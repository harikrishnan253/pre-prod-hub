import os
import re
import pythoncom
import win32com.client as win32

class ReferenceValidator:
    def __init__(self, filepath):
        self.filepath = os.path.abspath(filepath)
        self.word = None
        self.doc = None
        self.results = {
            'total_references': 0,
            'total_citations': 0,
            'missing_references': set(),
            'unused_references': set(),
            'sequence_issues': [],
            'citation_sequence': []
        }

    def __enter__(self):
        pythoncom.CoInitialize()
        self.word = win32.Dispatch("Word.Application")
        self.word.Visible = False
        self.word.ScreenUpdating = False
        # Open read-only by default for validation. Renumbering will reopen writable when needed.
        self.doc = self.word.Documents.Open(self.filepath, ReadOnly=True)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc:
            self.doc.Close(SaveChanges=False)
        if self.word:
            self.word.Quit()
        pythoncom.CoUninitialize()

    def validate(self, auto_renumber=False, save_path=None):
        ref_numbers = self._get_reference_numbers()
        self.results['total_references'] = len(ref_numbers)

        citations = self._get_citations()
        self.results['total_citations'] = len(citations)

        cited_numbers = set()
        for citation in citations:
            cited_numbers.update(citation['numbers'])
            self.results['citation_sequence'].extend(citation['numbers'])

        self.results['missing_references'] = sorted(cited_numbers - ref_numbers)
        self.results['unused_references'] = sorted(ref_numbers - cited_numbers)
        self._check_citation_sequence()
        # Optionally auto-renumber when sequence issues are found
        if auto_renumber and 'NOT in sequence' in self.results.get('sequence_message', ''):
            ren = self.renumber_if_needed(save_path=save_path)
            self.results['renumber_attempt'] = ren
            # If renumbering happened, refresh results by reopening the saved/original document
            if ren.get('renumbered'):
                reopen_path = save_path if save_path else self.filepath
                try:
                    if self.doc:
                        self.doc.Close(SaveChanges=False)
                except Exception:
                    pass
                try:
                    self.doc = self.word.Documents.Open(os.path.abspath(reopen_path), ReadOnly=True)
                except Exception:
                    try:
                        self.doc = self.word.Documents.Open(reopen_path)
                    except Exception:
                        # couldn't reopen, return current results with renumber info
                        return self.results

                # recompute results
                ref_numbers = self._get_reference_numbers()
                self.results['total_references'] = len(ref_numbers)
                citations = self._get_citations()
                self.results['total_citations'] = len(citations)
                cited_numbers = set()
                self.results['citation_sequence'] = []
                for citation in citations:
                    cited_numbers.update(citation['numbers'])
                    self.results['citation_sequence'].extend(citation['numbers'])
                self.results['missing_references'] = sorted(cited_numbers - ref_numbers)
                self.results['unused_references'] = sorted(ref_numbers - cited_numbers)
                self._check_citation_sequence()

        return self.results

    def renumber_if_needed(self, save_path=None):
        """
        If citation sequence is not ordered, renumber citations and reference list.
        Writes changes back to the document (overwrites original unless save_path provided).
        """
        # Ensure we have the latest sequence data
        sequence = self.results.get('citation_sequence', [])
        if not sequence:
            return {'renumbered': False, 'message': 'No citations found.'}

        seen = set()
        unique_sequence = []
        for num in sequence:
            if num not in seen:
                unique_sequence.append(num)
                seen.add(num)

        is_ordered = unique_sequence == sorted(unique_sequence)
        if is_ordered:
            return {'renumbered': False, 'message': 'Citations already in sequence.'}

        # Build renumber map: first-appearance order becomes 1..n. Include any reference-only numbers afterwards.
        ref_numbers = sorted(list(self._get_reference_numbers()))
        ordered_all = list(unique_sequence)
        
        # Add any references that weren't cited (append them at the end in their original relative order)
        for n in ref_numbers:
            if n not in seen:
                ordered_all.append(n)

        renumber_map = {old: new for new, old in enumerate(ordered_all, start=1)}
        new_to_old = {v: k for k, v in renumber_map.items()}

        # Reopen document writable
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
            self.doc = self.word.Documents.Open(self.filepath, ReadOnly=False)
        except Exception:
            # try opening without ReadOnly named arg
            self.doc = self.word.Documents.Open(self.filepath)

        # ---------------------------------------------------------
        # 1. Update Citations in Text (cite_bib)
        # ---------------------------------------------------------
        try:
            cite_style = self.doc.Styles("cite_bib")
        except Exception:
            cite_style = None

        if cite_style:
            rng = self.doc.Content
            rng.Find.ClearFormatting()
            rng.Find.Style = cite_style
            rng.Find.Text = ""
            while rng.Find.Execute():
                text = rng.Text.strip()
                if not text:
                    rng.Collapse(0)
                    continue
                nums = self._extract_numbers(text)
                if not nums:
                    rng.Collapse(0)
                    continue

                # Map numbers
                mapped = [renumber_map.get(n, n) for n in nums]
                new_text = self._numbers_to_string(mapped)

                # Replace only the numeric portion
                new_display = re.sub(r'\d+(-\d+)?', lambda m: self._mapped_segment(m.group(0), renumber_map), text)
                if not re.search(r'\d', new_display):
                    new_display = new_text

                try:
                    rng.Text = new_display
                except Exception:
                    pass
                rng.Collapse(0)

        # ---------------------------------------------------------
        # 2. Sort and Renumber Reference List (REF-N)
        # ---------------------------------------------------------
        # We need to physically move paragraphs to match the new order (1, 2, 3...)
        # Strategy: Copy all REF-N paragraphs to a temp doc, then paste them back in the correct order.
        
        try:
            refpara_style = self.doc.Styles("REF-N")
        except Exception:
            refpara_style = None

        if refpara_style:
            # A. Collect all REF-N paragraphs from source
            # We assume the current order in doc corresponds to ref_numbers sorted (1, 2, 3...)
            # We need to capture them and map them to their OLD number.
            
            source_paras = [] # List of Range objects (pointers might be unstable if we edit, so we copy to temp doc immediately)
            
            # Create Temp Doc
            temp_doc = self.word.Documents.Add(Visible=False)
            
            # Iterate and copy to temp doc
            # We need to know which old number corresponds to which paragraph.
            # Assumption: The first REF-N paragraph found is Ref #1, second is #2, etc.
            # This relies on the input doc being ordered 1..N initially (even if citations are out of order).
            # If the input doc is ALREADY out of order (e.g. Ref list is 1, 5, 2...), this assumption fails.
            # BUT, `_get_reference_numbers` just scrapes numbers.
            # Let's assume standard scientific writing: References are listed 1..N.
            
            original_ref_map = {} # old_number -> temp_doc_range
            
            current_ref_idx = 0
            for para in self.doc.Paragraphs:
                try:
                    if para.Range.Style == refpara_style or getattr(para.Range.Style, 'NameLocal', '') == 'REF-N':
                        # This is a reference paragraph.
                        # Which number is it?
                        # We can try to extract the number from it to be sure.
                        txt = para.Range.Text
                        extracted = self._extract_numbers(txt)
                        if extracted:
                            # Use the first number found as the identifier
                            old_num = extracted[0]
                        else:
                            # Fallback to index if extraction fails (risky)
                            if current_ref_idx < len(ref_numbers):
                                old_num = ref_numbers[current_ref_idx]
                            else:
                                old_num = -1 # Unknown
                        
                        current_ref_idx += 1
                        
                        # Copy to temp doc
                        para.Range.Copy()
                        # PasteAppend
                        rng_end = temp_doc.Content
                        rng_end.Collapse(0) # End
                        rng_end.Paste()
                        
                        # Store the range in temp doc that corresponds to this old_num
                        # The pasted paragraph is the last one in temp_doc
                        original_ref_map[old_num] = temp_doc.Paragraphs.Last.Range
                except Exception:
                    continue
            
            # B. Renumber Citations and References (No Physical Sorting)
            # Strategy: Build a mapping of old numbers to new numbers based on citation order,
            # then update the numbers in both citations and references
            
            # Step 1: Build citation order to create renumber map
            citation_order = []  # Track unique citation numbers in order of appearance
            try:
                cite_style = self.doc.Styles("cite_bib")
            except Exception:
                cite_style = None
            
            if cite_style:
                # Find all citations and track their order
                rng = self.doc.Content
                rng.Find.ClearFormatting()
                rng.Find.Style = cite_style
                rng.Find.Text = ""
                
                while rng.Find.Execute():
                    text = rng.Text.strip()
                    nums = self._extract_numbers(text)
                    
                    # Add each number to citation_order if not already there
                    for num in nums:
                        if num not in citation_order:
                            citation_order.append(num)
                    
                    rng.Collapse(0)
            
            # Build renumber map from citation order
            if citation_order:
                renumber_map = {old: new for new, old in enumerate(citation_order, start=1)}
            else:
                renumber_map = renumber_map  # Use the existing one from earlier
            
            # Step 2: Update citation numbers in text
            if cite_style:
                rng = self.doc.Content
                rng.Find.ClearFormatting()
                rng.Find.Style = cite_style
                rng.Find.Text = ""
                
                while rng.Find.Execute():
                    text = rng.Text.strip()
                    nums = self._extract_numbers(text)
                    
                    if nums:
                        # Map numbers
                        mapped = [renumber_map.get(n, n) for n in nums]
                        new_text = self._numbers_to_string(mapped)
                        
                        # Replace the numeric portion
                        new_display = re.sub(r'\d+(-\d+)?', lambda m: self._mapped_segment(m.group(0), renumber_map), text)
                        if not re.search(r'\d', new_display):
                            new_display = new_text
                        
                        try:
                            rng.Text = new_display
                        except Exception:
                            pass
                    
                    rng.Collapse(0)
            
            # Step 3: Update reference list numbers (bib_number style)
            try:
                bib_style = self.doc.Styles("bib_number")
            except Exception:
                bib_style = None
            
            if bib_style:
                rng = self.doc.Content
                rng.Find.ClearFormatting()
                rng.Find.Style = bib_style
                rng.Find.Text = ""
                
                while rng.Find.Execute():
                    txt = rng.Text.strip()
                    if txt:
                        # Extract the number and replace with mapped value
                        match = re.search(r'\d+', txt)
                        if match:
                            old_num = int(match.group())
                            new_num = renumber_map.get(old_num, old_num)
                            new_txt = re.sub(r'\d+', str(new_num), txt)
                            try:
                                rng.Text = new_txt
                            except Exception:
                                pass
                    
                    rng.Collapse(0)

            # C. Update Numbers in the Sorted List
            # Now that the list is sorted, we need to update the visible numbers (e.g. change "2." to "1.")
            # We iterate again.
            
            current_idx = 1
            for para in self.doc.Paragraphs:
                 if para.Range.Style == refpara_style or getattr(para.Range.Style, 'NameLocal', '') == 'REF-N':
                     # Find bib_number in this paragraph
                     p_rng = para.Range
                     try:
                         bib_style = self.doc.Styles("bib_number")
                         p_rng.Find.ClearFormatting()
                         p_rng.Find.Style = bib_style
                         p_rng.Find.Text = ""
                         if p_rng.Find.Execute():
                             # Found the number. Replace it.
                             old_txt = p_rng.Text
                             # Regex replace the number part
                             new_txt = re.sub(r'\d+', str(current_idx), old_txt)
                             p_rng.Text = new_txt
                     except:
                         pass
                     
                     current_idx += 1

            # Close temp doc
            temp_doc.Close(SaveChanges=False)

        # Save document
        try:
            if save_path:
                self.doc.SaveAs(FileName=os.path.abspath(save_path))
            else:
                # Overwrite existing document
                self.doc.Save()
        except Exception:
            # Some Word versions use SaveAs2
            try:
                if save_path:
                    self.doc.SaveAs2(FileName=os.path.abspath(save_path))
                else:
                    self.doc.Save()
            except Exception as e:
                return {'renumbered': False, 'message': f'Failed to save document: {e}'}

        return {'renumbered': True, 'map': renumber_map}

    def _mapped_segment(self, seg, renumber_map):
        """Map a segment like '2' or '2-4' using renumber_map and return a string representation."""
        if '-' in seg:
            a, b = seg.split('-', 1)
            a_n = renumber_map.get(int(a), int(a))
            b_n = renumber_map.get(int(b), int(b))
            # If mapped numbers are contiguous, show as range
            if b_n - a_n >= 1 and self._is_contiguous_mapping(int(a), int(b), renumber_map):
                return f"{a_n}-{b_n}"
            else:
                # Return comma-separated mapped numbers
                return ','.join(str(renumber_map.get(int(x), int(x))) for x in range(int(a), int(b)+1))
        else:
            n = int(seg)
            return str(renumber_map.get(n, n))

    def _is_contiguous_mapping(self, a, b, renumber_map):
        vals = [renumber_map.get(i, i) for i in range(a, b+1)]
        return vals == list(range(min(vals), max(vals)+1))

    def _numbers_to_string(self, numbers):
        """Convert list of integers into a compact string like '1,3-5,7'."""
        if not numbers:
            return ''
        nums = sorted(numbers)
        ranges = []
        start = prev = nums[0]
        for n in nums[1:]:
            if n == prev + 1:
                prev = n
                continue
            else:
                if start == prev:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{prev}")
                start = prev = n
        # finalize
        if start == prev:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{prev}")
        return ','.join(ranges)

    def _get_reference_numbers(self):
        numbers = set()

        # Strategy: Only extract the reference LIST number (e.g., "1", "2", "3"), 
        # NOT all numbers in the reference text (years, pages, volumes, etc.)
        
        # First try to find reference paragraphs styled as 'REF-N' (paragraph style for the whole ref)
        try:
            refpara_style = self.doc.Styles("REF-N")
        except Exception:
            refpara_style = None

        # Also get bib_number style for extracting just the number
        try:
            bib_char_style = self.doc.Styles("bib_number")
        except Exception:
            bib_char_style = None

        if refpara_style:
            for para in self.doc.Paragraphs:
                try:
                    if para.Range.Style == refpara_style or getattr(para.Range.Style, 'NameLocal', '') == 'REF-N':
                        # Look for bib_number style within this paragraph
                        found_number = False
                        if bib_char_style:
                            # Search for bib_number style in this paragraph
                            p_rng = para.Range
                            p_rng.Find.ClearFormatting()
                            p_rng.Find.Style = bib_char_style
                            p_rng.Find.Text = ""
                            if p_rng.Find.Execute():
                                # Extract only the first number from the bib_number styled text
                                bib_text = p_rng.Text.strip()
                                # Extract just the first number (ignore ranges, just get the leading number)
                                match = re.search(r'\d+', bib_text)
                                if match:
                                    numbers.add(int(match.group()))
                                    found_number = True
                        
                        # Fallback: if no bib_number found, extract first number from paragraph
                        if not found_number:
                            text = para.Range.Text.strip()
                            match = re.search(r'\d+', text)
                            if match:
                                numbers.add(int(match.group()))
                except Exception:
                    continue

        # Also check for standalone bib_number occurrences (if not already found via REF-N)
        # This handles cases where bib_number might be used without REF-N paragraph style
        if bib_char_style and not refpara_style:
            for item in self._find_style_ranges("bib_number"):
                bib_text = item['text'].strip()
                match = re.search(r'\d+', bib_text)
                if match:
                    numbers.add(int(match.group()))

        return numbers

    def _get_citations(self):
        # Find citations that use 'cite_bib' character style anywhere in the text
        try:
            _ = self.doc.Styles("cite_bib")
        except Exception:
            raise ValueError("'cite_bib' style not found")

        citations = []
        for item in self._find_style_ranges("cite_bib"):
            text = item['text'].strip()
            if text:
                numbers = self._extract_numbers(text)
                if numbers:
                    citations.append({
                        'text': text,
                        'numbers': numbers,
                        'range_start': item.get('range_start'),
                        'range_end': item.get('range_end')
                    })

        return citations

    def _find_style_ranges(self, style_name):
        """Return list of dicts {'text','range_start','range_end'} for occurrences of a style.
        Supports both paragraph and character styles by running a paragraph scan and a Find on the document content.
        """
        results = []
        # Try to get the style object (may be paragraph or character style)
        try:
            style_obj = self.doc.Styles(style_name)
        except Exception:
            style_obj = None

        # 1) Paragraph-level scan: find paragraphs whose paragraph style matches style_name
        if style_obj:
            try:
                for para in self.doc.Paragraphs:
                    try:
                        if para.Range.Style == style_obj or getattr(para.Range.Style, 'NameLocal', '') == style_name:
                            results.append({'text': para.Range.Text, 'range_start': para.Range.Start, 'range_end': para.Range.End})
                    except Exception:
                        continue
            except Exception:
                pass

        # 2) Character-style Find across the document content (this will also find character style runs)
        try:
            rng = self.doc.Content
            rng.Find.ClearFormatting()
            if style_obj:
                rng.Find.Style = style_obj
            rng.Find.Text = ""
            rng.Find.Format = True
            # Execute Find: when Style is a character style this finds character runs; for paragraph styles it may re-find paragraphs too
            while rng.Find.Execute():
                results.append({'text': rng.Text, 'range_start': rng.Start, 'range_end': rng.End})
                rng.Collapse(0)
        except Exception:
            pass

        return results

    def _extract_numbers(self, text):
        numbers = []

        # Handle ranges first
        for match in re.finditer(r'(\d+)-(\d+)', text):
            start, end = int(match.group(1)), int(match.group(2))
            
            # Safety check for massive ranges (e.g. typos like 1-1000000 or phone numbers)
            if end - start > 999:
                # If range is too large, treat as separate numbers to avoid memory issues
                numbers.append(start)
                numbers.append(end)
            elif end >= start:
                numbers.extend(range(start, end + 1))

        # Get individual numbers (excluding those that were part of ranges)
        text_no_ranges = re.sub(r'\d+-\d+', '', text)
        numbers.extend([int(match.group()) for match in re.finditer(r'\b\d+\b', text_no_ranges)])

        return numbers

    def _check_citation_sequence(self):
        sequence = self.results['citation_sequence']
        if len(sequence) < 2:
            self.results['sequence_message'] = "Citations are in proper sequence."
            return

        seen = set()
        unique_sequence = []
        for num in sequence:
            if num not in seen:
                unique_sequence.append(num)
                seen.add(num)

        is_ordered = unique_sequence == sorted(unique_sequence)

        if not is_ordered:
            self.results['sequence_issues'].append(sequence)
            self.results['sequence_message'] = "Citations are NOT in sequence."
        else:
            self.results['sequence_message'] = "Citations are in proper sequence."
