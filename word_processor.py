import os
import time
import subprocess
import pythoncom
import win32com.client as win32
import pywintypes
from config import COMMON_MACRO_FOLDER, DEFAULT_MACRO_NAME, WORD_START_RETRIES, ROUTE_MACROS
from utils import log_errors

class OptimizedDocumentProcessor:
    def __init__(self):
        self.word = None
        self.docs = []
        self.macro_template_loaded = False

    def __enter__(self):
        pythoncom.CoInitialize()
        self.word = self._start_word_optimized()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._cleanup()

    def _start_word_optimized(self):
        for attempt in range(WORD_START_RETRIES):
            try:
                subprocess.run(["taskkill", "/f", "/im", "winword.exe"],
                               capture_output=True, check=False)

                word = win32.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False
                word.AutomationSecurity = 1
                word.ScreenUpdating = False
                word.Options.DoNotPromptForConvert = True
                word.Options.ConfirmConversions = False
                return word
            except Exception as e:
                if attempt == WORD_START_RETRIES - 1:
                    raise RuntimeError(f"Failed to start Word: {e}")
                time.sleep(1)

    def _load_macro_template(self):
        if self.macro_template_loaded:
            return True

        try:
            macro_path = os.path.join(COMMON_MACRO_FOLDER, DEFAULT_MACRO_NAME)
            if not os.path.exists(macro_path):
                return False

            for addin in self.word.AddIns:
                try:
                    if hasattr(addin, 'FullName') and addin.FullName.lower().endswith(DEFAULT_MACRO_NAME.lower()):
                        self.macro_template_loaded = True
                        return True
                except:
                    continue

            self.word.AddIns.Add(macro_path, True)
            self.macro_template_loaded = True
            return True

        except Exception as e:
            log_errors([f"Failed to load macro template: {str(e)}"])
            return False

    def process_documents_batch(self, file_paths, selected_tasks, route_type):
        errors = []

        if not self._load_macro_template():
            errors.append("Failed to load macro template")
            return errors

        route_macros = ROUTE_MACROS.get(route_type, {}).get('macros', [])

        for doc_path in file_paths:
            try:
                abs_path = os.path.abspath(doc_path)
                if not os.path.exists(abs_path):
                    errors.append(f"File not found: {abs_path}")
                    continue

                doc = self.word.Documents.Open(abs_path, ReadOnly=False, AddToRecentFiles=False)
                self.docs.append(doc)

                for task_index in selected_tasks:
                    try:
                        idx = int(task_index)
                        if 0 <= idx < len(route_macros):
                            macro_name = route_macros[idx]
                            try:
                                self.word.Run(macro_name)
                            except pywintypes.com_error as ce:
                                errors.append(f"COM error running '{macro_name}': {ce}")
                            except Exception as me:
                                errors.append(f"Macro '{macro_name}' failed: {me}")
                        else:
                            errors.append(f"Invalid task index {idx} for route {route_type}")
                    except ValueError:
                        errors.append(f"Invalid task index: {task_index}")

                try:
                    doc.Save()
                    doc.Close(SaveChanges=False)
                    self.docs.remove(doc)
                except Exception as se:
                    errors.append(f"Failed to save document: {se}")

            except Exception as doc_err:
                errors.append(f"Document processing failed: {doc_err}")

        return errors

    def _cleanup(self):
        for doc in self.docs:
            try:
                doc.Close(SaveChanges=False)
            except:
                pass

        if self.word:
            try:
                self.word.Quit()
            except:
                pass

        try:
            pythoncom.CoUninitialize()
        except:
            pass
