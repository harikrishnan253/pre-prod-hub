import re
import sys
import os

# Mocking the validator method to test logic without full environment
class MockValidator:
    def _extract_numbers(self, text):
        numbers = []

        # Handle ranges first
        for match in re.finditer(r'(\d+)-(\d+)', text):
            start, end = int(match.group(1)), int(match.group(2))
            
            # Safety check for massive ranges (e.g. typos like 1-1000000 or phone numbers)
            if end - start > 999:
                # If range is too large, treat as separate numbers to avoid memory issues
                print(f"  -> Range too large ({start}-{end}), treating as separate numbers")
                numbers.append(start)
                numbers.append(end)
            elif end >= start:
                numbers.extend(range(start, end + 1))

        # Get individual numbers (excluding those that were part of ranges)
        text_no_ranges = re.sub(r'\d+-\d+', '', text)
        numbers.extend([int(match.group()) for match in re.finditer(r'\b\d+\b', text_no_ranges)])

        return numbers

validator = MockValidator()

test_cases = [
    "See figures 1-5",
    "Year 2024-2025",
    "Typo 1-1000000",
    "Phone 555-1234"
]

print("Verifying fix logic:")
for t in test_cases:
    print(f"Testing: '{t}'")
    nums = validator._extract_numbers(t)
    print(f"  Result count: {len(nums)}")
    if len(nums) > 1000:
        print("  FAIL: Result too large!")
    else:
        print("  PASS")
