import re

def extract_numbers(text):
    numbers = []
    # Handle ranges first
    for match in re.finditer(r'(\d+)-(\d+)', text):
        start, end = int(match.group(1)), int(match.group(2))
        print(f"Found range: {start}-{end}")
        # numbers.extend(range(start, end + 1)) # Commented out to avoid crash
        if end - start > 1000:
            print(f"WARNING: Massive range detected! Size: {end - start}")
    return numbers

test_cases = [
    "See figures 1-5",
    "Year 2024-2025",
    "Typo 1-1000000",
    "Phone 555-1234"
]

for t in test_cases:
    print(f"Testing: '{t}'")
    extract_numbers(t)
