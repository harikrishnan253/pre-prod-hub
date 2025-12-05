import urllib.request
import urllib.parse
from html.parser import HTMLParser
import traceback

class CSRFParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.csrf_token = None
    
    def handle_starttag(self, tag, attrs):
        if tag == 'input':
            attrs_dict = dict(attrs)
            if attrs_dict.get('name') == 'csrf_token':
                self.csrf_token = attrs_dict.get('value')

try:
    # Get the login page to extract CSRF token
    response = urllib.request.urlopen('http://localhost:5001/login')
    html_content = response.read().decode('utf-8')
    
    parser = CSRFParser()
    parser.feed(html_content)
    csrf_token = parser.csrf_token

    print(f"CSRF Token: {csrf_token}")

    if csrf_token:
        # Now try to login with the token
        login_data = urllib.parse.urlencode({
            'username': 'admin',
            'password': 'admin123',
            'csrf_token': csrf_token
        }).encode('utf-8')
        
        try:
            login_response = urllib.request.urlopen('http://localhost:5001/login', login_data)
            print(f"Status: {login_response.status}")
            print("Login successful!")
        except urllib.error.HTTPError as e:
            print(f"Status: {e.code}")
            print(f"Response: {e.read().decode('utf-8')[:500]}")
    else:
        print("Could not find CSRF token")
except Exception as e:
    print(f"Error: {e}")
    traceback.print_exc()
