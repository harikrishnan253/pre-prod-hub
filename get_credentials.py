import sqlite3
from config import DATABASE

# Connect to the database
conn = sqlite3.connect(DATABASE)
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Query all users
cursor.execute("SELECT username, password, email, is_admin FROM users")
users = cursor.fetchall()

print("=" * 80)
print("LOGIN CREDENTIALS FROM DATABASE")
print("=" * 80)
print()

if users:
    for user in users:
        print(f"Username: {user['username']}")
        print(f"Email: {user['email']}")
        print(f"Admin: {user['is_admin']}")
        print(f"Password Hash: {user['password']}")
        print("-" * 80)
else:
    print("No users found in database")

conn.close()
