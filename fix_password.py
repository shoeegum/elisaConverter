#!/usr/bin/env python3
"""
Fix the password hash in the environment variable.
This is a one-time script to fix the issue where the password is stored
instead of the hash in the APP_PASSWORD_HASH environment variable.
"""

import hashlib
import os

# The correct password is "IRelisa2017!"
password = "IRelisa2017!"
correct_hash = hashlib.sha1(password.encode()).hexdigest()

# Print current environment variable value
current_env = os.environ.get("APP_PASSWORD_HASH", "not set")
print(f"Current environment variable: {current_env}")
print(f"Correct hash should be: {correct_hash}")

# Instructions for fixing
print("\nTo fix this issue, please add a new Replit Secret:")
print("1. In Replit, click on 'Tools' in the left sidebar")
print("2. Select 'Secrets'")
print("3. Delete the existing APP_PASSWORD_HASH secret")
print("4. Click 'Add a new secret'")
print("5. For the key, enter: APP_PASSWORD_HASH")
print("6. For the value, enter the correct hash:")
print(f"   {correct_hash}")
print("7. Click 'Add Secret'")
print("8. Restart your application")
print("\nAfter doing this, you should be able to login with the password: IRelisa2017!")