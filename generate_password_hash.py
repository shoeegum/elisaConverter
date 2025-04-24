#!/usr/bin/env python3
"""
Password Hash Generator for ELISA Parser Application

This script generates a SHA-1 hash for a password that can be used with the ELISA Parser application.
"""

import hashlib
import getpass

def generate_password_hash(password):
    """Generate a SHA-1 hash for the given password."""
    return hashlib.sha1(password.encode()).hexdigest()

if __name__ == "__main__":
    print("ELISA Parser Password Hash Generator")
    print("====================================")
    print("This tool will generate a SHA-1 hash for your password")
    print("that can be used with the ELISA Parser application.")
    print()
    
    # Get password (without echoing to screen)
    password = getpass.getpass("Enter your password: ")
    confirm_password = getpass.getpass("Confirm your password: ")
    
    if password != confirm_password:
        print("Passwords do not match. Please try again.")
        exit(1)
    
    if not password:
        print("Password cannot be empty. Please try again.")
        exit(1)
    
    # Generate hash
    password_hash = generate_password_hash(password)
    
    print("\nYour password hash (SHA-1):")
    print(password_hash)
    print("\nTo use this password, add it to your Replit Secrets with the key 'APP_PASSWORD_HASH'")