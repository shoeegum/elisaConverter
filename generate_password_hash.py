#!/usr/bin/env python3
"""
Password Hash Generator for ELISA Parser Application

This script generates a SHA-1 hash for a password that can be used with the ELISA Parser application.
It provides instructions for setting up Replit Secrets for secure password storage.
"""

import hashlib
import getpass
import os
import sys

def generate_password_hash(password):
    """Generate a SHA-1 hash for the given password."""
    return hashlib.sha1(password.encode()).hexdigest()

def check_current_password():
    """Check if a default password is currently in use."""
    default_hash = "fe6a972039480fa98cafede1c8e048e0798b0f46"  # Hash for "IRelisa2017!"
    env_hash = os.environ.get("APP_PASSWORD_HASH")
    
    if env_hash:
        return (f"Current password is set from environment variable (hash: {env_hash[:8]}...{env_hash[-8:]})")
    else:
        return "Current password is the default: IRelisa2017!"

def show_replit_instructions(password_hash):
    """Show instructions for setting up a Replit Secret."""
    print("\nTo use your new password, follow these steps to create a Replit Secret:")
    print("1. In Replit, click on 'Tools' in the left sidebar")
    print("2. Select 'Secrets'")
    print("3. Click 'Add a new secret'")
    print("4. For the key, enter: APP_PASSWORD_HASH")
    print("5. For the value, enter your password hash:")
    print(f"   {password_hash}")
    print("6. Click 'Add Secret'")
    print("\nAfter adding the secret, restart your application.")
    print("You should now be able to log in with your new password.")

if __name__ == "__main__":
    print("ELISA Parser Password Hash Generator")
    print("====================================")
    print("This tool will generate a SHA-1 hash for your password")
    print("that can be used with the ELISA Parser application.")
    print()
    print(check_current_password())
    print()
    
    if len(sys.argv) > 1 and sys.argv[1] == "--check":
        # Only check current password
        exit(0)
    
    print("Enter a new password to generate a hash.")
    print("(Press Ctrl+C to cancel at any time)")
    print()
    
    try:
        # Get password (without echoing to screen)
        password = getpass.getpass("Enter new password: ")
        confirm_password = getpass.getpass("Confirm new password: ")
        
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
        
        show_replit_instructions(password_hash)
        
    except KeyboardInterrupt:
        print("\nPassword generation cancelled.")
        exit(0)