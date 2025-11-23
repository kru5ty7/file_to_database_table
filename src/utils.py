"""
Utility functions for file to database converter
"""

import os
import re
import logging
from datetime import datetime
from cryptography.fernet import Fernet
import base64

# Configure logging
def setup_logging():
    """Setup logging configuration with both file and console handlers"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_filename = os.path.join(log_dir, f"file_to_db_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    # Create logger
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)

    # Clear any existing handlers
    logger.handlers = []

    # File handler - detailed logging
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s')
    file_handler.setFormatter(file_formatter)

    # Console handler - less verbose
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('%(levelname)s - %(message)s')
    console_handler.setFormatter(console_formatter)

    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logger.info(f"Logging initialized. Log file: {log_filename}")
    return logger

logger = setup_logging()

# Encryption key management
def get_or_create_key():
    """Get or create encryption key for password storage"""
    key_file = '.encryption_key'
    if os.path.exists(key_file):
        with open(key_file, 'rb') as f:
            key = f.read()
    else:
        key = Fernet.generate_key()
        with open(key_file, 'wb') as f:
            f.write(key)
        logger.info("Generated new encryption key")
    return key

def encrypt_password(password):
    """Encrypt password using Fernet symmetric encryption"""
    if not password:
        return ""
    key = get_or_create_key()
    f = Fernet(key)
    encrypted = f.encrypt(password.encode())
    return base64.urlsafe_b64encode(encrypted).decode()

def decrypt_password(encrypted_password):
    """Decrypt password using Fernet symmetric encryption"""
    if not encrypted_password:
        return ""
    try:
        key = get_or_create_key()
        f = Fernet(key)
        decoded = base64.urlsafe_b64decode(encrypted_password.encode())
        decrypted = f.decrypt(decoded)
        return decrypted.decode()
    except Exception as e:
        logger.error(f"Failed to decrypt password: {e}")
        # If decryption fails, assume it's plain text (backward compatibility)
        return encrypted_password

def sanitize_name(name):
    """
    Sanitize table and column names:
    - Convert to lowercase
    - Remove special characters (keep only alphanumeric and underscores)
    - Replace whitespace with underscores
    - Remove leading/trailing underscores
    - Ensure it doesn't start with a number
    """
    original_name = name
    # Convert to lowercase
    name = name.lower()
    # Replace whitespace with underscores
    name = re.sub(r'\s+', '_', name)
    # Remove special characters, keep only alphanumeric and underscores
    name = re.sub(r'[^a-z0-9_]', '', name)
    # Remove leading/trailing underscores
    name = name.strip('_')
    # Ensure it doesn't start with a number
    if name and name[0].isdigit():
        name = 'col_' + name
    # Ensure it's not empty
    if not name:
        name = 'unnamed'

    if original_name != name:
        logger.debug(f"Sanitized name: '{original_name}' -> '{name}'")

    return name
