import os

from dotenv import load_dotenv

load_dotenv()

gmail_user = os.getenv('gmail_user')
gmail_password = os.getenv('gmail_password')