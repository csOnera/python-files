# on branch

# try if commit locally will change in github
import os
from dotenv import load_dotenv

load_dotenv()

POP_UN = os.getenv('POP_UN')
print(POP_UN)