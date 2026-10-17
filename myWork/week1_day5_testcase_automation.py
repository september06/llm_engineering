import os
import json 
from dotenv import load_dotenv 
from openai import OpenAI
import sys
import re
import gradio as gr
#need to pip install pandas 
try:
    import pandas as pd
    print("✅ pandas imported successfully")
except ImportError:
    print("❌ Error: 'pandas' module not found.")
    print("💡 Install with: pip install pandas")
    sys.exit(1)

#-docx
try:
    from docx import Document
    print("✅ python-docx imported successfully")
except ImportError:
    print("❌ Error: 'python-docx' module not found.")
    print("💡 Install with: pip install python-docx")
    sys.exit(1)
load_dotenv(override=True)
GOOGLE_KEY = os.getenv('GEMINI_API_KEY')
MODEL = "gemini-3.8-flash"

class Get_document:
    def __init__(self, file_path):
        self.file_path = file_path

    def
