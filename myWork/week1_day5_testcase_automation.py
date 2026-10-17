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

    def get_text_from_file(self):
        """Reads the content of the .docx file if it is exist
      and extract the paragraphs as list of strings"""
        if not self.file_path or os.path.exists(self.file_path)
            return ""
        try:
            doc = Document(self.file_path)
            paragraphs_list_as_text = []
            for para in doc.paragraphs:
                if para.text.strip():
                    paragraphs_list_as_text.append(para.text)
            return  "\n".join( paragraphs_list_as_text)
        except Exception as e:
            print(f"An error occurred:{e}")
            return ""

