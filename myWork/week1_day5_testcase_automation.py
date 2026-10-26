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

GOOGLE_KEY = os.getenv('GEMINI_API_KEY')
MODEL = "gemini-3.8-flash"

class Get_document:
    def __init__(self, file_path):
        self.file_path = file_path

    def get_text_from_file(self):
        """Reads the content of the .docx file if it is exist
      and extract the paragraphs as list of strings"""
        if not self.file_path or os.path.exists(self.file_path):
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

load_dotenv(override=True)
api_key = GOOGLE_KEY
openai = OpenAI(
    api_key = api_key,
    base_url  = "https://generativelanguage.googleapis.com/v1beta/openai/"
)


req_doc_system_prompt = "You are provided with a complete requirements specifications document. \
You are able to decide which content from that document are related to actual requirements, identify each requirement as \
functional or non-functional and list them all.\n"
req_doc_system_prompt += "If the document is empty or do not contain requirements or if you cannot extract them, please respond as such.\
Do not make up your own requirements. \n"
req_doc_system_prompt += "You should respond in JSON as in this example:"
req_doc_system_prompt += """
{
    "requirements": [
        {"RequirementNo": "FR-01", "Requirement Description": "description of this functional requirement goes here"},
        {"RequirementNo": "FR-02": "Requirement Description": "description of this functional requirement goes here"},
        {"RequirementNo": "NFR-01": "Requirement Description": "description of this non-functional requirement goes here"},
        {"RequirementNo": "NFR-02": "Requirement Description": "description of this non-functional requirement goes here"}
    ]
}
"""
def req_doc_user_prompt(doc):
    user_prompt = "Here is the contents from a requirement document.\n"
    user_prompt += f"{doc.extract()} \n"
    user_prompt += "Please scan through the document and extract only the  actual requirements. For example, ignore sections or \
paragraphs such as Approvers, table of contents and similar sections which are not really requirements.\
You must respond in a JSON format"
    user_prompt += "If the content is empty, respond that there are no valid requirements you could extract and ask for a proper document.\n"
    user_prompt = user_prompt[:25_000] # Truncate if more than 25,000 characters
    return user_prompt

def process_requirements_and_generate_tests(file_path):
    if not file_path:
        yield "please check file path ", None
        return
    reqdoc = Get_document(file_path)
    text = reqdoc.extract()
    if not text.strip():
        yield "Empty file !! ", None
        return
    try:

        response = openai.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": req_doc_system_prompt},
                {"role": "user", "content": req_doc_user_prompt(reqdoc)}
            ],
        response_format={"type": "json_object"}
        )
        requirements_data = response.choices[0].message.content
    except Exception as e:
        yield f" Something Wrong {str(e)}", None
        return
    return json.loads(requirements_data)

def get_test_system_prompt():
    test_system_prompt = """You are an expert QA engineer. that receives a list of functional and non functional requirements in JSON format. You are the expert in generating unit test cases for each requirement. \
    You will create as many different test cases as needed for each requirement and produce a result in a table. Order the table by requirement No. Provide clear details on test case pass criteria. \
    Generate detailed unit test cases for the provided requirements list.
    Output the result ONLY as a Markdown Table with these exact columns:
    | S No | Requirement No | Requirement Description | Test Case ID | Test Case Summary | Test Case Description | Success Criteria |
    """
    test_system_prompt += "If you are provided with an empty list, ask for a proper requirement doc\n"
    return test_system_prompt

def get_test_user_prompt(file_path):
        user_prompt = "You are looking at the following list of requirements. \n"
        user_prompt += f"{process_requirements_and_generate_tests(file_path)}\n"
        user_prompt += "Prepare unit test cases for each of these requirements in a table and send that table as response. \n"
        user_prompt += user_prompt[:25000]
        return user_prompt


def create_testcase_gradio(response, is_response_ready, is_cleared, file_input):
    if is_cleared or file_input == None:  # Prevent OpenAI call if "Clear" was clicked
        return "", False
    stream = openai.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": get_test_system_prompt},
            {"role": "user", "content":get_test_user_prompt(file_input)}
          ],
        stream=True
    )
    #Modified for Gradio
    result = ""
    for chunk in stream:
        result += chunk.choices[0].delta.content or ""
        #print(result)
        yield result, False
        excel_path = export_markdown_table_to_excel(result)
        yield result, excel_path

def export_markdown_table_to_excel(markdown_text):
    lines = [line.strip() for line in markdown_text.split('\n') if line.strip().startswith('|')]
    if len(lines) < 3: # חייב להיות לפחות כותרת, מפריד, ושורה אחת
        return None
    
    # חילוץ כותרות ונתונים
    headers = [col.strip() for col in lines[0].split('|')[1:-1]]
    data = []
    for line in lines[2:]: # דילוג על שורת המפריד (|---|---|)
        row = [col.strip() for col in line.split('|')[1:-1]]
        if len(row) == len(headers):
            data.append(row)

    if not data:
        return None

    df = pd.DataFrame(data, columns=headers)
    output_file = "test_cases.xlsx"
    df.to_excel(output_file, index=False)
    return output_file

force_dark_mode = """
function refresh() {
    const url = new URL(window.location);
    if (url.searchParams.get('__theme') !== 'dark') {
        url.searchParams.set('__theme', 'dark');
        window.location.href = url.href;
    }
}
"""

with gr.Blocks(theme=gr.themes.Soft()) as demo:
    gr.Markdown("# 📄 אוטומציה ליצירת תרחישי בדיקה (Test Cases)")
    
    with gr.Row():
        file_input = gr.File(label="העלה מסמך דרישות (.docx)", file_types=[".docx"])
    
    submit_btn = gr.Button("צור בדיקות", variant="primary")
    
    with gr.Row():
        markdown_output = gr.Markdown(label="תוצאות בדיקה")
    
    with gr.Row():
        excel_output = gr.File(label="הורד קובץ Excel", visible=True)

    # הפעלה בלחיצה
    submit_btn.click(
        fn=process_requirements_and_generate_tests,
        inputs=[file_input],
        outputs=[markdown_output, excel_output]
    )

demo.launch()