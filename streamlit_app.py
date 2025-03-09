import streamlit as st
import pandas as pd
from langchain_community.embeddings import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_community.chat_models import ChatOpenAI
from langchain.prompts import PromptTemplate
from langchain.chains import RetrievalQA
from langchain.text_splitter import RecursiveCharacterTextSplitter
import docx
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import smtplib
from email.message import EmailMessage

# Verify environment variable
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    st.error("OpenAI API key not found! Set it in your Streamlit secrets or your .env file.")
    st.stop()

# Add Synchrony logo aligned left and smaller
st.image("syf logo.png", width=50)

# File paths
sas_code_path = "sas_code_example.docx"
requirements_path = "campaign_requirements_example.docx"
excel_path = "project_segment_details.xlsx"

# Verify all required files exist
missing_files = [fp for fp in [sas_code_path, requirements_path, excel_path] if not os.path.exists(fp)]
if missing_files:
    st.error(f"Required file(s) not found: {', '.join(missing_files)}")
    st.stop()

# Load Excel data with sheet verification
data = pd.ExcelFile(excel_path)
st.write("Available Excel Sheets:", data.sheet_names)

if "Project Details" not in data.sheet_names or "Segment Details" not in data.sheet_names:
    st.error("Required Excel sheets ('Project Details' and 'Segment Details') are missing or incorrectly named.")
    st.stop()

project_details = data.parse("Project Details")
segment_details = data.parse("Segment Details")

# Load SAS code and campaign requirements
def load_docx(file_path):
    doc = docx.Document(file_path)
    return '\n'.join(para.text for para in doc.paragraphs)

sas_text = load_docx(sas_code_path)
campaign_text = load_docx(requirements_path)

# Prepare vector DB
text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
sas_chunks = text_splitter.split_text(sas_text)
campaign_chunks = text_splitter.split_text(campaign_text)

embeddings = OpenAIEmbeddings(openai_api_key=openai_api_key)
vector_db = FAISS.from_texts(sas_chunks + campaign_chunks, embeddings)

# Streamlit UI enhancements
st.title("✨ AI-Based Campaign Operation Programming ✨")

wf_number = st.text_input("🔢 Enter Workfront Number (Numeric Only)")
user_email = st.text_input("📧 Enter Your Email")

if st.button("🚀 Submit"):
    if not wf_number.isdigit():
        st.error("Workfront number must be numeric.")
    else:
        if st.confirm(f"You entered {wf_number}. Is this correct?"):
            wf_number_int = int(wf_number)
            project_info = project_details[project_details['WFNO'] == wf_number_int]
            segment_info = segment_details[segment_details['WFNO'] == wf_number_int]

            if project_info.empty or segment_info.empty:
                st.error("No matching details found for this Workfront number.")
            else:
                campaign_req = project_info.iloc[0]['Campaign Requirements']
                suppressions = [field for field in ['Marketing', 'Risk', 'Output'] if project_info.iloc[0][field] == 'Y']
                outfile_type = project_info.iloc[0]['Outfile Required']
                misc_info = project_info.iloc[0]['Misc']

                standard_prompt = f"Generate SAS code for this campaign from '{campaign_req}' with suppressions: {', '.join(suppressions)}. Outfile type: {outfile_type}. Misc info: {misc_info}."

                llm = ChatOpenAI(model_name="gpt-4-turbo", temperature=0, openai_api_key=openai_api_key)
                prompt_template = PromptTemplate(
                    input_variables=["context", "question"],
                    template="Context: {context}\n\nTask: {question}",
                )

                qa_chain = RetrievalQA.from_chain_type(
                    llm=llm,
                    chain_type="stuff",
                    retriever=vector_db.as_retriever(),
                    chain_type_kwargs={"prompt": prompt_template},
                )

                sas_code_response = qa_chain.run(standard_prompt)

                st.subheader("📄 Generated SAS Code")
                st.code(sas_code_response, language='sas')

                wb = Workbook()
                ws1 = wb.active
                ws1.title = "Project Details"
                for r in dataframe_to_rows(project_info, index=False, header=True):
                    ws1.append(r)

                wb.create_sheet("Waterfall").append(["Waterfall Data Placeholder"])
                ws3 = wb.create_sheet("Segment Details")
                for r in dataframe_to_rows(segment_info, index=False, header=True):
                    ws3.append(r)
                wb.create_sheet("Output File Layout").append(["Output File Layout", outfile_type])

                excel_buffer = BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)

                try:
                    msg = EmailMessage()
                    msg['Subject'] = f"Campaign Details - Workfront {wf_number}"
                    msg['From'] = "mirza.22sept@gmail.com"
                    msg['To'] = user_email
                    msg.set_content(f"Attached campaign details for Workfront Number {wf_number}.")
                    msg.add_attachment(excel_buffer.read(), maintype='application', subtype='xlsx', filename=f"Campaign_{wf_number}.xlsx")

                    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
                        smtp.starttls()
                        smtp.login("mirza.22sept@gmail.com", "your_password")
                        smtp.send_message(msg)

                    st.success(f"Excel file successfully sent to {user_email}")
                except Exception as e:
                    st.error(f"Email sending failed: {e}")
