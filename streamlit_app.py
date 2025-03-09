import streamlit as st
import pandas as pd
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import FAISS
from langchain.chat_models import ChatOpenAI
from langchain.chains import RetrievalQA
from langchain.prompts import PromptTemplate
from langchain.text_splitter import RecursiveCharacterTextSplitter
import docx
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import smtplib
from email.message import EmailMessage

# Load Excel data
excel_path = "project_segment_details.xlsx"
data = pd.ExcelFile(excel_path)
project_details = data.parse("project details")
segment_details = data.parse("segment details")

# Load campaign and SAS code into vector DB
sas_code_doc = docx.Document("sas_code_example.docx")
campaign_doc = docx.Document("campaign_requirements_example.docx")

sas_text = "\n".join([para.text for para in sas_code_doc.paragraphs])
campaign_text = "\n".join([para.text for para in campaign_doc.paragraphs])

text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
sas_chunks = text_splitter.split_text(sas_text)
campaign_chunks = text_splitter.split_text(campaign_text)

openai_api_key = os.getenv("OPENAI_API_KEY")
embeddings = OpenAIEmbeddings(openai_api_key=openai_api_key)
vector_db = FAISS.from_texts(sas_chunks + campaign_chunks, embeddings)

# Streamlit UI setup
st.title("Campaign Operation Programming Solution")

def is_numeric(input_str):
    return input_str.isdigit()

wf_number = st.text_input("Enter Workfront Number (Numeric Only)")
user_email = st.text_input("Enter Your Email")

if st.button("Submit"):
    if not is_numeric(wf_number):
        st.error("Workfront number must be numeric.")
    else:
        if st.confirm(f"You entered {wf_number}. Is this correct?"):
            wf_number_int = int(wf_number)
            project_info = project_details[project_details['WFNO'] == wf_number_int]
            segment_info = segment_details[segment_details['WFNO'] == wf_number_int]

            if project_info.empty or segment_info.empty:
                st.error("No matching details found for this Workfront number.")
            else:
                campaign_req = project_info.iloc[0]['Campaign Requirement']
                suppressions = [field for field in ['Marketing', 'Risk', 'Output'] if project_info.iloc[0][field] == 'Y']
                outfile_type = project_info.iloc[0]['Outfile']
                misc_info = project_info.iloc[0]['Misc']

                standard_prompt = f"Generate SAS code for this campaign from '{campaign_req}' with suppressions: {', '.join(suppressions)}. Output file type: {outfile_type}. Miscellaneous info: {misc_info}."

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

                st.subheader("Generated SAS Code")
                st.code(sas_code_response, language='sas')

                wb = Workbook()

                ws1 = wb.active
                ws1.title = "Project Details"
                for r in dataframe_to_rows(project_info, index=False, header=True):
                    ws1.append(r)

                ws2 = wb.create_sheet("Waterfall")
                ws2.append(["Waterfall Data Placeholder"])

                ws3 = wb.create_sheet("Segment Details")
                for r in dataframe_to_rows(segment_info, index=False, header=True):
                    ws3.append(r)

                ws4 = wb.create_sheet("Output File Layout")
                ws4.append(["Output File Layout", outfile_type])

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
