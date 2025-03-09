import streamlit as st
import pandas as pd
import docx
import os
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_community.chat_models import ChatOpenAI
from langchain.prompts import PromptTemplate
from langchain.chains import RetrievalQA
from langchain.text_splitter import RecursiveCharacterTextSplitter
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
import smtplib
from email.message import EmailMessage

# Verify the API key is loaded
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    st.error("OpenAI API key not found! Please set it in your Streamlit secrets or your .env file.")
    st.stop()

# Verify required files
required_files = ["sas_code_example.docx", "campaign_requirements_example.docx", "project_segment_details.xlsx", "syf logo.png"]
missing_files = [file for file in required_files if not os.path.exists(file)]
if missing_files:
    st.error(f"Missing required file(s): {', '.join(missing_files)}")
    st.stop()

# Display Synchrony logo
st.image("syf logo.png", width=100)
st.title("✨ AI-Based Campaign Operation Programming ✨")

# Load and process Excel data
excel_path = "project_segment_details.xlsx"
data = pd.ExcelFile(excel_path)

if "Project Details" not in data.sheet_names or "Segment Details" not in data.sheet_names:
    st.error(f"Required Excel sheets missing. Available sheets: {data.sheet_names}")
    st.stop()

project_details = data.parse("Project Details")
segment_details = data.parse("Segment Details")

# Load campaign and SAS code
def load_docx(file_path):
    doc = docx.Document(file_path)
    return '\n'.join(para.text for para in doc.paragraphs)

sas_text = load_docx("sas_code_example.docx")
campaign_text = load_docx("campaign_requirements_example.docx")

# Prepare vector DB
text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
sas_chunks = text_splitter.split_text(sas_text)
campaign_chunks = text_splitter.split_text(campaign_text)

embeddings = OpenAIEmbeddings(openai_api_key=openai_api_key)
vector_db = FAISS.from_texts(sas_chunks + campaign_chunks, embeddings)

# UI Input
wf_number = st.text_input("🔢 Enter Workfront Number (Numeric Only)")
user_email = st.text_input("📧 Enter Your Email")

if st.button("🚀 Submit"):
    if not wf_number.isdigit():
        st.error("Workfront number must be numeric.")
    else:
        confirm = st.checkbox(f"Confirm Workfront Number: {wf_number}")
        if confirm:
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

                standard_prompt = (f"Generate SAS code for this campaign from '{campaign_req}'.\n"
                    f"Apply suppressions for any fields that contain 'Y' among: {', '.join(suppressions)}.\n"
                    f"Outfile type: {outfile_type}. If 'EM', use an email file layout; if 'DM', use a direct mail file layout; if 'DE', use both.\n"
                    f"Misc info: {misc_info}.\n")

                # Retrieve relevant SAS code snippets
                retriever = vector_db.as_retriever()
                retrieved_docs = retriever.get_relevant_documents(standard_prompt)
                retrieved_text = '\n'.join([doc.page_content for doc in retrieved_docs])

                st.write("**Debugging Info:**")
                st.write("Campaign Requirement:", campaign_req)
                st.write("Suppressions Applied:", suppressions)
                st.write("Outfile Type:", outfile_type)
                st.write("Misc Info:", misc_info)
                st.write("Retrieved SAS Snippets:", retrieved_text)

                if not retrieved_text.strip():
                    st.error("No relevant SAS code found in the vector database.")
                else:
                    try:
                        llm = ChatOpenAI(model_name="gpt-4o-mini", temperature=0, openai_api_key=openai_api_key)
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
                        
                        st.write("**Executing LLM Query with Prompt:**", standard_prompt)
                        sas_code_response = qa_chain.run(retrieved_text + "\n" + standard_prompt)
                        
                        if not sas_code_response.strip():
                            st.error("LLM did not return a SAS code. Please check OpenAI API usage.")
                        else:
                            st.subheader("📄 Generated SAS Code")
                            st.code(sas_code_response, language='sas')
                    except Exception as e:
                        st.error(f"Error generating SAS code: {e}")
