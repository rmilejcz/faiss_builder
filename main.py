import os
import logging
from docx import Document
import pandas as pd
from langchain.document_loaders import PyPDFLoader, CSVLoader, TXTLoader
from langchain.vectorstores import FAISS
from langchain.embeddings.openai import OpenAIEmbeddings

INDEX_FILE_NAME = "faiss.index"

logging.basicConfig(level=logging.INFO)

def doc_to_txt(doc_path):
    try:
        document = Document(doc_path)
        text = "\n".join([paragraph.text for paragraph in document.paragraphs])
        with open(f"{doc_path}.txt", "w") as text_file:
            text_file.write(text)
    except Exception as e:
        logging.warning(f"Could not process {doc_path}: {str(e)}")

def xls_to_csv(xls_path):
    try:
        read_file = pd.read_excel(xls_path)
        read_file.to_csv(f"{xls_path}.csv", index=None, header=True)
    except Exception as e:
        logging.warning(f"Could not process {xls_path}: {str(e)}")

def get_document_loader(file_extension):
    return {
        '.pdf': PyPDFLoader,
        '.csv': CSVLoader,
        '.txt': TXTLoader,
    }.get(file_extension)

root_dir = '/path/to/your/directory'
openai_key = 'YOUR_OPENAI_KEY'
index_filename = INDEX_FILE_NAME
embeddings = OpenAIEmbeddings(openai_api_key=openai_key)
faiss_index = None

for root, dirs, files in os.walk(root_dir):
    for name in files:
        try:
            _, file_extension = os.path.splitext(name)
            file_path = os.path.join(root, name)
            if file_extension in ['.doc', '.docx']:
                doc_to_txt(file_path)
                file_path = f"{file_path}.txt"
            elif file_extension in ['.xls', '.xlsx']:
                xls_to_csv(file_path)
                file_path = f"{file_path}.csv"
            loader = get_document_loader(file_extension)
            if loader:
                document = loader(file_path).load()
                if not faiss_index:
                    faiss_index = FAISS.from_document(document, embeddings)
                else:
                    faiss_index.add_document(document)
        except Exception as e:
            logging.warning(f"Error processing file: {file_path}. Error: {str(e)}")

faiss_index.save_local(index_filename)
