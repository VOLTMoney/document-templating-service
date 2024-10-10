import base64
from fastapi import Body, FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pydantic import Json
from docxtpl import DocxTemplate
from typing import Any, Dict
import aiofiles
from utils import remove_temporary_files, get_env
import requests
import uuid
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)  # Set to DEBUG for more detailed logs
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Document Template Processing Service",
    description="""
        This is the documentation of the REST API exposed by the document template processing microservice.
        This will allow you to inject data in a specific word document template and get the pdf format as a result. 🚀🚀🚀
    """,
    version="1.0.0"
)

SERVICE_STATUS = {'status': 'Service is healthy !'}

@app.get('/')
async def livenessprobe():
    remove_temporary_files()
    return SERVICE_STATUS

@app.get('/health-check')
async def healthcheck():
    remove_temporary_files()
    return SERVICE_STATUS

@app.post('/api/v1/process-template-document')
async def process_document_template(data: Json = Body(...), file: UploadFile = File(...)):
    if file.filename == '':
        return JSONResponse({'status': 'error', 'message': 'file is required'}, status_code=400)
    if data is None or len(data) == 0:
        return JSONResponse({'status': 'error', 'message': 'data is required'}, status_code=400)
    resourceURL = '{}/forms/libreoffice/convert'.format(get_env('GOTENBERG_API_URL')) 
    file_path = 'temp/{}'.format(file.filename)
    pdf_file_path = 'temp/{}.pdf'.format(file.filename.split('.')[0])
    async with aiofiles.open(file_path, 'wb') as out_file:
        while content := await file.read(1024):
            await out_file.write(content)
    document = DocxTemplate(file_path)
    document.render(data)
    document.save(file_path)
    response = requests.post(url=resourceURL, files={'file': open(file_path, 'rb')})
    async with aiofiles.open(pdf_file_path, 'wb') as out_file:
        await out_file.write(response.content)
    return FileResponse(pdf_file_path, media_type='application/pdf')

# Define a custom temporary folder to store uploaded files
UPLOAD_FOLDER = 'temp/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Only allow .docx files
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.post('/api/v1/process-template-document/upload-file')
async def process_document_template(file: UploadFile = File(...)):
    # Check if the uploaded file is a .docx file
    if not allowed_file(file.filename):
        return JSONResponse(content={"error": "Only .docx files are allowed"}, status_code=400)

    # Save the uploaded file to the custom temporary folder
    file_location = os.path.join(UPLOAD_FOLDER, file.filename)

    with open(file_location, "wb") as buffer:
        buffer.write(await file.read())

    # Return success message with file path
    return {"message": "File uploaded successfully", "file_path": file_location}



@app.post('/api/v1/process-template-document/tata-kfs-review')
async def process_document_template(data: Dict[str, Any] = Body(...)):
    logger.info("Received request to process document template data {}".format(data))
    if not data or 'fileName' not in data or 'data' not in data:
        return JSONResponse({'status': 'error', 'message': 'fileName and data are required'}, status_code=400)

    resourceURL = f"{get_env('GOTENBERG_API_URL')}/forms/libreoffice/convert"
    file_name = data['fileName'].replace('.docx', '')  # Remove the extension for filename purposes
    file_path = f'temp/{data["fileName"]}'

    # Generate unique filenames
    unique_id = str(uuid.uuid4())
    modified_file_path = f'temp/modified_{file_name}_{unique_id}.docx'
    pdf_file_path = f'temp/{file_name}_{unique_id}.pdf'

    # Load and modify the document
    document = DocxTemplate(file_path)
    document.render(data['data'])
    document.save(modified_file_path)

    # Convert to PDF
    try:
        with open(modified_file_path, 'rb') as f:
            response = requests.post(url=resourceURL, files={'file': f})
            response.raise_for_status()  # Check for errors in the response
    except Exception as e:
        return JSONResponse({'status': 'error', 'message': str(e)}, status_code=500)

    # Save the PDF
    async with aiofiles.open(pdf_file_path, 'wb') as out_file:
        await out_file.write(response.content)

    # Read the PDF file and encode to Base64
    async with aiofiles.open(pdf_file_path, 'rb') as out_file:
        pdf_content = await out_file.read()
        pdf_base64 = base64.b64encode(pdf_content).decode('utf-8')

    return JSONResponse({'status': 'success', 'pdf_base64': pdf_base64})
