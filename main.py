import base64
from fastapi import Body, FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pydantic import Json
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from io import BytesIO
from typing import Any, Dict
import aiofiles
from utils import remove_temporary_files, get_env
import requests
import uuid
import io
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


async def process_parallel_sections(data: Dict[str, Any]):
    """Helper function for parallel processing of document sections"""
    # Ensure directories exist
    temp_dir = 'temp'
    sections_dir = 'docx-sections'
    for dir_path in [temp_dir, sections_dir]:
        if not os.path.exists(dir_path):
            os.makedirs(dir_path)

    resourceURL = f"{get_env('GOTENBERG_API_URL')}/forms/libreoffice/convert"
    folder_name = data['folderName']
    folder_path = f'{sections_dir}/{folder_name}'

    if not os.path.exists(folder_path):
        return JSONResponse({'status': 'error', 'message': f'Section folder not found: {folder_name}'}, status_code=404)

    import threading
    from concurrent.futures import ThreadPoolExecutor, as_completed

    try:
        # Get all files in the folder (both DOCX and PDF)
        all_files = os.listdir(folder_path)
        all_docx_files = [f for f in all_files if f.endswith('.docx')]
        all_pdf_files = [f for f in all_files if f.endswith('.pdf')]

        # Sort files to ensure proper order (handles 01, 02, 03... 10, 11 correctly)
        def natural_sort_key(filename):
            import re
            # Extract number from filename (e.g., "01_dynamic_agreement.docx" -> 1)
            match = re.search(r'^(\d+)_', filename)
            if match:
                return int(match.group(1))
            return 0  # Files without numbers go first

        all_docx_files.sort(key=natural_sort_key)
        all_pdf_files.sort(key=natural_sort_key)

        # Separate dynamic and static files
        dynamic_files = [f for f in all_docx_files if '_dynamic_' in f]
        static_pdf_files = [f for f in all_pdf_files if '_static_' in f]

        # Static files are always PDF-only - create mapping for ordering
        static_files = []
        static_file_mapping = {}  # Maps display name to actual file path

        for static_pdf in static_pdf_files:
            # Create a display name using the DOCX convention for ordering
            static_docx = static_pdf.replace('.pdf', '.docx')
            static_files.append(static_docx)  # Use DOCX name for ordering
            static_file_mapping[static_docx] = os.path.join(folder_path, static_pdf)

        # Create the complete file list for merging (all files in order)
        all_files_for_merging = []

        # Combine all files and sort by number
        all_file_names = dynamic_files + static_files
        all_file_names.sort(key=natural_sort_key)

        for file_name in all_file_names:
            if file_name in dynamic_files:
                all_files_for_merging.append(file_name)
            elif file_name in static_files:
                all_files_for_merging.append(file_name)

        if not all_docx_files:
            return JSONResponse({'status': 'error', 'message': f'No DOCX files found in folder: {folder_name}'}, status_code=404)

        # Thread-safe storage for processed dynamic files
        processed_dynamic_pdfs = {}
        processing_lock = threading.Lock()

        def process_dynamic_section(docx_file, index):
            """Process a single dynamic section in a thread"""
            file_path = os.path.join(folder_path, docx_file)

            try:
                # Process with data
                document = DocxTemplate(file_path)
                context = data['data']

                # Process image if provided (for all dynamic sections)
                if 'image' in data:
                    image_info = data['image']
                    base64_image = image_info.get('content')
                    image_width = image_info.get('width', 2)
                    image_height = image_info.get('height', 2)

                    if base64_image:
                        image_data = base64.b64decode(base64_image)
                        image_file = BytesIO(image_data)
                        context["image_placeholder"] = InlineImage(document, image_file, width=Inches(image_width), height=Inches(image_height))

                # Render with data
                document.render(context)

                # Convert to PDF
                output_stream = BytesIO()
                document.save(output_stream)
                output_stream.seek(0)

                # Convert to PDF
                response = requests.post(
                    url=resourceURL,
                    files={'file': (
                    f'dynamic_{index}.docx', output_stream, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
                )
                response.raise_for_status()

                # Clean up BytesIO stream
                output_stream.close()

                if response.content:
                    # Thread-safe storage of results
                    with processing_lock:
                        processed_dynamic_pdfs[docx_file] = {
                            'content': response.content
                        }

            except Exception as e:
                return None

            return {
                'filename': docx_file
            }

        # Process dynamic files in parallel using ThreadPoolExecutor
        if dynamic_files:
            with ThreadPoolExecutor(max_workers=min(len(dynamic_files), 6)) as executor:
                # Submit all dynamic tasks
                future_to_section = {
                    executor.submit(process_dynamic_section, docx_file, i): (docx_file, i) 
                    for i, docx_file in enumerate(dynamic_files)
                }

                # Wait for all tasks to complete
                completed_sections = []
                for future in as_completed(future_to_section):
                    docx_file, index = future_to_section[future]
                    try:
                        result = future.result()
                        if result:
                            completed_sections.append(result)
                    except Exception as e:
                        pass

        # Load pre-existing static PDFs (no processing needed)
        static_pdfs = {}

        for static_file in static_files:
            try:
                # Use the mapping to get the actual file path
                static_pdf_path = static_file_mapping[static_file]

                if os.path.exists(static_pdf_path):
                    # Read the pre-existing PDF file
                    with open(static_pdf_path, 'rb') as pdf_file:
                        pdf_content = pdf_file.read()

                    static_pdfs[static_file] = {
                        'content': pdf_content
                    }

            except Exception as e:
                continue

        # Merge all PDFs in correct order (both dynamic and static)
        try:
            from PyPDF2 import PdfMerger

            # Create PDF merger
            merger = PdfMerger()

            # Merge in the order of all files (dynamic and static combined)
            pdf_streams = []
            for docx_file in all_files_for_merging:
                if docx_file in processed_dynamic_pdfs:
                    # Add processed dynamic PDF
                    pdf_data = processed_dynamic_pdfs[docx_file]
                    pdf_stream = BytesIO(pdf_data['content'])
                    merger.append(pdf_stream)
                    pdf_streams.append(pdf_stream)
                elif docx_file in static_pdfs:
                    # Add static PDF
                    pdf_data = static_pdfs[docx_file]
                    pdf_stream = BytesIO(pdf_data['content'])
                    merger.append(pdf_stream)
                    pdf_streams.append(pdf_stream)

            # Create merged PDF
            merged_pdf_stream = BytesIO()
            merger.write(merged_pdf_stream)
            merger.close()

            merged_pdf = merged_pdf_stream.getvalue()

            # Clean up PDF streams
            for stream in pdf_streams:
                stream.close()
            merged_pdf_stream.close()

        except Exception as e:
            # Get first available PDF as fallback
            if processed_dynamic_pdfs:
                first_dynamic = list(processed_dynamic_pdfs.keys())[0]
                merged_pdf = processed_dynamic_pdfs[first_dynamic]['content']
            elif static_pdfs:
                first_static = list(static_pdfs.keys())[0]
                merged_pdf = static_pdfs[first_static]['content']
            else:
                return JSONResponse({'status': 'error', 'message': 'No PDFs were successfully processed'}, status_code=500)

        # Encode final PDF
        pdf_base64 = base64.b64encode(merged_pdf).decode('utf-8')

        # Explicit memory cleanup to ensure no memory leaks
        del processed_dynamic_pdfs
        del static_pdfs
        del merged_pdf
        del merged_pdf_stream
        del merger
        del pdf_streams
        del all_files_for_merging
        del static_file_mapping
        del dynamic_files
        del static_files
        del all_docx_files
        del all_pdf_files
        del all_files

    except Exception as e:
        return JSONResponse({'status': 'error', 'message': f"Error in parallel processing: {str(e)}"}, status_code=500)

    return JSONResponse({
        'status': 'success', 
        'pdf_base64': pdf_base64
    })


@app.post('/api/v1/process-template-document/docx-to-pdf')
async def process_document_template(data: Dict[str, Any] = Body(...)):
    # Check if folderName is provided and fileName is empty/null - use parallel processing
    if data and 'folderName' in data and data.get('folderName') and (not data.get('fileName') or data.get('fileName') == ''):
        return await process_parallel_sections(data)

    if not data or 'fileName' not in data or 'data' not in data:
        return JSONResponse({'status': 'error', 'message': 'fileName and data are required'}, status_code=400)

    # Ensure the temp directory exists
    temp_dir = 'temp'
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
        print(f"Created temp directory: {temp_dir}")

    # List the contents of the temp directory
    files_in_temp = os.listdir(temp_dir)
    print(f"Current contents of the temp directory: {files_in_temp}")

    resourceURL = f"{get_env('GOTENBERG_API_URL')}/forms/libreoffice/convert"
    file_name = data['fileName'].replace('.docx', '')  # Remove the extension for filename purposes
    file_path = f'docx-template/{data["fileName"]}'

    # Generate unique filenames
    unique_id = str(uuid.uuid4())
    modified_file_path = f'temp/modified_{file_name}_{unique_id}.docx'

    output_stream = BytesIO()

    # Load and modify the document
    try:
        document = DocxTemplate(file_path)
        # Start with the provided data as the context
        context = data['data']
        # Process image if provided in nested 'image' data
        if 'image' in data:
            image_info = data['image']
            base64_image = image_info.get('content')
            image_width = image_info.get('width', 2)  # Default to 2 inches if not provided
            image_height = image_info.get('height', 2)  # Default to 2 inches if not provided

            if base64_image:
                # Decode the base64 string and use BytesIO to create a file-like object
                image_data = base64.b64decode(base64_image)
                image_file = BytesIO(image_data)

                # Add the InlineImage to the context under a key that matches the placeholder in the template
                context["image_placeholder"] = InlineImage(document, image_file, width=Inches(image_width), height=Inches(image_height))

        # Render the document once with the combined context
        document.render(context)

        document.save(output_stream)
        output_stream.seek(0)  # Reset stream position for reading
    except Exception as e:
        return JSONResponse({'status': 'error', 'message': f"Error rendering or saving docx: {str(e)}"}, status_code=500)

    # Convert to PDF
    try:
        response = requests.post(
            url=resourceURL,
            files={'file': (
            'modified.docx', output_stream, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
        )
        response.raise_for_status()  # Check for errors in the response
    except requests.exceptions.RequestException as e:
        return JSONResponse({'status': 'error', 'message': f"PDF conversion failed: {str(e)}"}, status_code=500)

    # Ensure the response contains the PDF content
    if not response.content:
        return JSONResponse({'status': 'error', 'message': 'PDF conversion returned empty content'}, status_code=500)

    # Directly encode PDF content to Base64 without saving it
    try:
        pdf_base64 = base64.b64encode(response.content).decode('utf-8')
    except Exception as e:
        return JSONResponse({'status': 'error', 'message': f"Error encoding PDF to Base64: {str(e)}"}, status_code=500)

    return JSONResponse({'status': 'success', 'pdf_base64': pdf_base64})


