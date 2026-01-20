# main.py
from fastapi import FastAPI, UploadFile, File
from typing import Annotated
from generator.generator import CertificateGenerator
from fastapi.responses import FileResponse
import shutil

# Create a FastAPI instance
app = FastAPI()

# Define a GET path operation decorator
@app.get("/", tags=["Root"])
def read_root():
    return {"message": "Hello World"}


@app.post("/gerar", tags=["Certificados"])
def read_certificados(file: Annotated[UploadFile, File(...)]):
    with open(f"../generator/table/atletas.xlsx", "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    generator = CertificateGenerator()
    result = generator.generate()
    if result:
        # Return the FileResponse
        return FileResponse(
            path='../generator/certificates/certificados_final.docx',
            filename='certificados_final.docx', # The name the browser will use for the download
            media_type='application/octet-stream' # Generic media type for binary data
        )
    else:
        # Return the FileResponse
        return FileResponse(
            path='../generator/logs/error_log.txt',
            filename='error_log.txt', # The name the browser will use for the download
            media_type='application/octet-stream' # Generic media type for binary data
        )

