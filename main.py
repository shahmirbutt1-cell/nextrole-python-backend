from fastapi import FastAPI, UploadFile, File
from docx import Document
import shutil

app = FastAPI()

@app.get("/")
def root():
    return {"status": "NextRole Python backend running"}

@app.post("/analyze-resume")
async def analyze_resume(file: UploadFile = File(...)):
    
    temp_file = f"temp_{file.filename}"
    
    with open(temp_file, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    doc = Document(temp_file)
    
    text = "\n".join([p.text for p in doc.paragraphs])

    return {
        "file_name": file.filename,
        "text_length": len(text),
        "preview": text[:200]
    }
