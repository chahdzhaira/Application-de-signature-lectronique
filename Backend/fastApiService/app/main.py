from typing import Optional
from fastapi import FastAPI, HTTPException , Header , Depends
from dotenv import load_dotenv
import os
from msal import ConfidentialClientApplication
import requests
import fitz  # PyMuPDF
from io import BytesIO
from fastapi.responses import StreamingResponse
from PIL import Image
import io
from pymongo import MongoClient
from datetime import datetime
from pydantic import BaseModel
import base64
import uuid
from fastapi.middleware.cors import CORSMiddleware

# pyHanko imports
from pyhanko.sign import signers
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign.fields import SigFieldSpec , append_signature_field
from pyhanko.sign.signers.pdf_signer import PdfSignatureMetadata
from pyhanko.pdf_utils.generic import DictionaryObject, NameObject, ArrayObject

#hash
import hashlib


# FastAPI instance
app = FastAPI()

origins = [
    "http://localhost:3000",  
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,           
    allow_credentials=True,
    allow_methods=["*"],             
    allow_headers=["*"],             
)

# Load environment variables
load_dotenv()

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
SHAREPOINT_TENANT_DOMAIN = os.getenv("SHAREPOINT_TENANT_DOMAIN")
SHAREPOINT_SITE_NAME = os.getenv("SHAREPOINT_SITE_NAME")
LIBRARY_NAME = os.getenv("LIBRARY_NAME")
SECRET_API_TOKEN = os.getenv("SECRET_API_TOKEN")
# MongoDB connection
client = MongoClient("mongodb://localhost:27017/")
db = client["Electronic_Signature_App"]
collection = db["Documents"]

# Pydantic Models
class DocumentMetadata(BaseModel):
    file_id: str
    file_name: str
    file_size: int
    signed: bool
    created_at: datetime
    modified_at: datetime
    sharepoint_url: str
    download_url: str
    signed_by: str 
    signerEmail : str
    signed_at: datetime
    jobTitle : str
    original_filename:str
    signature_code: str
    pdf_hash: str
    signer_index: Optional[int] = None
    total_signers: Optional[int] = 1
    signingMode: str
    requestorEmail : str

class SignatureBase64Payload(BaseModel):
    pdf_file: str
    image_file: str
    original_filename:str
    signed_by: str 
    signerEmail : str
    signed_at: datetime
    jobTitle : str
    signer_index: Optional[int] = 1
    total_signers: Optional[int] = 1
    signature_field_name: Optional[str] = "signature1"
    signingMode : str
    requestorEmail : str


# Auth to Microsoft Graph
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["https://graph.microsoft.com/.default"]

msal_app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

def get_access_token():
    result = msal_app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception("Erreur lors de l'obtention du token Graph")

# Get site ID
url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_TENANT_DOMAIN}:/sites/{SHAREPOINT_SITE_NAME}"
headers = { "Authorization": f"Bearer {get_access_token()}" }
response = requests.get(url, headers=headers)

if response.status_code == 200:
    site_id = response.json().get("id")
    print(site_id)
else:
    print("Erreur site ID:", response.status_code, response.text)
    exit()

# Get drive ID
drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
response_drive = requests.get(drive_url, headers=headers)

if response_drive.status_code == 200:
    drives = response_drive.json()['value']
    documents_library_drive = next((d for d in drives if d['name'] == LIBRARY_NAME), None)

    if not documents_library_drive:
        print(f"Erreur : bibliothèque '{LIBRARY_NAME}' non trouvée.")
        exit()

    drive_id = documents_library_drive['id']
    print(drive_id)
else:
    print("Erreur lors de la récupération des bibliothèques :", response_drive.status_code, response_drive.text)
    exit()
    



def verify_token(authorization: Optional[str] = Header(None)):
    if authorization is None or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing or invalid Authorization header")

    token = authorization.split(" ")[1]
    if token != SECRET_API_TOKEN:
        raise HTTPException(status_code=403, detail="Invalid token")

    return True


@app.get("/hello")
async def hello():
    return {"message": "Hello World !"}

@app.post("/add_signature/")
async def add_signature(payload: SignatureBase64Payload, token_ok: bool = Depends(verify_token)):
    try:
        # Decode PDF
        try:
            pdf_bytes = base64.b64decode(payload.pdf_file)
            pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
            # Ajouter une page blanche si le nombre de signataires dépasse 4
            if payload.signer_index and payload.signer_index > 4:
                width = pdf_document[-1].rect.width
                height = pdf_document[-1].rect.height
                pdf_document.new_page(width=width, height=height)  # ajoute à la fin

        except Exception:
            raise HTTPException(status_code=400, detail="Erreur décodage PDF")

        # Decode Image
        try:
            image_bytes = base64.b64decode(payload.image_file)
            img = Image.open(BytesIO(image_bytes))
        except Exception:
            raise HTTPException(status_code=400, detail="Erreur décodage image")

        # Resize image
        img = img.resize((150, 86), Image.LANCZOS)
        img_byte_arr = BytesIO()
        img.save(img_byte_arr, format="PNG")
        img_byte_arr.seek(0)

        # Rechercher tous les documents avec le même original_filename dans MongoDB
        document_metadata = collection.find({"original_filename": payload.original_filename})

        if not document_metadata:
            signatures_count = 0
        else:
            signatures_count = collection.count_documents({"original_filename": payload.original_filename})

        signer_info = f"{payload.signed_by}\n{payload.signed_at.strftime('%Y-%m-%d %H:%M')}"

        # Insert image in last page
        last_page = pdf_document[-1]
        pw, ph = last_page.rect.width, last_page.rect.height

        signature_spacing = 123  
        
        # Si c'est la première signature, la placer à la position par défaut
        if signatures_count == 0:
            x, y = 50, ph - 106  
        else:
            x = 50 + (signatures_count * signature_spacing)
            y = ph - 106        
        last_page.insert_image((x, y, x + 150, y + 86), stream=img_byte_arr)

        formatted_date = payload.signed_at.strftime("%d/%m/%Y %H:%M")

        signature_code = uuid.uuid4().hex[:10]

        text_to_display = f"{payload.jobTitle}\n{payload.signed_by} - {formatted_date}\nVerification code: {signature_code}"

        text_x = x
        text_y = y - 15 
        last_page.insert_text((text_x, text_y), text_to_display, fontsize=8, fontname="helv", fill=(0, 0, 0))

        


        # Save modified PDF
        modified_pdf = BytesIO()
        pdf_document.save(modified_pdf)
        pdf_document.close()

        modified_pdf.seek(0)

         # --- Signature numérique avec pyHanko ---
        if payload.signer_index == payload.total_signers:
       
            # Chargement du la clé privée et le certificat 
            signer = signers.SimpleSigner.load(
                key_file = "app/cert/private_key.pem",
                cert_file = "app/cert/cert.pem",
                key_passphrase=None  
            )

            # Préparer le writer en mode incrémental pour ne pas casser la structure du PDF
            pdf_input = modified_pdf.read()
            pdf_in = BytesIO(pdf_input)
            w = IncrementalPdfFileWriter(pdf_in)

            signature_field_name = f"Signature{payload.signer_index }"  

            if '/AcroForm' not in w.root:
                w.root['/AcroForm'] = DictionaryObject()
                w.root['/AcroForm']['/Fields'] = ArrayObject()


            existing_fields = w.root['/AcroForm']['/Fields']
            field_names = []
            for field in existing_fields:
                obj = field.get_object()
                name = obj.get('/T') if obj else None
                if name:
                    field_names.append(name)
                    
            if signature_field_name not in field_names:
                append_signature_field(w, SigFieldSpec(sig_field_name=signature_field_name))

            signed_pdf_bytes = BytesIO()
            await signers.async_sign_pdf(
                w,
                signer=signer,
                output=signed_pdf_bytes,
                signature_meta=PdfSignatureMetadata(field_name=signature_field_name),
                existing_fields_only=True
            )
            signed_pdf_bytes.seek(0)
            pdf_hash = hashlib.sha256(signed_pdf_bytes.read()).hexdigest()

        else:
            signed_pdf_bytes = modified_pdf
            pdf_hash = hashlib.sha256(signed_pdf_bytes.read()).hexdigest()

        # Generate file name
        base_name = os.path.splitext(payload.original_filename)[0]
        timestamp = payload.signed_at.strftime("%Y%m%d_%H%M%S")
        signed_file_name = f"{base_name}_signed_{timestamp}.pdf"

        upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/SignedDoc/{signed_file_name}:/content"
        upload_headers = {
            "Authorization": f"Bearer {get_access_token()}",
            "Content-Type": "application/pdf"
        }

        response_upload = requests.put(upload_url, headers=upload_headers, data=signed_pdf_bytes.getvalue())
        if response_upload.status_code != 201:
            return {"error": f"Erreur upload : {response_upload.status_code} - {response_upload.text}"}

        uploaded_file_info = response_upload.json()
        uploaded_file_id = uploaded_file_info.get("id")

        # Update "Signed" field
        update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{uploaded_file_id}/listItem/fields"
        print(f"Updating field at: {update_url}")

        update_headers = {
            "Authorization": f"Bearer {get_access_token()}",
            "Content-Type": "application/json"
        }

        update_data = {
            "Status": "Signed",
            "SignedBy": payload.signerEmail
        }

        response_update = requests.patch(update_url, headers=update_headers, json=update_data)
        print(f"Update response: {response_update.status_code} - {response_update.text}")

        if response_update.status_code != 200:
            return {"error": f"Error updating Signed field : {response_update.status_code} - {response_update.text}"}

        response_upload_json = response_upload.json()
        response_update_json = response_update.json()
        

        # Save in MongoDB
        metadata = DocumentMetadata(
            file_id=uploaded_file_id,
            file_name=signed_file_name,
            file_size=len(signed_pdf_bytes.getvalue()),
            signed=True,
            created_at=datetime.utcnow(),
            modified_at=datetime.utcnow(),
            sharepoint_url=f"https://{SHAREPOINT_TENANT_DOMAIN}.sharepoint.com/sites/{SHAREPOINT_SITE_NAME}/_layouts/15/Doc.aspx?sourcedoc={uploaded_file_id}",
            download_url=f"https://{SHAREPOINT_TENANT_DOMAIN}.sharepoint.com/sites/{SHAREPOINT_SITE_NAME}/_layouts/15/Download.aspx?file={signed_file_name}",
            original_filename=payload.original_filename,
            signed_by=payload.signed_by,
            signerEmail=payload.signerEmail,
            signed_at=payload.signed_at,
            jobTitle=payload.jobTitle,
            signature_code=signature_code,
            pdf_hash=pdf_hash, 
            signer_index = payload.signer_index,
            total_signers=payload.total_signers,
            signingMode = payload.signingMode,
            requestorEmail = payload.requestorEmail,

        )

        metadata_dict = metadata.dict()
        print(f"Metadata inserted: {metadata_dict}")
        collection.insert_one(metadata_dict)

        return {"message": "Fichier signé et uploadé avec succès.", "file_name": signed_file_name}
        
    except Exception as e:
        return {"error": str(e)}
