from pydantic import BaseModel
from typing import List, Optional

class LoginRequest(BaseModel):
    username: str
    password: str

class User(BaseModel):
    email: str
    clients: List[str]
    access: str

class UserCreate(BaseModel):
    email: str
    clients: List[str]
    access: str

class FolderRequest(BaseModel):
    folder: str

class TemplateRequest(BaseModel):
    folder: str
    dl_type: str
    template: str

class ModeRequest(BaseModel):
    mode: str

class OutputFormatRequest(BaseModel):
    format: str  # "zip" or "print"

class AuditEntry(BaseModel):
    client: str
    processed_by: str
    total_accounts: int
    mode: str
    template_folder: Optional[str] = None
    dl_type: Optional[str] = None
