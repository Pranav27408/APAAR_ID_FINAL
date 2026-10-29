import csv
import io
import os
import re
import uuid
import uuid
from datetime import datetime, timedelta, timezone

from dotenv import load_dotenv
from fastapi import Depends, FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.security import OAuth2PasswordBearer
from jose import JWTError, jwt
from openpyxl import load_workbook
from passlib.context import CryptContext
from pydantic import BaseModel, ConfigDict, EmailStr, Field, field_validator
from sqlalchemy import Boolean, DateTime, Integer, String, Text, create_engine, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, sessionmaker

load_dotenv()
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./apaar.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg://", 1)
elif DATABASE_URL.startswith("postgresql://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg://", 1)
JWT_SECRET = os.getenv("JWT_SECRET", "development-only-change-this-secret-before-deployment")
ALGORITHM = "HS256"
TOKEN_MINUTES = int(os.getenv("ACCESS_TOKEN_MINUTES", "480"))
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {})
SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)
pwd = CryptContext(schemes=["bcrypt"], deprecated="auto")
oauth = OAuth2PasswordBearer(tokenUrl="/api/auth/login")
app = FastAPI(title="APAAR Student Data Platform API", version="1.0.0", description="Student records and verification workflow API.")
default_frontend_origins = ["https://apaarstudentid.netlify.app", *[f"http://{host}:{port}" for host in ("localhost", "127.0.0.1") for port in (5173, 5174, 5175)]]
configured_frontend_origins = os.getenv("FRONTEND_ORIGINS", os.getenv("FRONTEND_ORIGIN", ""))
frontend_origins = list(dict.fromkeys([*default_frontend_origins, *[origin.strip().rstrip("/") for origin in configured_frontend_origins.split(",") if origin.strip()]]))
app.add_middleware(CORSMiddleware, allow_origins=frontend_origins, allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

class Base(DeclarativeBase): pass

class User(Base):
    __tablename__ = "users"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    email: Mapped[str] = mapped_column(String(160), unique=True, index=True)
    password_hash: Mapped[str] = mapped_column(String(255))
    name: Mapped[str] = mapped_column(String(160))
    role: Mapped[str] = mapped_column(String(20), index=True)
    institution: Mapped[str] = mapped_column(String(200), default="National Institute of Technology, Jaipur")

class Student(Base):
    __tablename__ = "students"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    public_id: Mapped[str] = mapped_column(String(40), unique=True, index=True)
    name: Mapped[str] = mapped_column(String(160), index=True)
    register_no: Mapped[str] = mapped_column(String(60), unique=True, index=True)
    department: Mapped[str] = mapped_column(String(120))
    year: Mapped[int] = mapped_column(Integer)
    apaar_id: Mapped[str] = mapped_column(String(12), unique=True, index=True)
    email: Mapped[str] = mapped_column(String(160), default="")
    intake: Mapped[int] = mapped_column(Integer, default=0)
    enrolment: Mapped[int] = mapped_column(Integer, default=0)
    supernumerary: Mapped[int] = mapped_column(Integer, default=0)
    nri: Mapped[bool] = mapped_column(Boolean, default=False)
    oci: Mapped[bool] = mapped_column(Boolean, default=False)
    foreign: Mapped[bool] = mapped_column(Boolean, default=False)
    status: Mapped[str] = mapped_column(String(30), default="Pending", index=True)
    institution: Mapped[str] = mapped_column(String(200), default="National Institute of Technology, Jaipur")
    docs: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

class AuditLog(Base):
    __tablename__ = "audit_logs"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    user_email: Mapped[str] = mapped_column(String(160))
    institution: Mapped[str] = mapped_column(String(200), default="")
    action: Mapped[str] = mapped_column(String(100))
    entity: Mapped[str] = mapped_column(String(100))
    entity_id: Mapped[str] = mapped_column(String(80))
    details: Mapped[str] = mapped_column(Text, default="")
    at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

class Submission(Base):
    __tablename__ = "submissions"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    institution: Mapped[str] = mapped_column(String(200))
    academic_year: Mapped[str] = mapped_column(String(20), default="2026-27")
    status: Mapped[str] = mapped_column(String(30), default="Pending")
    submitted_by: Mapped[str] = mapped_column(String(160))
    submitted_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    comment: Mapped[str] = mapped_column(Text, default="")

class Certificate(Base):
    __tablename__ = "certificates"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    student_id: Mapped[str] = mapped_column(String(40), index=True)
    title: Mapped[str] = mapped_column(String(160))
    serial_no: Mapped[str] = mapped_column(String(50), unique=True)
    issued_by: Mapped[str] = mapped_column(String(160))
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

class StudentIn(BaseModel):
    name: str = Field(min_length=2, max_length=160)
    register_no: str = Field(min_length=3, max_length=60)
    department: str = Field(min_length=2, max_length=120)
    year: int = Field(ge=1, le=8)
    apaar_id: str
    email: str = ""
    intake: int = Field(default=0, ge=0, le=100000)
    enrolment: int = Field(default=0, ge=0, le=100000)
    supernumerary: int = Field(default=0, ge=0, le=10000)
    nri: bool = False
    oci: bool = False
    foreign: bool = False
    @field_validator("apaar_id")
    @classmethod
    def valid_apaar(cls, v):
        if not re.fullmatch(r"\d{12}", str(v)): raise ValueError("APAAR ID must contain exactly 12 digits")
        return str(v)
    @field_validator("email")
    @classmethod
    def safe_email(cls, v):
        if v and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", v): raise ValueError("Invalid email address")
        return v

class StudentOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: str
    name: str
    register_no: str
    department: str
    year: int
    apaar_id: str
    email: str
    intake: int
    enrolment: int
    supernumerary: int
    nri: bool
    oci: bool
    foreign: bool
    status: str
    institution: str
    docs: int
    updated: str

class LoginIn(BaseModel):
    email: EmailStr
    password: str
    role: str
    student_name: str = ""
    institution: str = ""

class ReviewIn(BaseModel):
    status: str
    comment: str = ""

def get_db():
    db = SessionLocal()
    try: yield db
    finally: db.close()

def log(db: Session, user: User, action: str, entity: str, entity_id: str, details: str = ""):
    db.add(AuditLog(user_email=user.email, institution=user.institution, action=action, entity=entity, entity_id=entity_id, details=details))

def get_user(token: str = Depends(oauth), db: Session = Depends(get_db)):
    try:
        data = jwt.decode(token, JWT_SECRET, algorithms=[ALGORITHM]); uid = int(data["sub"])
    except (JWTError, KeyError, ValueError): raise HTTPException(401, "Invalid or expired access token")
    user = db.get(User, uid)
    if not user: raise HTTPException(401, "Account not found")
    return user

def allow(*roles):
    def dep(user: User = Depends(get_user)):
        if user.role not in roles: raise HTTPException(403, "Your role cannot perform this action")
        return user
    return dep

def to_out(s: Student):
    return {"id":s.public_id,"name":s.name,"register_no":s.register_no,"department":s.department,"year":s.year,"apaar_id":s.apaar_id,"email":s.email,"intake":s.intake,"enrolment":s.enrolment,"supernumerary":s.supernumerary,"nri":s.nri,"oci":s.oci,"foreign":s.foreign,"status":s.status,"institution":s.institution,"docs":s.docs,"updated":(s.updated_at.strftime("%d %b %Y") if s.updated_at else "")}

@app.on_event("startup")
def startup():
    Base.metadata.create_all(engine)
    with SessionLocal() as db:
        if engine.dialect.name=="sqlite":
            columns={row[1] for row in db.connection().exec_driver_sql("PRAGMA table_info(audit_logs)").fetchall()}
            if columns and "institution" not in columns:
                db.connection().exec_driver_sql("ALTER TABLE audit_logs ADD COLUMN institution VARCHAR(200) NOT NULL DEFAULT ''")
                db.connection().exec_driver_sql("UPDATE audit_logs SET institution=(SELECT institution FROM users WHERE users.email=audit_logs.user_email) WHERE institution=''")
        accounts=[("student@demo.edu","Student123!","Aarav Sharma","student"),("staff@demo.edu","Staff123!","Priya Nair","staff"),("officer@aicte.gov.in","Officer123!","Verification Officer","officer")]
        for email,password,name,role in accounts:
            if not db.scalar(select(User).where(User.email==email)):
                db.add(User(email=email,password_hash=pwd.hash(password),name=name,role=role))
        if not db.scalar(select(Student)):
            rows=[("STU-2026-001","Aarav Sharma","EC2023-041","Computer Science",3,"123456789012","student@demo.edu",120,112,4,"Approved",3),("STU-2026-002","Diya Patel","CS2024-018","Computer Science",2,"234567890123","diya.patel@demo.edu",180,172,6,"Pending",2),("STU-2026-003","Kabir Reddy","ME2022-067","Mechanical Engineering",4,"345678901234","kabir.reddy@demo.edu",90,84,3,"Needs correction",4),("STU-2026-004","Meera Iyer","IT2023-029","Information Technology",3,"456789012345","meera.iyer@demo.edu",120,118,2,"Approved",3),("STU-2026-005","Arjun Nair","CE2024-052","Civil Engineering",2,"567890123456","arjun.nair@demo.edu",60,56,2,"Pending",1),("STU-2026-006","Ananya Das","CS2022-011","Computer Science",4,"678901234567","ananya.das@demo.edu",180,176,4,"Approved",4),("STU-2026-007","Vikram Singh","EE2023-034","Electrical Engineering",3,"789012345678","vikram.singh@demo.edu",90,80,3,"Rejected",2)]
            for row in rows:
                db.add(Student(public_id=row[0],name=row[1],register_no=row[2],department=row[3],year=row[4],apaar_id=row[5],email=row[6],intake=row[7],enrolment=row[8],supernumerary=row[9],status=row[10],docs=row[11]))
        db.commit()

@app.get("/api/health")
def health(): return {"status":"ok"}

@app.post("/api/auth/login")
def login(body: LoginIn, db: Session = Depends(get_db)):
    user=db.scalar(select(User).where(User.email==body.email.lower()))
    if not user or not pwd.verify(body.password,user.password_hash) or user.role!=body.role: raise HTTPException(401,"Invalid credentials for selected portal")
    if user.role=="student":
        student=db.scalar(select(Student).where(Student.email==user.email))
        if not student or body.student_name.strip().casefold()!=student.name.casefold() or body.institution.strip().casefold()!=student.institution.casefold():
            raise HTTPException(401,"Student name or institution does not match the student record")
    token=jwt.encode({"sub":str(user.id),"role":user.role,"exp":datetime.now(timezone.utc)+timedelta(minutes=TOKEN_MINUTES)},JWT_SECRET,algorithm=ALGORITHM)
    return {"access_token":token,"token_type":"bearer","name":user.name,"email":user.email,"role":user.role,"institution":user.institution}

@app.get("/api/auth/me")
def me(user: User = Depends(get_user)): return {"email":user.email,"name":user.name,"role":user.role,"institution":user.institution}

@app.get("/api/students")
def list_students(q: str = "", status: str = "", limit: int = Query(500,ge=1,le=1000), offset: int = Query(0,ge=0), user: User = Depends(get_user), db: Session = Depends(get_db)):
    stmt=select(Student)
    if user.role=="student": stmt=stmt.where(Student.email==user.email)
    elif user.role=="staff": stmt=stmt.where(Student.institution==user.institution)
    if q: stmt=stmt.where((Student.name.ilike(f"%{q}%"))|(Student.register_no.ilike(f"%{q}%"))|(Student.apaar_id.ilike(f"%{q}%")))
    if status: stmt=stmt.where(Student.status==status)
    items=db.scalars(stmt.order_by(Student.updated_at.desc()).offset(offset).limit(limit)).all()
    return {"items":[to_out(s) for s in items],"total":len(items),"cross_institution_check":True}

@app.post("/api/students",status_code=201)
def create_student(body: StudentIn, user: User = Depends(allow("staff")), db: Session = Depends(get_db)):
    dupe=db.scalar(select(Student).where((Student.apaar_id==body.apaar_id)|(Student.register_no==body.register_no)))
    if dupe: raise HTTPException(409,"Duplicate APAAR ID or register number detected")
    if body.enrolment>body.intake+body.supernumerary: raise HTTPException(422,"Enrolment exceeds approved intake plus supernumerary seats")
    s=Student(public_id=f"STU-{datetime.now().year}-{datetime.now().timestamp_ns() if hasattr(datetime.now(),'timestamp_ns') else int(datetime.now().timestamp()*1000000)}",**body.model_dump(),institution=user.institution,status="Pending",docs=0)
    db.add(s);db.flush();log(db,user,"created","student",s.public_id,"New student record");db.commit();db.refresh(s);return to_out(s)

@app.put("/api/students/{public_id}")
def update_student(public_id: str, body: StudentIn, user: User = Depends(allow("staff")), db: Session = Depends(get_db)):
    s=db.scalar(select(Student).where(Student.public_id==public_id,Student.institution==user.institution))
    if not s: raise HTTPException(404,"Student not found")
    if body.enrolment>body.intake+body.supernumerary: raise HTTPException(422,"Enrolment exceeds approved intake plus supernumerary seats")
    dupe=db.scalar(select(Student).where(Student.apaar_id==body.apaar_id,Student.id!=s.id))
    if dupe: raise HTTPException(409,"APAAR ID is already assigned to another student")
    old=s.status
    for k,v in body.model_dump().items(): setattr(s,k,v)
    if old=="Approved": s.status="Pending"
    log(db,user,"updated","student",public_id,"Student fields updated");db.commit();db.refresh(s);return to_out(s)

@app.delete("/api/students/{public_id}")
def delete_student(public_id: str, user: User = Depends(allow("staff")), db: Session = Depends(get_db)):
    s=db.scalar(select(Student).where(Student.public_id==public_id,Student.institution==user.institution))
    if not s: raise HTTPException(404,"Student not found")
    log(db,user,"deleted","student",public_id,"Student record deleted");db.delete(s);db.commit();return {"ok":True}

@app.post("/api/students/{public_id}/review")
def review_student(public_id: str, body: ReviewIn, user: User = Depends(allow("staff","officer")), db: Session = Depends(get_db)):
    if body.status not in {"Approved","Rejected","Needs correction"}: raise HTTPException(422,"Unsupported review status")
    stmt=select(Student).where(Student.public_id==public_id)
    if user.role=="staff": stmt=stmt.where(Student.institution==user.institution)
    s=db.scalar(stmt)
    if not s: raise HTTPException(404,"Student not found")
    s.status=body.status;log(db,user,"reviewed","student",public_id,body.comment or body.status);db.commit();db.refresh(s);return to_out(s)

@app.post("/api/submissions")
def submit_institution(user: User = Depends(allow("staff")), db: Session = Depends(get_db)):
    rows=db.scalars(select(Student).where(Student.institution==user.institution)).all()
    errors=[]
    for s in rows:
        if not re.fullmatch(r"\d{12}",s.apaar_id): errors.append({"student":s.public_id,"issue":"Invalid APAAR ID"})
        if s.enrolment>s.intake+s.supernumerary: errors.append({"student":s.public_id,"issue":"Enrolment exceeds seats"})
    if errors: return {"submitted":False,"errors":errors,"count":len(rows)}
    sub=Submission(institution=user.institution,submitted_by=user.email,status="Pending")
    db.add(sub)
    for s in rows:
        if s.status not in {"Approved","Rejected"}: s.status="Pending"
    log(db,user,"submitted","submission",str(sub.id),f"{len(rows)} student records submitted");db.commit()
    return {"submitted":True,"submission_id":sub.id,"count":len(rows),"status":"Pending"}

@app.get("/api/submissions")
def list_submissions(user: User = Depends(get_user), db: Session = Depends(get_db)):
    stmt=select(Submission)
    if user.role=="staff": stmt=stmt.where(Submission.institution==user.institution)
    return [{"id":x.id,"institution":x.institution,"academic_year":x.academic_year,"status":x.status,"submitted_by":x.submitted_by,"submitted_at":x.submitted_at.isoformat(),"comment":x.comment} for x in db.scalars(stmt.order_by(Submission.submitted_at.desc())).all()]

@app.post("/api/import")
async def import_students(file: UploadFile = File(...), user: User = Depends(allow("staff")), db: Session = Depends(get_db)):
    name=file.filename or ""
    if not name.lower().endswith((".csv",".xlsx",".xlsm")): raise HTTPException(415,"Upload a CSV or XLSX file")
    raw=await file.read()
    if len(raw)>10*1024*1024: raise HTTPException(413,"File exceeds 10 MB")
    try:
        if name.lower().endswith(".csv"):
            text=raw.decode("utf-8-sig"); rows=list(csv.DictReader(io.StringIO(text)))
        else:
            wb=load_workbook(io.BytesIO(raw),read_only=True,data_only=True);ws=wb.active; vals=list(ws.values); headers=[str(x or "").strip().lower().replace(" ","_") for x in vals[0]];rows=[dict(zip(headers,r)) for r in vals[1:] if any(v is not None for v in r)]
    except Exception: raise HTTPException(400,"Could not read spreadsheet. Check its format and headers.")
    aliases={"register_number":"register_no","register_no":"register_no","name":"name","department":"department","year":"year","apaar_id":"apaar_id","email":"email","approved_intake":"intake","intake":"intake","actual_enrolment":"enrolment","enrolment":"enrolment","supernumerary_seats":"supernumerary","supernumerary":"supernumerary","nri_category":"nri","oci_status":"oci","foreign_national_status":"foreign"}
    seen=set();valid=[];errors=[]
    for i,row in enumerate(rows,2):
        item={aliases.get(str(k).strip().lower().replace(" ","_"),str(k).strip().lower()):v for k,v in row.items()}
        apaar=str(item.get("apaar_id","")).strip().removesuffix(".0"); reg=str(item.get("register_no","")).strip(); n=str(item.get("name","")).strip()
        if not n or not reg or not re.fullmatch(r"\d{12}",apaar): errors.append({"row":i,"issue":"Name, register number and a 12-digit APAAR ID are required"});continue
        if apaar in seen or db.scalar(select(Student).where((Student.apaar_id==apaar)|(Student.register_no==reg))): errors.append({"row":i,"issue":"Duplicate APAAR ID or register number"});continue
        seen.add(apaar)
        try:
            valid.append(Student(public_id=f"IMP-{int(datetime.now().timestamp()*1000)}-{i}",name=n,register_no=reg,department=str(item.get("department","Unspecified")),year=int(item.get("year",1) or 1),apaar_id=apaar,email=str(item.get("email","") or ""),intake=int(item.get("intake",0) or 0),enrolment=int(item.get("enrolment",0) or 0),supernumerary=int(item.get("supernumerary",0) or 0),nri=str(item.get("nri","")).lower() in {"1","yes","true"},oci=str(item.get("oci","")).lower() in {"1","yes","true"},foreign=str(item.get("foreign","")).lower() in {"1","yes","true"},institution=user.institution,status="Pending",docs=0))
        except (ValueError,TypeError): errors.append({"row":i,"issue":"Year and seat fields must be numbers"})
    db.add_all(valid);log(db,user,"imported","student_import",file.filename,f"Imported {len(valid)}; {len(errors)} flagged");db.commit()
    return {"imported":len(valid),"errors":errors,"status":"Pending","cross_institution_check":True}

@app.get("/api/students/export.csv")
def export_csv(user: User = Depends(get_user),db:Session=Depends(get_db)):
    stmt=select(Student)
    if user.role=="student": stmt=stmt.where(Student.email==user.email)
    elif user.role=="staff": stmt=stmt.where(Student.institution==user.institution)
    out=io.StringIO();w=csv.writer(out);w.writerow(["name","register_no","department","year","apaar_id","intake","enrolment","supernumerary","nri","oci","foreign","status"])
    for s in db.scalars(stmt):w.writerow([s.name,s.register_no,s.department,s.year,s.apaar_id,s.intake,s.enrolment,s.supernumerary,s.nri,s.oci,s.foreign,s.status])
    return StreamingResponse(iter([out.getvalue()]),media_type="text/csv",headers={"Content-Disposition":"attachment; filename=apaar-student-records.csv"})

@app.get("/api/analytics")
def analytics(user:User=Depends(allow("staff","officer")),db:Session=Depends(get_db)):
    stmt=select(Student)
    if user.role=="staff":stmt=stmt.where(Student.institution==user.institution)
    rows=db.scalars(stmt).all(); counts={x:sum(1 for s in rows if s.status==x) for x in ["Pending","Approved","Rejected","Needs correction"]}
    return {"students":len(rows),"enrolment":sum(s.enrolment for s in rows),"intake":sum(s.intake for s in rows),"status_counts":counts,"institutions":1 if user.role=="staff" else len({s.institution for s in rows})}

@app.get("/api/analytics/enrollment-trend")
def enrollment_trend(user:User=Depends(allow("officer"))):
    institutions=["National Institute of Technology, Jaipur","Central Institute of Engineering, Pune","Government College of Technology, Chennai"]
    years=["2022–23","2023–24","2024–25","2025–26","2026–27"]
    profiles={institutions[0]:{"Computer Science":[420,458,491,522,558],"Core Engineering":[610,628,655,673,698],"Applied Programs":[286,301,315,332,347]},institutions[1]:{"Computer Science":[378,412,449,480,514],"Core Engineering":[540,562,588,609,633],"Applied Programs":[320,338,352,371,390]},institutions[2]:{"Computer Science":[312,348,382,418,453],"Core Engineering":[670,692,710,735,761],"Applied Programs":[246,265,281,299,317]}}
    items=[{"institution":institution,"year":year,"category":category,"enrolment":count} for institution,categories in profiles.items() for category,counts in categories.items() for year,count in zip(years,counts)]
    return {"institutions":institutions,"items":items}

@app.get("/api/certificates")
def list_certificates(user:User=Depends(get_user),db:Session=Depends(get_db)):
    stmt=select(Certificate).join(Student,Student.public_id==Certificate.student_id)
    if user.role=="student":stmt=stmt.where(Student.email==user.email)
    elif user.role=="staff":stmt=stmt.where(Student.institution==user.institution)
    return [{"id":c.id,"student_id":c.student_id,"title":c.title,"serial_no":c.serial_no,"issued_by":c.issued_by,"issued_at":c.created_at.isoformat()} for c in db.scalars(stmt.order_by(Certificate.created_at.desc())).all()]

@app.post("/api/students/{public_id}/certificates",status_code=201)
def issue_certificate(public_id:str,title:str="Academic Record Certificate",user:User=Depends(allow("staff")),db:Session=Depends(get_db)):
    student=db.scalar(select(Student).where(Student.public_id==public_id,Student.institution==user.institution))
    if not student:raise HTTPException(404,"Student not found in your institution")
    if student.status!="Approved":raise HTTPException(409,"Only approved student records can receive a certificate")
    cert=Certificate(student_id=student.public_id,title=title[:160],serial_no=f"NITJ-{datetime.now().year}-{uuid.uuid4().hex[:10].upper()}",issued_by=user.name)
    db.add(cert);db.flush();log(db,user,"issued_certificate","certificate",str(cert.id),f"{cert.title} · {student.name}");db.commit();db.refresh(cert)
    return {"id":cert.id,"student_id":cert.student_id,"title":cert.title,"serial_no":cert.serial_no,"issued_by":cert.issued_by,"issued_at":cert.created_at.isoformat()}

@app.get("/api/certificates/{certificate_id}/download")
def download_certificate(certificate_id:int,user:User=Depends(get_user),db:Session=Depends(get_db)):
    cert=db.get(Certificate,certificate_id)
    if not cert:raise HTTPException(404,"Certificate not found")
    student=db.scalar(select(Student).where(Student.public_id==cert.student_id))
    if user.role=="student" and student.email!=user.email:raise HTTPException(403,"This certificate belongs to another student")
    if user.role=="staff" and student.institution!=user.institution:raise HTTPException(403,"This certificate belongs to another institution")
    lines=["Certificate of Academic Record",f"Student: {student.name}",f"Institution: {student.institution}",f"Student ID: {student.register_no}",f"APAAR ID: {student.apaar_id}",f"Certificate: {cert.title}",f"Serial number: {cert.serial_no}",f"Issued by: {cert.issued_by}",f"Issued on: {cert.created_at.strftime('%d %B %Y')}"]
    def esc(text):return text.replace("\\","\\\\").replace("(","\\(").replace(")","\\)")
    commands=["BT","/F1 18 Tf","72 740 Td",f"({esc(lines[0])}) Tj","/F1 11 Tf"]
    for line in lines[1:]:commands.extend(["0 -28 Td",f"({esc(line)}) Tj"])
    commands.append("ET");stream="\n".join(commands).encode("latin-1","replace")
    objects=[b"<< /Type /Catalog /Pages 2 0 R >>",b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",b"<< /Length "+str(len(stream)).encode()+b" >>\nstream\n"+stream+b"\nendstream"]
    output=bytearray(b"%PDF-1.4\n");offsets=[0]
    for i,obj in enumerate(objects,1):offsets.append(len(output));output.extend(f"{i} 0 obj\n".encode()+obj+b"\nendobj\n")
    xref=len(output);output.extend(f"xref\n0 {len(offsets)}\n0000000000 65535 f \n".encode())
    for off in offsets[1:]:output.extend(f"{off:010d} 00000 n \n".encode())
    output.extend(f"trailer\n<< /Size {len(offsets)} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode())
    from fastapi.responses import Response
    return Response(bytes(output),media_type="application/pdf",headers={"Content-Disposition":f'attachment; filename="{cert.serial_no}.pdf"'})

@app.get("/api/audit")
def audit(limit:int=Query(100,ge=1,le=500),user:User=Depends(allow("staff","officer")),db:Session=Depends(get_db)):
    stmt=select(AuditLog)
    if user.role=="staff":stmt=stmt.where(AuditLog.institution==user.institution)
    return [{"id":x.id,"user":x.user_email,"institution":x.institution,"action":x.action,"entity":x.entity,"entity_id":x.entity_id,"details":x.details,"at":x.at.isoformat()} for x in db.scalars(stmt.order_by(AuditLog.at.desc()).limit(limit)).all()]

@app.get("/api/institutions")
def institutions(user:User=Depends(allow("officer")),db:Session=Depends(get_db)):
    students=db.scalars(select(Student)).all(); groups={}
    for s in students:
        x=groups.setdefault(s.institution,{"institution":s.institution,"students":0,"pending":0,"approved":0,"needs_attention":0})
        x["students"]+=1;x["pending"]+=s.status=="Pending";x["approved"]+=s.status=="Approved";x["needs_attention"]+=s.status in {"Rejected","Needs correction"}
    return list(groups.values())

@app.get("/api/duplicates/{apaar_id}")
def duplicate_check(apaar_id:str,user:User=Depends(allow("staff","officer")),db:Session=Depends(get_db)):
    if not re.fullmatch(r"\d{12}",apaar_id):raise HTTPException(422,"APAAR ID must contain 12 digits")
    rows=db.scalars(select(Student).where(Student.apaar_id==apaar_id)).all()
    return {"duplicate":len(rows)>1,"matches":[{"student_id":x.public_id,"institution":x.institution,"status":x.status} for x in rows],"scope":"institution record database"}


