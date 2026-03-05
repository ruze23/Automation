import pymupdf
import re
from datetime import date, timedelta
from pathlib import Path
from docxtpl import DocxTemplate
from docx2pdf import convert

PATTERN_CASE = r'(\d{2}\s?(-)?DAY)\s+(SUMMARY\/CASE)'
PATTERN_CONFERENCE = r'(CONFERENCE)'
PATTERN_PT = r'(\s-\sPT)'
PATTERN_PATIENT_DOB = r'(?<=DOB:)\s+[0-9]+\/[0-9]+\/[0-9]+'

def extract(match: re.Match) -> str:
    return match.group(0).strip() if match else ""


def extract_by_pattern(text: str, pattern: str, flags: int = 0) -> str:
    return extract(re.search(pattern, text, flags))


def extract_case_info(
    text: str,
    case_pattern: str = PATTERN_CASE,
    conference_pattern: str = PATTERN_CONFERENCE,
    pt_pattern: str = PATTERN_PT,
) -> tuple[str, str, str]:
    case = extract_by_pattern(text, case_pattern).strip()
    conference = extract_by_pattern(text, conference_pattern).strip()
    pt = extract_by_pattern(text, pt_pattern)
    return case, conference, pt


def extract_patient_name(text: str, pattern: str) -> str:
    return extract_by_pattern(text, pattern).upper()


def extract_patient_dob(text: str, pattern: str = PATTERN_PATIENT_DOB) -> str:
    return extract_by_pattern(text, pattern)


def extract_physician_name(text: str, pattern: str) -> str:
    return extract_by_pattern(text, pattern).upper()


def extract_physician_phone(text: str, pattern: str) -> str:
    return extract_by_pattern(text, pattern)


def extract_physician_fax(text: str, pattern: str) -> str:
    return extract_by_pattern(text, pattern)


def extract_bottom_date(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def extract_bottom_date1(text: str, pattern: str) -> str:
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(0).strip() if match else ""


def extract_surname(full_name: str) -> str:
    return full_name.split(",")[0].upper() if full_name else ""

def capture_text(page, start: list, startTitle: list, end: list) -> str:
    
    rx0 = start[0].x0
    ry0 = start[0].y0
    rx1 = startTitle[0].x1
    ry1 = end[0].y1
    
    cr = pymupdf.Rect(rx0, ry0, rx1, ry1)
    return page.get_text(clip=cr, sort=True)


def build_and_send_signature_request(
    template_path: Path,
    context: dict,
    output_dir: Path,
    sending_dir: Path,
    file_name: str,
    original_pdf_path: Path,
) -> bool:
    doc = DocxTemplate(str(template_path))

    try:
        doc.render(context)
    except Exception as e:
        print("Render error:", e)
        raise

    sending_dir.mkdir(parents=True, exist_ok=True)

    output_file = output_dir / f"{file_name}_signature_request.docx"
    doc.save(str(output_file))

    pdf_output_file = output_dir / f"{file_name}_signature_request.pdf"
    try:
        convert(str(output_file), str(pdf_output_file))
        print(f"PDF generated: {pdf_output_file}")
    except Exception as e:
        print("PDF conversion failed:", e)
        return False

    with pymupdf.open(str(pdf_output_file)) as doc_a, pymupdf.open(str(original_pdf_path)) as doc_b:
        doc_a.insert_file(doc_b)
        doc_a.save(str(sending_dir / f"{file_name}.pdf"))

    #output_file.unlink(missing_ok=True)
    pdf_output_file.unlink(missing_ok=True)
    return True


def process_case_summary_files(
    pdf_path: Path,
    output_dir: Path,
    template_path: Path,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Processing summary case: {pdf_path.name}")

    if "Sixty" in pdf_path.name:
        title = "60 DAY SUMMARY/CASE"
    else:
        title = "30 DAY SUMMARY/CASE"
        
    PATTERN_PATIENT_NAME = r'(?<=Patient Name:)\s+\w+,\s\w*\s([A-Z])?'
    PATTERN_PHYSICIAN_NAME = r'(?<=Physician:)(\s+\w*,?\s\w*\s\w*\W?\w?\W)(?!DNR:)'
    PATTERN_PHYSICIAN_PHONE = r'(?<=Physician Phone:)\s+\([0-9]+\) [0-9]+-[0-9]+'
    PATTERN_PHYSICIAN_FAX = r'(?<=Physician Fax:)\s+\([0-9]+\) [0-9]+-[0-9]+'
    PATTERN_BOTTOM_DATE = r'\b(?:RN|PT)\b\s+([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})'
    
    with pymupdf.open(str(pdf_path)) as read_pdf:
        page = read_pdf[0]

        pages = read_pdf.page_count + 2

        start = page.search_for("Compassionate Home Health LTD")
        startTitle = page.search_for(title)

        end = page.search_for("Homebound Status")
        
        try:
            text = capture_text(page, start, startTitle, end)
        except Exception as e:
            print("Error capturing text:", e)
            if "Sixty" in pdf_path.name:
                startTitle = page.search_for("60-DAY SUMMARY/CASE")
                text = capture_text(page, start, startTitle, end)
            else:
                startTitle = page.search_for("30-DAY SUMMARY/CASE")
                text = capture_text(page, start, startTitle, end)    

        case, conference, pt = extract_case_info(text)
        
        patientName = extract_patient_name(text, PATTERN_PATIENT_NAME)
        patientSurname = extract_surname(patientName)
        print("Patient Name:" if patientName else "Patient Name not found", patientName)

        patientDOB = extract_patient_dob(text)
        print("Patient DOB:" if patientDOB else "Patient DOB not found", patientDOB)

        physicianName = extract_physician_name(text, PATTERN_PHYSICIAN_NAME)
        
        physicianSurname = extract_surname(physicianName)
        print("Physician Name:" if physicianName else "Physician Name not found", physicianName)

        physicianPhone = extract_physician_phone(text, PATTERN_PHYSICIAN_PHONE)
        print("Physician Phone:" if physicianPhone else "Physician Phone not found", physicianPhone)

        physicianFax = extract_physician_fax(text, PATTERN_PHYSICIAN_FAX)
        print("Physician Fax:" if physicianFax else "Physician Fax not found", physicianFax)

        fileType = f"{case} {conference} {pt}".strip()

        start = page.search_for("Signature:")

        end = page.search_for("Axxess")

        text = capture_text(page, start, startTitle, end)

        dateBottom = extract_bottom_date(text, PATTERN_BOTTOM_DATE)
        print("Bottom date:" if dateBottom else "Date on bottom not found", dateBottom)
        
        patientNameForFile = extract_by_pattern(patientName,r'(\w+,\s\w*)')
        fileName = f"{patientNameForFile} {case[0:11]} {pt[2:]} {dateBottom.replace('/','-')}".strip()

        context = {
            'physician_name': physicianName,
            'physician_fax': physicianFax,
            'physician_phone': physicianPhone,
            'no_pages': pages,
            'patient_name': patientName,
            'patient_dob': patientDOB,
            'date_today': (date.today() - timedelta(days=1)).strftime('%m/%d/%Y'),
            'physician_surname': physicianSurname,
            'file_type': fileType,
            'date_bottom': dateBottom
        }

        sending_dir = Path(f"D://automation//output//{patientSurname}//")
        if not build_and_send_signature_request(
            template_path=template_path,
            context=context,
            output_dir=output_dir,
            sending_dir=sending_dir,
            file_name=fileName,
            original_pdf_path=pdf_path,
        ):
            return

def process_physician_order(pdf_path: Path, 
                            output_dir: Path,
                            template_path: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Processing file name: {pdf_path.name}")
    
    with pymupdf.open(str(pdf_path)) as read_pdf:
        page = read_pdf[0]
        pages = read_pdf.page_count + 2
        
        start = page.search_for("Patient:")
        startTitle = page.search_for("Physician Order")
        end = page.search_for("Mbi:")
        
        text = capture_text(page, start, startTitle, end) 
        
        fileType = "PHYSICIAN ORDER"
        
        PATTERN_PATIENT_NAME = r'(?<=Patient:)\s+\w+,\s\w*\s([A-Z])?'
        PATTERN_PHYSICIAN_NAME = r'(?<=Practitioner:)(\s+\w*,?\s\w*)'
        PATTERN_PHYSICIAN_PHONE = r'(?<=Phone:)\s+\([0-9]+\) [0-9]+-[0-9]+'
        PATTERN_PHYSICIAN_FAX = r'(?<=Fax:)\s+\([0-9]+\) [0-9]+-[0-9]+'
        PATTERN_DATE_BOTTOM = r'(?<=Date:\s)\d{2}\/\d{2}\/\d{4}'
        
        patientName = extract_patient_name(text, PATTERN_PATIENT_NAME)
        patientSurname = extract_surname(patientName)
        
        patientDOB = extract_patient_dob(text)
        
        physicianName = extract_physician_name(text, PATTERN_PHYSICIAN_NAME)
        
        if "M.D." in text:
            physicianName = f'{physicianName} M.D.'
        physicianSurname = extract_surname(physicianName)
        
        physicianPhone = extract_physician_phone(text, PATTERN_PHYSICIAN_PHONE)
        physicianFax = extract_physician_fax(text, PATTERN_PHYSICIAN_FAX)
        
        start = page.search_for("Order Date:")
        end = page.search_for("Summary:")
        
        text = capture_text(page, start, startTitle, end)
        
        dateBottom = extract_bottom_date1(text, PATTERN_DATE_BOTTOM)
        
        patientNameForFile = extract_by_pattern(patientName,r'(\w+,\s\w*)')
        fileName = f"{patientNameForFile} PO {dateBottom.replace('/','-')}".strip()
        
        context = {
            'physician_name': physicianName,
            'physician_fax': physicianFax,
            'physician_phone': physicianPhone,
            'no_pages': pages,
            'patient_name': patientName,
            'patient_dob': patientDOB,
            'date_today': date.today().strftime('%m/%d/%Y'),
            'physician_surname': physicianSurname,
            'file_type': fileType,
            'date_bottom': dateBottom
        }

        sending_dir = Path(f"D://automation//output//{patientSurname}//")
        if not build_and_send_signature_request(
            template_path=template_path,
            context=context,
            output_dir=output_dir,
            sending_dir=sending_dir,
            file_name=fileName,
            original_pdf_path=pdf_path,
        ):
            return

def process_therapy_of_Care(pdf_path: Path, 
                            output_dir: Path,
                            template_path: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Processing file name: {pdf_path.name}")
    
    with pymupdf.open(str(pdf_path)) as read_pdf:
        page = read_pdf[0]
        pages = read_pdf.page_count + 2
        
        start = page.search_for("Patient Name")
        startTitle = page.search_for("Physical Therapy Plan of Care with Full Evaluation")
        end = page.search_for("Advance Directives")
        
        text = capture_text(page, start, startTitle, end) 
        
        fileType = "PT Plan of Care".upper()
        
        PATTERN_PATIENT_NAME = r'(?<=Patient:)\s+\w+,\s\w*\s([A-Z])?'
        PATTERN_PHYSICIAN_NAME = r'(?<=Practitioner:)(\s+\w*,?\s\w*)'
        PATTERN_PHYSICIAN_PHONE = r'(?<=Phone:)\s+\([0-9]+\) [0-9]+-[0-9]+'
        PATTERN_PHYSICIAN_FAX = r'(?<=Fax:)\s+\([0-9]+\) [0-9]+-[0-9]+'
        PATTERN_DATE_BOTTOM = r'(?<=Date:\s)\d{2}\/\d{2}\/\d{4}'
        
        patientName = extract_patient_name(text, PATTERN_PATIENT_NAME)
        patientSurname = extract_surname(patientName)
        
        patientDOB = extract_patient_dob(text)
        
        physicianName = extract_physician_name(text, PATTERN_PHYSICIAN_NAME)
        
        if "M.D." in text:
            physicianName = f'{physicianName} M.D.'
        physicianSurname = extract_surname(physicianName)
        
        physicianPhone = extract_physician_phone(text, PATTERN_PHYSICIAN_PHONE)
        physicianFax = extract_physician_fax(text, PATTERN_PHYSICIAN_FAX)
        
        start = page.search_for("Order Date:")
        end = page.search_for("Summary:")
        
        text = capture_text(page, start, startTitle, end)
        
        dateBottom = extract_bottom_date1(text, PATTERN_DATE_BOTTOM)
        
        patientNameForFile = extract_by_pattern(patientName,r'(\w+,\s\w*)')
        fileName = f"{patientNameForFile} PO {dateBottom.replace('/','-')}".strip()
        
        context = {
            'physician_name': physicianName,
            'physician_fax': physicianFax,
            'physician_phone': physicianPhone,
            'no_pages': pages,
            'patient_name': patientName,
            'patient_dob': patientDOB,
            'date_today': date.today().strftime('%m/%d/%Y'),
            'physician_surname': physicianSurname,
            'file_type': fileType,
            'date_bottom': dateBottom
        }

        sending_dir = Path(f"D://automation//output//{patientSurname}//")
        if not build_and_send_signature_request(
            template_path=template_path,
            context=context,
            output_dir=output_dir,
            sending_dir=sending_dir,
            file_name=fileName,
            original_pdf_path=pdf_path,
        ):
            return
        
if __name__ == "__main__":
    input_dir = Path("D://automation//real data//")
    output_dir = Path("D://automation//output//")
    all_pdf_files = sorted(input_dir.glob("*.pdf"))

    if not all_pdf_files:
        print(f"No PDF files found in: {input_dir}")
        raise SystemExit(0)

    for pdf_path in all_pdf_files:
        if "DaySummary" in pdf_path.name:
            process_case_summary_files(
                pdf_path=pdf_path,
                output_dir=output_dir,
                template_path=Path("D://automation//templates//30_60 DAY CASE SUMMARY SIGNATURE REQUEST.docx"),
            )
        elif "PhysicianOrder" in pdf_path.name:
            process_physician_order(pdf_path=pdf_path, 
                                    output_dir=output_dir,
                                    template_path=Path("D://automation//templates//other signature request.docx"),
            )
        else:
            print('Not yet handled, skipping')  
            
            