import pymupdf
import re
from datetime import date
from pathlib import Path
from zipfile import ZipFile
from docxtpl import DocxTemplate
from docx2pdf import convert

if __name__ == "__main__":
    input_dir = Path("D://automation//real data//")

    matching_files = sorted(
        file_path for file_path in input_dir.glob("*.pdf")
    )

    if not matching_files:
        print(f"No matching PDF files found in: {input_dir}")
        raise SystemExit(0)

    def extract(match: re.Match) -> str:
        return match.group(0).strip() if match else ""

    output_dir = Path("D://automation//output//")
    sending_dir = Path("D://automation//output//for sending//")
    output_dir.mkdir(parents=True, exist_ok=True)
    sending_dir.mkdir(parents=True, exist_ok=True)

    for pdf_path in matching_files:
        print(f"Processing: {pdf_path.name}")
        
        if(pdf_path.name.__contains__("SixtyDay")):
            title = "60 DAY SUMMARY/CASE"
        else:
            title = "30 DAY SUMMARY/CASE"

        with pymupdf.open(str(pdf_path)) as read_pdf:
            page = read_pdf[0]

            pages = read_pdf.page_count + 2

            start = page.search_for("Compassionate Home Health LTD")
            startTitle = page.search_for(title)

            end = page.search_for("Homebound Status")

            rx0 = start[0].x0
            ry0 = start[0].y0
            rx1 = startTitle[0].x1
            ry1 = end[0].y1

            cr = pymupdf.Rect(rx0, ry0, rx1, ry1)

            text = page.get_text(clip=cr, sort=True)
            
            #capture case
            case_match = re.search(
                r'(\d{2} DAY)\s+(SUMM)',
                text
            )
            conference_match = re.search(
                r'(?:\s*-\s*PT)',
                text
            )
            case = extract(case_match).strip()
            conference = extract(conference_match).strip()
            
            # capture Patient Name:
            patientName_match = re.search(
                r'(?<=Patient Name:)\s+\w+,\s\w*\s[A-Z]?',
                text
            )
            patientName = extract(patientName_match).upper()
            patientSurname = patientName.split(",")[0].upper() if patientName else ""
            print("Patient Name:" if patientName else "Patient Name not found", patientName)
            
            # capture Patient DOB:
            patientDOB_match = re.search(
                r'(?<=DOB:)\s+[0-9]+\/[0-9]+\/[0-9]+',
                text
            )
            patientDOB = extract(patientDOB_match)
            print("Patient DOB:" if patientDOB else "Patient DOB not found", patientDOB)

            # capture Physician Name:
            physicianName_match = re.search(
                r'(?<=Physician:)\s+\w+, ?\w+',
                text
            )
            physicianName = extract(physicianName_match).upper()
            physicianSurname = physicianName.split(",")[0].upper() if physicianName else ""
            print("Physician Name:" if physicianName else "Physician Name not found", physicianName)

            # capture Physician Phone:
            physicianPhone_match = re.search(
                r'(?<=Physician Phone:)\s+\([0-9]+\) [0-9]+-[0-9]+',
                text
            )
            physicianPhone = extract(physicianPhone_match)
            print("Physician Phone:" if physicianPhone else "Physician Phone not found", physicianPhone)

            # capture Physician Fax:
            physicianFax_match = re.search(
                r'(?<=Physician Fax:)\s+\([0-9]+\) [0-9]+-[0-9]+',
                text
            )
            physicianFax = extract(physicianFax_match)
            print("Physician Fax:" if physicianFax else "Physician Fax not found", physicianFax)
            
            fileName = f"{patientSurname} {case} {conference}"

            # Capture Date on bottom, Date today, Number of pages +
            start = page.search_for("Signature:")
            startTitle = page.search_for(title)

            end = page.search_for("Axxess")

            rx0 = start[0].x0
            ry0 = start[0].y0
            rx1 = startTitle[0].x1
            ry1 = end[0].y1

            cr = pymupdf.Rect(rx0, ry0, rx1, ry1)

            text = page.get_text(clip=cr, sort=True)

            # capture date on bottom:
            dateBottom_match = re.search(
                r'\b(?:RN|PT)\b\s+([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})',
                text,
                re.IGNORECASE
            )
            dateBottom = dateBottom_match.group(1).strip() if dateBottom_match else ""
            print("Bottom date:" if dateBottom else "Date on bottom not found", dateBottom)

            doc = DocxTemplate("D://automation//templates//30_60 DAY CASE SUMMARY SIGNATURE REQUEST.docx")

            context = {
                'physician_name': physicianName,
                'physician_fax': physicianFax,
                'physician_phone': physicianPhone,
                'no_pages': pages,
                'patient_name': patientName,
                'patient_dob': patientDOB,
                'date_bottom': date.today().strftime('%m/%d/%Y'),
                'physician_surname': physicianSurname,
                'file_name': fileName,
                'date_file': dateBottom 
            }

            try:
                doc.render(context)
            except Exception as e:
                print("Render error:", e)
                raise

            output_file = output_dir / f"{fileName}_signature_request.docx"
            doc.save(str(output_file))

            pdf_output_file = output_dir / f"{fileName}_signature_request.pdf"
            try:
                convert(str(output_file), str(pdf_output_file))
                print(f"PDF generated: {pdf_output_file}")
            except Exception as e:
                print("PDF conversion failed:", e)
                
            with pymupdf.open(str(pdf_output_file)) as doc_a,pymupdf.open(str(pdf_path)) as doc_b:
                doc_a.insert_file(doc_b)
                doc_a.save(f'{sending_dir}/{patientName} {conference}.pdf') 

            # Basic validation of generated docx container
            try:
                size = output_file.stat().st_size
                print(f"Generated file size: {size} bytes")
                with ZipFile(output_file) as z:
                    names = set(z.namelist())
                    required = {"[Content_Types].xml", "word/document.xml"}
                    missing = required - names
                    if missing:
                        print(f"Missing required parts: {missing}")
                    else:
                        print("Docx zip structure looks intact.")
            except Exception as e:
                print("Docx validation failed:", e)

            print(f"Document generated: {output_file}")
            
            