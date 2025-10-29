import pymupdf  
import re

if __name__ == "__main__":
    read_pdf = pymupdf.open("D://automation//test data//ThirtyDaySummary_638966391093975385.pdf")
    
    page = read_pdf[0]
    
    start = page.search_for("Compassionate Home Health LTD")
    startTitle = page.search_for("30 DAY SUMMARY/CASE")
    
    end = page.search_for("Homebound Status")
    
    rx0 = start[0].x0
    ry0 = start[0].y0
    rx1 = startTitle[0].x1
    ry1 = end[0].y1
    
    cr = pymupdf.Rect(rx0,ry0,rx1,ry1)
    
    text = page.get_text(clip=cr, sort=True)
    
    print(text)
    
    #text to find : Patient Name, Patient DOB, Physician Name, Physician Phone, Physician Fax
    
    #capture Patient Name:
    patientName = re.search(
        r'(?<=Patient Name:)\s+\w+,\s\w*\s[A-Z]?',
        text
    )
    if(patientName):
        print("Patient Name: " + patientName.group(0).strip())
    else:
        print("Patient Name not found")
        
    #capture Patient DOB:
    patientDOB = re.search(
        r'(?<=DOB:)\s+[0-9]+\/[0-9]+\/[0-9]+',
        text
    )
    if(patientDOB):
        print("Patient DOB: " + patientDOB.group(0).strip())
    else:
        print("Patient DOB not found")
    
    #capture Physician Name:
    physicianName = re.search(
        r'(?<=Physician:)\s+\w+, ?\w+',
        text
    )
    if(physicianName):
        print("Physician Name: " + physicianName.group(0).strip())
    else:
        print("Physician Name not found")
        
    #capture Physician Phone:
    physicianPhone = re.search(
        r'(?<=Physician Phone:)\s+\([0-9]+\) [0-9]+-[0-9]+',
        text
    )
    if(physicianPhone):
        print("Physician Phone: " + physicianPhone.group(0).strip())
    else:
        print("Physician Phone not found")
        
    #capture Physician Fax:
    physicianFax = re.search(
        r'(?<=Physician Fax:)\s+\([0-9]+\) [0-9]+-[0-9]+',
        text
    )
    if(physicianFax):
        print("Physician Fax: " + physicianFax.group(0).strip())
    else:
        print("Physician Fax not found")
        
    #Capture Date on bottom, Date today, Number of pages + 
    start = page.search_for("Signature:")
    startTitle = page.search_for("30 DAY SUMMARY/CASE")
    
    end = page.search_for("Axxess")
    
    rx0 = start[0].x0
    ry0 = start[0].y0
    rx1 = startTitle[0].x1
    ry1 = end[0].y1
    
    cr = pymupdf.Rect(rx0,ry0,rx1,ry1)
    
    text = page.get_text(clip=cr, sort=True)

    #capture date on bottom:
    dateBottom = re.search(
        r'(?<=PT)\s+[0-9]+\/[0-9]+\/[0-9]+',
        text
    )
    if(dateBottom):
        print("Bottom date: " + dateBottom.group(0).strip())
    else:
        print("Date on bottom not found")