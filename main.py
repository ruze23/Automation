import pymupdf  

if __name__ == "__main__":
    read_pdf = pymupdf.open("D://automation//test data//ThirtyDaySummary_638966391093975385.pdf")
    
    page = read_pdf[0]
    
    start = page.search_for("Compassionate Home Health LTD")
    startTitle = page.search_for("30 DAY SUMMARY/CASE")
    
    end = page.search_for("Homebound Status")
    
    rx0 = start[0].x0
    ry0 = start[0].y0
    rx1 = end[0].x1
    ry1 = end[0].y1
    
    cr = pymupdf.Rect(rx0,ry0,rx1,ry1)
    
    text = page.get_text(clip=cr, sort=True)
    
    print(text)
    
    #text to find : Patient Name, Patient DOB, Physician Name, Physician Phone, Physician Fax, Date on bottom, Date today, Number of pages + 2