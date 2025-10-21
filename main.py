import pypdf

if __name__ == "__main__":
    read_pdf = pypdf.PdfReader("D://automation//test data//ThirtyDaySummary_638966391093975385.pdf")
    
    page = read_pdf.pages[0]
    
    print(page.extract_text())