import os
import sys
import PyPDF2
import collections
import re
import csv
import subprocess
import openpyxl




jobNumber = input("\r\nProvide DS job number:  ")
while type(jobNumber) != int:
    jobNumber = input("Please provide a number:  ")
while len(str(jobNumber)) != 5:
    jobNumber = input("Please provide a number with the correct number of characters:  ")  


jobFolder = os.path.join(r'P:\Vanguard', "{} MTA Benefits Confirm".format(jobNumber))

inputPDF = os.path.join(jobFolder, "Documents", "combined", "combined.pdf")
pg_list = os.path.join(jobFolder, "Documents", "combined", "record_pages_list.dat")

printOutput = os.path.join(jobFolder, "Print", "{} MTA Benefits Confirm Letters_%h06 Sheets.%e".format(jobNumber))
sheetCounts = os.path.join(jobFolder, "Documents", "counts", "Sheet and Impression Counts.csv")
logFile = os.path.join(jobFolder, "Documents", "counts", "{} job.log".format(jobNumber))

with open(inputPDF, 'rb') as a:
    with open(pg_list, 'wb') as o:
        
        ## Search pattern for text only on the first page
        search_title = re.compile("Confirmation of Benefits Elections")
        search_page = re.compile("1\n\sof\n")
        
        ## Dictionary of records and pages per record.
        record_pages_dict = collections.defaultdict(list)
        
        document = PyPDF2.PdfFileReader(a)
        totalpages = document.getNumPages()
        
        record_sequence = 0
        current_page_count = 0
        page_in_record = 1
        
        for i in range(totalpages):
            current_page_count = i+1
            pageObj = document.getPage(i)
            pageText = pageObj.extractText()
            
            ## Find the first page for each record, then count the page for each record
            ## Use record sequence as key for each record. Use PDF page counter as values.
            if search_title.findall(pageText) and search_page.findall(pageText):
                record_sequence += 1
                page_in_record = 1
                record_pages_dict[record_sequence].append((current_page_count, page_in_record))
            else:
                page_in_record += 1
                record_pages_dict[record_sequence].append((current_page_count, page_in_record))
            
            if current_page_count % 500 == 0:
                print "{} pages processed.".format(current_page_count)
        
        csvOut = csv.writer(o, quoting=csv.QUOTE_ALL)
        csvOut.writerow(["Record Number", "PDF Page Number", "Page In Record", "Number of Impressions"])
        
        ## Loop through each record. 
        ## Write the record number, page number and pages per record to file
        for record in sorted(record_pages_dict.keys()):
            record_page_list = record_pages_dict[record]
            num_of_impressions = len(record_page_list)
            
            for pdf_pg, rec_pg in record_page_list:
                csvOut.writerow([record, pdf_pg, rec_pg, num_of_impressions])
                
                
                
## Create print files

subprocess.call(["G:\PrintNet T Designer\PNetTC.exe",
                 "P:\Vanguard\MTA BeneConfirm_FMLA\PrintNet\PDF_by_Pg_Counts.wfd",
                 "-difPDF_Data", pg_list,
                 "-printDuplexFinishing", "True", 
                 "-PDFPDF_Param", inputPDF, 
                 "-o", "Print", 
                 "-c", "P:\Vanguard\MTA BeneConfirm_FMLA\PrintNet\MTA_BENE.job",
                 "-pc", "OCE",
                 "-dc", "MTA_BENE",
                 "-f", printOutput,
                 "-e", "AdobePostScript3",
                 "-la", logFile,
                 "-splitbygroup"])
                 
## Create Counts csv

subprocess.call(["G:\PrintNet T Designer\PNetTC.exe",
                 "P:\Vanguard\MTA BeneConfirm_FMLA\PrintNet\PDF_by_Pg_Counts.wfd",
                 "-difPDF_Data", pg_list,
                 "-printDuplexFinishing", "True", 
                 "-PDFPDF_Param", inputPDF, 
                 "-o", "SheetsCounts", 
                 "-f", sheetCounts,
                 "-la", logFile]) 
                 

                 
## Read contents of Counts csv and write them to Excel.

with open(sheetCounts, 'rb') as counts_csv:
    csvInCounts = csv.reader(counts_csv, quoting=csv.QUOTE_ALL)
    dataToWrite = [row for row in csvInCounts]
    
    header = dataToWrite[0]
    # NOTE: header is ["Impressions_per_Record", "Sheets_per_Record", "Records_Per_Group", "Impressions_Per_Group"]
    
    # Add up the record and impression counts and append them to the report. 
    total_records = sum([int(row[header.index("Records_Per_Group")]) for row in dataToWrite[1:]])
    total_impressions = sum([int(row[header.index("Impressions_Per_Group")]) for row in dataToWrite[1:]])
    dataToWrite.append(["TOTALS", "", str(total_records), str(total_impressions)])
    
    # Write counts to Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in dataToWrite:        
        ws.append(row)
        
    xls = os.path.join(jobFolder, "Documents", "counts", "Sheet and Impression Counts.xlsx")
    wb.save(xls)
    
