'''
Combine PDFs placed in a folder.
Combined PDF intended be printed duplex and sometimes stapled.

1. Count number of pages for each PDF to be merged. Add a blank
   to PDFs with odd page counts.
   
2. Combine the PDFs together.

3. Create a text file containing the following:
   - A list of each PDF and its page range within the combined PDF.
     (i.e., 1-20    Letter.pdf
            21-36   Notice.pdf)
   - The total number of pages in the combined PDF.
   - A comma separated list of the Page Ranges for Print Production
     to copy and paste into the OCE Page Programmer Software. 


Created by Shaun Thomas
Version 2.0
Created 7/26/2017

'''

import os
import sys
import PyPDF2


def main():

    # Number choices and the corresponding Job type name and print coversheet     
    jobTypeDict = {
        "1" : ("MTA Benefits Confirm", "N/A"),
        "2" : ("MTA FMLA Certification", 4),
        "3" : ("MTA FMLA Designation", 2),
        "4" : ("MTA FMLA Expiration", 1)
        }
        
    jobNumber = inputJobNumber()
    number_choice = selectJobTypeNumber()

    jobType = jobTypeDict[number_choice][0]
    pagesPerRecord = jobTypeDict[number_choice][1]
    
    
    jobFolder = os.path.join(r'P:\Vanguard', "{} {}".format(jobNumber, jobType))

    pdfFolder = os.path.abspath(os.path.join(jobFolder, "Documents"))   
    
    countsFile = os.path.join(jobFolder, "Documents", "{}_pdf_counts.txt".format(jobNumber))
    
    combinedPDF = os.path.join(jobFolder, "Print", "{} {} Letters.pdf".format(jobNumber, jobType))
    if jobType == "MTA Benefits Confirm":
        combinedPDF = os.path.join(jobFolder, "Documents", "combined", "combined.pdf".format(jobNumber, jobType))
       
            
    PDFfiles = os.listdir(pdfFolder)
    
    # List of PDFs and page counts 
    pdfDict = {}    
    
    # Combined PDF Merger
    mergedPDF = PyPDF2.PdfFileMerger()   
    
    for file in PDFfiles:
        pdffile = os.path.join(pdfFolder, file)
        if os.path.isfile(pdffile) and file[-4:].upper() == ".PDF":
            
            with open(pdffile, 'rb') as o:
                document = PyPDF2.PdfFileReader(o)
                mergedPDF.append(fileobj=document)
                pdfPageCount = document.getNumPages()
                
                if jobType == "MTA Benefits Confirm":
                    pdfDict[file] = (pdfPageCount, pagesPerRecord)
                else:
                    pdfDict[file] = (pdfPageCount, pdfPageCount/pagesPerRecord)
        else:
            continue
            
    # Create combined PDF
    with open(combinedPDF, 'wb') as pdfout:
        mergedPDF.write(pdfout)
    
    impressionCounts = sum ([n[0] for n in pdfDict.values()])
    recordCounts = sum([n[1] for n in pdfDict.values()]) if jobType != "MTA Benefits Confirm" else "N/A"
    
    # Print list to screen
    print "{} - PDF Files merged:\r\n".format(jobType)
    print "File\t\t\tPg Count\t\t\tRecords (Pg/{})".format(pagesPerRecord)
    
    for pdf in sorted(pdfDict.keys()):
        print "{}:\t\t\t{}\t\t\t{}\r\n".format(pdf, pdfDict[pdf][0], pdfDict[pdf][1])
    
    print "No of Files:  {}".format(len(pdfDict.keys()))
    print "No of Impressions:  {}".format(impressionCounts)
    print "No of Records:  {}".format(recordCounts)
    
    
    # Print counts to file
    with open(countsFile, 'wb') as c:
        c.write("{} - PDF Files merged:\r\n\r\n".format(jobType))
        c.write("File                         Pg Count          Records (Pg/{})\r\n".format(pagesPerRecord))
        
        for pdf in sorted(pdfDict.keys()):
            c.write("{}           {}          {}\r\n".format(pdf, pdfDict[pdf][0], pdfDict[pdf][1]))
        
        c.write("\r\nNo of Files:  {}\r\n".format(len(pdfDict.keys())))
        c.write("No of Impressions:  {}\r\n".format(impressionCounts))
        c.write("No of Records:  {}\r\n".format(recordCounts))


def inputJobNumber():
    """ User input to provide DS job ticket number. """
    
    jobNumber = input("\r\nProvide DS job number:  ")
    while type(jobNumber) != int:
        jobNumber = input("Please provide a number:  ")
    while len(str(jobNumber)) != 5:
        jobNumber = input("Please provide a number with the correct number of characters:  ")
    return jobNumber  
        
        
def selectJobTypeNumber():        
    """ User input to select the MTA job type """
    
    number_choice = str(input("\r\n".join(["\r\nSelect job type:",
        "1 = MTA Benefits Confirm",
        "2 = MTA FMLA Certification",
        "3 = MTA FMLA Designation",
        "4 = MTA FMLA Expiration",
        "-->  "])))
    while number_choice not in ["1", "2", "3", "4"]:
        number_choice = str(input("Please select an appropriate choice:  "))
        
    return number_choice 
         
        
if __name__ == '__main__':
    main()
