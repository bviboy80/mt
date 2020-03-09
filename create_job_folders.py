"""
Create job folders for MTA Benefit Confirmation/FMLA

User provides the job number and job type

Script will 
1) Create the necessary folders.
2) Copy the appropriate print coversheet, envelope admark 
   layout and counts excel spreadsheet from the project 
   folder to the job folder.
3) Create the PrintNet paths text file if job type is Benefit Confirmation.


Usage: python createFolders.py 
(or just double click script if Python is the default for opening ".py" files)

v1.0
Create by Shaun Thomas
10/12/2018

"""


import sys
import os
import shutil


def main():
    # Number choices and the corresponding Job type name and print coversheet     
    jobTypeDict = {
        "1" : ("MTA Benefits Confirm", "coversheet_bene.docx"),
        "2" : ("MTA FMLA Certification", "coversheet_cert.docx"),
        "3" : ("MTA FMLA Designation", "coversheet_des.docx"),
        "4" : ("MTA FMLA Expiration", "coversheet_exp.docx")
        }
        
    jobNumber = inputJobNumber()
    number_choice = selectJobTypeNumber()

    jobType = jobTypeDict[number_choice][0]
    jobCoversheet = jobTypeDict[number_choice][1]
    
    jobFolder = os.path.join(r'P:\Vanguard', "{} {}".format(jobNumber, jobType))
        
    createJobFolders(jobFolder, jobType) 
    copyCoverSheetsToJobFolder(jobFolder, jobCoversheet)

    
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


def createJobFolders(jobFolder, jobType):
    """ Create job directories """
    if jobType == "MTA Benefits Confirm":
        os.makedirs(os.path.join(jobFolder, "Documents", "combined"))   
        os.mkdir(os.path.join(jobFolder, "Documents", "counts"))
        os.mkdir(os.path.join(jobFolder, "Print"))
    else:
        os.makedirs(os.path.join(jobFolder, "Documents"))
        os.mkdir(os.path.join(jobFolder, "Print"))
        

def copyCoverSheetsToJobFolder(jobFolder, jobCoversheet):
    """ Copy the counts excel, env layout and cover sheet to the job folder.  """   
                
    shutil.copy2(os.path.join(r'P:\Vanguard\MTA BeneConfirm_FMLA', "envelope layout.pptx"),
                os.path.join(jobFolder, "envelope layout.pptx"))
                
    shutil.copy2(os.path.join(r'P:\Vanguard\MTA BeneConfirm_FMLA', jobCoversheet),
                os.path.join(jobFolder, jobCoversheet))    


if __name__ == '__main__':
    main()
