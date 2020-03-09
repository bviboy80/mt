"""
Sort PDFs by letter type for MTA Benefit Confirmation/FMLA

User provides the job number and job type

Script will 


v1.0
Create by Shaun Thomas
10/12/2018

"""


import sys
import os
import shutil
import re


def main():
    # Number choices and the corresponding Job type name and print coversheet     
    
    zip_pdf_folder = r'P:\Vanguard\MTA BeneConfirm_FMLA\zips'
    letters_dict = createLtrFolders(zip_pdf_folder)   
    pdf_file_list = make_pdfs_to_process_list(zip_pdf_folder)

    for pdf in pdf_file_list:
        zip_pdf_path = os.path.join(zip_pdf_folder, pdf)
        
        for ltr in letters_dict.keys():
            if letters_dict[ltr]["pattern"].match(pdf) != None:
                new_pdf_path = os.path.join(letters_dict[ltr]["folder"], pdf)
                shutil.move(zip_pdf_path, new_pdf_path)
                break
            else:
                continue


def createLtrFolders(zip_pdf_folder):
    
    letter_types = {
                    "benefit" : {"pattern" : re.compile(r'^bas005.+$'),
                                 "folder" : os.path.join(zip_pdf_folder, "benefit")},
                                 
                    "certification" : {"pattern" : re.compile(r'^B_FMLA_CERT.+$'),
                                       "folder" : os.path.join(zip_pdf_folder, "certification")},
                                 
                    "design" : {"pattern" : re.compile(r'^B_FMLA_DSGN.+$'),
                                 "folder" : os.path.join(zip_pdf_folder, "design")}, 
                                 
                    "expired": {"pattern" : re.compile(r'^B_FMLA_EXPRD.+$'),
                                "folder" : os.path.join(zip_pdf_folder, "expired")}
                   }
    
    for k in letter_types.keys():
        ltr_folder = letter_types[k]["folder"]
        if not os.path.exists(ltr_folder):
            os.mkdir(ltr_folder)
            
    return letter_types

    
def make_pdfs_to_process_list(zip_pdf_folder):
    pdfs_to_process_list = []
    for f in os.listdir(zip_pdf_folder):
        file_ext = f.split(".")[-1]
        if file_ext.upper() == "PDF":
            pdfs_to_process_list.append(f)
            
    return pdfs_to_process_list
    


if __name__ == '__main__':
    main()
