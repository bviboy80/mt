import sys
import struct
import csv
import os
import re
import openpyxl
import subprocess
import math


def main():
    excel_file = os.path.abspath(sys.argv[1])
    outdir = os.path.dirname(excel_file)
    
    print "Converting XLS to CSV..."
    csv_file = convertXLStoCSV(outdir, excel_file)

    print "Grouping employee and dependents..."
    records_dict = createRecordsDict(csv_file)
    
    print "Sorting records by number of Dependents..."
    mail_piece_dict = sortByDependentCount(records_dict)
    
    print "Sorting Foreign and Domestic records..."
    sortForeignAndDomestic(mail_piece_dict)
    
    print "Writing records to file..."
    writeRecordsToFile(outdir, mail_piece_dict)
    
    print "Generating counts...."
    getCounts(outdir, mail_piece_dict)
    


def convertXLStoCSV(outdir, excel_file):
    """ CSV file to write Excel output to.
    Create temp VB script which will create the 
    CSV. Can't use tempfile since 'cscript' looks for 
    the file extension of the script.  """
        
    excel_name = os.path.basename(excel_file)
    csv_name = ".".join(excel_name.split(".")[:-1]) + ".csv"
    csv_file = os.path.join(outdir, csv_name)
    
    # Temp XLS to CSV vb script
    temp_vb_script = os.path.join(outdir, "tempXLStoCSV-DO_NOT_TOUCH.vbs")
    
    vb_string = "\n".join(
    ["Dim oExcel",
     "Set oExcel = CreateObject(\"Excel.Application\")",
     "oExcel.DisplayAlerts = False",
     "Dim oBook",
     "Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))",
     "oBook.SaveAs WScript.Arguments.Item(1), 6",
     "oBook.Close False",
     "oExcel.Quit",
     "WScript.Echo \"Done\""
    ])

    with open(temp_vb_script, 'wb') as t:
        t.write(vb_string)
        
    # Process using command line arguments
    subprocess.call(["cscript", temp_vb_script, excel_file, csv_file])
    
    # Delete temp vb script
    os.remove(temp_vb_script)
    
    return csv_file

    

def createRecordsDict(csv_file):
    records_dict = {}
    
    with open(csv_file, 'rb') as inhandle:
    
        csvInfile = csv.reader(inhandle, quoting=csv.QUOTE_ALL)
        csvInfile.next()
        
        hdr = ["First_Name", "Last_Name", "BSC_ID", "Name 2", "Name 3", "Name 4", 
        "Address 1", "Address 2", "City", "ST", "Postal_Code", 
        "Dep First_Name", "Dep Last_Name"]

        
        for row in csvInfile:
            bsc_id = row[hdr.index("BSC_ID")]
            
            records_dict.setdefault(bsc_id, {"employee" : [],
                                             "dependents" : []
                                            })
                                            
            employeeRow = createEmployeeRow(row, hdr)
            dependentRow = createDependentRow(row, hdr)
            
            records_dict[bsc_id]["employee"] = employeeRow
            records_dict[bsc_id]["dependents"].append(dependentRow)
        
    return records_dict    





def createEmployeeRow(row, hdr):

    return [row[hdr.index("First_Name")], 
            row[hdr.index("Last_Name")],
            row[hdr.index("Name 2")],
            row[hdr.index("Name 3")],
            row[hdr.index("Name 4")],
            row[hdr.index("Address 1")],
            row[hdr.index("Address 2")],
            row[hdr.index("City")],
            row[hdr.index("ST")],
            row[hdr.index("Postal_Code")],
            row[hdr.index("BSC_ID")]]

def createDependentRow(row, hdr):
    return [row[hdr.index("BSC_ID")],
            row[hdr.index("Dep First_Name")],
            row[hdr.index("Dep Last_Name")]]
            
            
def sortByDependentCount(records_dict):
    mail_piece_dict = {}
    
    for bsc_id in sorted(records_dict.keys()):
        
        dependentList = records_dict[bsc_id]["dependents"]
        dependentCount = len(dependentList)
        
        mail_piece_dict.setdefault(dependentCount, [])
        mail_piece_dict[dependentCount].append(records_dict[bsc_id])
    
    return mail_piece_dict




        
def sortForeignAndDomestic(mail_piece_dict):

    for dependentCount in sorted(mail_piece_dict.keys()):
        
        foreign_records = []
        domestic_records = []
        
        mail_records = mail_piece_dict[dependentCount]
        
        for record_dict in mail_records:
            record_state = record_dict["employee"][8]
            
            if record_state == "":
                foreign_records.append(record_dict)
            else:    
                domestic_records.append(record_dict)
        
        mail_piece_dict[dependentCount] = domestic_records + foreign_records
        
      
def writeRecordsToFile(outdir, mail_piece_dict):
        
    employee_hdr = ["IM barcode Digits", "OEL", 
                    "Sack and Pack Numbers", "Presort Sequence",
                    "First Name", "Last Name", "Name1", "Name2", "Name3", 
                    "Delivery Address", "Alternate 1 Address",
                    "City", "State", "ZIP+4", 
                    "BSC_ID", "Dependent Count"]
                    
    dependents_hdr = ["BSC_ID", "Dep First_Name", "Dep Last_Name"]

    
    fullrate_file = os.path.join(outdir, "FULLRATE.csv")
    dependents_file = os.path.join(outdir, "Dependents.dat")
    
    with open(fullrate_file, 'wb') as emp_handle:
        with open(dependents_file, 'wb') as dep_handle:
        
            dep_csv_wtr = csv.writer(dep_handle, quoting=csv.QUOTE_ALL)
            dep_csv_wtr.writerow(dependents_hdr)
            
            emp_csv_wtr = csv.writer(emp_handle, quoting=csv.QUOTE_ALL)
            emp_csv_wtr.writerow(employee_hdr)
            
            for dependentCount in sorted(mail_piece_dict.keys()):
                
                mail_records = mail_piece_dict[dependentCount]

                if len(mail_records) >= 1000:
                    file_count = int(math.ceil((len(mail_records) + 1000)/1000))
                    print file_count
                    for count in range(file_count):

                        presort_file = os.path.join(outdir, "Addr_{}dep_PT_{}.csv".format(dependentCount, count + 1))
                        
                        with open(presort_file, 'wb') as presort_handle:
                            psrt_csv_wtr = csv.writer(presort_handle, quoting=csv.QUOTE_ALL)
                            psrt_csv_wtr.writerow(employee_hdr)
                        
                            for seq, record in enumerate(mail_records, start=1):
                                if int(math.ceil((seq-1)/1000)) == count:

                                    imb_oel_sack_seq = ["","","","{}".format(seq)]
                                    employeeRow = imb_oel_sack_seq + record["employee"] + ["{}".format(dependentCount)]
                                    psrt_csv_wtr.writerow(employeeRow)
                                    
                                    for dependent in record["dependents"]:
                                        dep_csv_wtr.writerow(dependent)
                else:
                    
                    for seq, record in enumerate(mail_records, start=1):

                        imb_oel_sack_seq = ["","","","{}".format(seq)]
                        employeeRow = imb_oel_sack_seq + record["employee"] + ["{}".format(dependentCount)]
                        emp_csv_wtr.writerow(employeeRow)
                        
                        for dependent in record["dependents"]:
                            dep_csv_wtr.writerow(dependent)
                            
                            
                
            
def getCounts(outdir, mail_piece_dict):
    
    wb = openpyxl.Workbook()
    ws = wb.create_sheet("COUNTS", 0)    

    # Add header to worksheet
    ws.append(["Dependents per Employee","Mail Pieces","Records","Impressions"])

    total_mail_pieces = 0
    total_records = 0
    total_impressions = 0

    # Add counts to worksheet
    for dependentCount in sorted(mail_piece_dict.keys()):
        
        employee_count = len(mail_piece_dict[dependentCount])
        record_count = employee_count * dependentCount
        impression_count = employee_count * (dependentCount + 1)
        
        ws.append([dependentCount, employee_count, record_count, impression_count])
        
        # Increment totals
        total_mail_pieces += employee_count
        total_records += record_count
        total_impressions += impression_count
        
    ws.append([""])
    ws.append(["", total_mail_pieces, total_records, total_impressions])
    

    # Save Excel workbook
    counts_file = os.path.join(outdir, "PRINT_COUNTS.xlsx")
    wb.save(counts_file)

if __name__ == '__main__':
    main()
