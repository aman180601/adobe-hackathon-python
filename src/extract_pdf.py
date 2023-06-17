# Import required libraries.
from adobe.pdfservices.operation.auth.credentials import Credentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.execution_context import ExecutionContext
from adobe.pdfservices.operation.io.file_ref import FileRef
from adobe.pdfservices.operation.pdfops.extract_pdf_operation import ExtractPDFOperation
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_pdf_options import ExtractPDFOptions
from adobe.pdfservices.operation.pdfops.options.extractpdf.extract_element_type import ExtractElementType
import os
import zipfile
import json
import logging
import re
import pandas as pd

# Directories.
base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
resource_path = base_path + "/resources/"
zip_file = base_path + "/output/ExtractTextInfoFromPDF.zip"

# Rename input pdf files (output0 -output9) so that it executes sequentially (in top down order as in directory).
for i in range(0,10):
    if(os.path.isfile(resource_path + 'output'+ str(i) + '.pdf')):
        os.rename(resource_path + 'output'+ str(i) + '.pdf',resource_path + 'output0'+ str(i) + '.pdf')

# Loop for extracting each pfd file from resource and appending the dataframes to an output csv file one by one.
index = 0
for filename in sorted(os.listdir(resource_path)):
      input_pdf = resource_path + filename

      try:
            # Initial setup, create credentials instance.
            credentials = Credentials.service_account_credentials_builder()\
            .from_file(base_path + "/pdfservices-api-credentials.json") \
            .build()

            # Create an ExecutionContext using credentials and create a new operation instance.
            execution_context = ExecutionContext.create(credentials)
            extract_pdf_operation = ExtractPDFOperation.create_new()

            # Set operation input from a source file.
            source = FileRef.create_from_local_file(input_pdf)
            extract_pdf_operation.set_input(source)

            # Build ExtractPDF options and set them into the operation
            extract_pdf_options: ExtractPDFOptions = ExtractPDFOptions.builder() \
            .with_element_to_extract(ExtractElementType.TEXT) \
            .build()
            extract_pdf_operation.set_options(extract_pdf_options)

            # Execute the operation.
            result: FileRef = extract_pdf_operation.execute(execution_context)

            # Save the result to the specified location.
            result.save_as(zip_file)

            # Read data from generated json file.
            archive = zipfile.ZipFile(zip_file,'r')
            jsonentry = archive.open('structuredData.json')
            jsondata = jsonentry.read()
            data = json.loads(jsondata)
            jsonentry.close()
            archive.close()
            os.remove(zip_file)
     
            # Create three dataframes from json file and append them to csv file.
            # Generate second dataframe (df2) containing informations about items ordered (name, quantity and rate) from pdf invoice.
            temp_series = pd.Series(["","",""])
            df2 = pd.DataFrame(columns = [0,1,2])
            b = False
            i = 0
            j = 0
            for element in data["elements"]:
                  if("Text" in element):
                        if(element["Text"].startswith("AMOUNT")):
                              b = True
                              continue
                        elif(element["Text"].startswith("Subtotal")):
                              break
                        if(b == True and i < 3):
                              temp_series[i] = element["Text"][:-1]
                              i = i + 1
                        elif(i == 3):
                              df2.loc[j] = temp_series
                              i = 0
                              j = j + 1

            # Extract data from json file to create other two dataframes (df1 and df3).     
            # Extract invoice number and issue date.
            invoice = ""
            b = False
            for element in data["elements"]:  
                  if("Text" in element):
                        if(element["Text"].startswith("BILL TO")):
                              if("Table" in element["Path"]):
                                    b = True
                              break
                        elif(element["Text"].startswith("Invoice")):
                              b = True
                        elif(element["Text"].startswith("NearBy")):
                              b = False      
                        if(b==True):
                              invoice = invoice + element["Text"]      

            # Extract bill to, details and due date of invoice depending on whether they are in table structure or not.
            bill = ""
            detail = ""
            due_date = ""
            b1 = False
            b2 = False
            if(b == False):
                  for element in data["elements"]:  
                        if("Text" in element):
                              if(element["Text"].startswith("BILL TO")):
                                    b1 = True
                              elif(element["Text"].startswith("DETAILS") or element["Text"].startswith("PAYMENT")):
                                    b2 = True
                                    b1 = False
                              elif(element["Text"].startswith("ITEM")):
                                    break 
                              if(element["Text"].startswith("Due date:")):
                                    due_date = due_date + element["Text"]    
                              if(b1 == True):
                                    bill = bill + element["Text"]
                              elif(b2 == True):
                                    detail = detail + element["Text"]
            else:
                  for element in data["elements"]:
                        if("Text" in element):
                              if(re.match("//Document/(Sect|Sect\[.\])/Table/TR\[.\]/TD/",element["Path"])):
                                    bill = bill + element["Text"]
                              elif(re.match("//Document/(Sect|Sect\[.\])/Table/TR\[.\]/TD\[2\]/",element["Path"])):
                                    detail = detail + element["Text"]
                              elif(re.match("//Document/(Sect|Sect\[.\])/Table/TR\[.\]/TD\[3\]/",element["Path"])):
                                    due_date = due_date + element["Text"]
                              elif(element["Text"] == "ITEM "):
                                    break                 

            bill = bill.replace("BILL TO ","")               
            bill = bill.replace("BILL TO","")
            detail = detail.replace("DETAILS ","")
            detail = detail.replace("DETAILS","")
            detail = detail.replace("PAYMENT ","")
            detail = detail.replace("PAYMENT","")
            detail = detail.replace(due_date,"")
            detail = detail.replace(".","")
            detail = detail.replace("$","")
            detail = re.sub("\d+","",detail)
      
            # Generate third dataframe (df3) using above informations (issue date, due date, invoice number, details and tax %).
            temp_series = pd.Series(["","",invoice.split(' ')[4],invoice.split(' ')[1],"10"])
            temp_series[0] = detail[:-1]
            temp_series[1] = due_date.split(' ')[2]
            n = df2.shape[0]
            df3 = pd.DataFrame(columns = [0,1,2,3,4])
            i = 0
            while(i < n):
                  df3.loc[i] = temp_series
                  i = i + 1
      
            # Generate first dataframe (df1) with rest of the informations.
            temp_series = pd.Series(["Jamestown","Tennessee, USA","We are here to serve you better. Reach out to us in case of any concern or feedbacks.","NearBy Electronics","3741 Glory Road","38556","","",bill.split(' ')[0] + ' ' + bill.split(' ')[1],bill.split(' ')[2],""])
            i = 0
            if(not bill.split(' ')[2].endswith(".com")):
                  i = 1
                  temp_series[9] = temp_series[9] + bill.split(' ')[3]
            temp_series[10] = bill.split(' ')[i + 3]
            temp_series[6] = bill.split(' ')[i + 4] + ' ' + bill.split(' ')[i+5] + ' ' + bill.split(' ')[i + 6]
            temp_series[7] = bill.split(' ',i + 7)[i + 7][:-1]
            df1 = pd.DataFrame(columns = [0,1,2,3,4,5,6,7,8,9,10])
            i = 0
            while(i < n):
                  df1.loc[i] = temp_series
                  i = i + 1

            # Generate a new dataframe (df) by merging all the three dataframes (df1 + df2 + df3).
            df = pd.concat([df1,df2,df3],axis = 1)
            df.columns = ['Bussiness__City','Bussiness__Country','Bussiness__Description','Bussiness__Name','Bussiness__StreetAddress','Bussiness__Zipcode','Customer__Address__line1','Customer__Address__line2','Customer__Email','Customer__Name','Customer__PhoneNumber','Invoice__BillDetails__Name','Invoice__BillDetails__Quantity','Invoice__BillDetails__Rate','Invoice__Description','Invoice__DueDate','Invoice__IssueDate','Invoice__Number','Invoice__Tax']
      
            # Lastly append the dataframe (df) to the output csv file.
            if (index > 0):
                  df.to_csv(base_path + '/output/ExtractedData.csv', mode = 'a',index = False,header = False)
            else:
                  df.to_csv(base_path + '/output/ExtractedData.csv',index = False)
            print("Successfully extracted pdf",index)
            index = index + 1

      except (ServiceApiException, ServiceUsageException, SdkException):
            logging.exception("Exception encountered while executing operation")     