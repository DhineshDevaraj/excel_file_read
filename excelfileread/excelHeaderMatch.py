import pandas as pd

file_read_header = ['CINVOICENBR' 'DINVDATE' 'CPRODCODE' 'CPRODTYPE' 'CAWBNO' 'Rough'
                    'Rough center' 'Purpose' 'Center' 'Rough.1' 'DBATCHDT' 'ORGIN'
                    'Destiination' 'NACTWGT' 'NCHRGWT' 'NPCS' 'NTOTALAMT' 'R off']

data_visualization_header = ['Sr_no' 'Lab_ID' 'REGISTRATION_DATE' 'PROJECT_CODE' 'CENTER_CODE'
                             'REGION' 'METROS/NON-METROS' 'MOTHER_NAME' 'CRM_NO' 'KIT_BOX_NO'
                             'DELIVERY_DATE' 'Delivery Date' 'RECEIPT_TIME' 'Courier' 'Docket NO'
                             'Receipt Dt & Time (Log)' 'Receipt date' 'Booking Date' 'Pick up City'
                             'Overall TAT' 'LOG TAT' 'BD TAT' 'Booking TAT' 'Overall TAT Status']

df = pd.read_excel("D:/DB.xlsx")

import_headers = df.axes[1]
print(import_headers)

data_count = [i for i in import_headers if i not in file_read_header]
if len(data_count) == 0:
    print("Start the Process")
else:
    print("File Contains Invalid Header")
    pass

