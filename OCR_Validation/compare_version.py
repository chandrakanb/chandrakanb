import sys
import pandas as pd

golden_value=open('Gold.txt', 'r').read()
actaul_value=open('Actaul.txt', 'r').read()

golden_value=golden_value.strip()
actaul_value=actaul_value.strip()

golden_value=golden_value[:7]
golden_value=golden_value[:-1]
for i in range(0,3):
    temp=golden_value+str(i)
    if temp==actaul_value:
        data = {'Sr. No.': [1,2,3],
                'Key Name': ['Result','actaul','golden'],
                'Value': ['Pass',actaul_value,golden_value]}
        df = pd.DataFrame(data)
        df.to_excel('D:/OCR_Validation/Result.xlsx', index=False)
        print(f"Pass : Actual : {actaul_value} Refrence : {temp}")
        sys.exit()

data = {'Sr. No.': [1,2,3],
        'Key Name': ['Result','actaul','golden'],
        'Value': ['Fail',actaul_value,golden_value]}
df = pd.DataFrame(data)
df.to_excel('D:/OCR_Validation/Result.xlsx', index=False)
print(f"Fail : Actual : {actaul_value} Refrence : {golden_value}")