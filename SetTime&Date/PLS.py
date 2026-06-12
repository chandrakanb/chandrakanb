import pandas as pd 
from datetime import datetime, timedelta
import sys

try:
      Year=int(datetime.now().strftime('%Y'))
      Month=int(datetime.now().strftime('%m'))
      Day=int(datetime.now().strftime('%d'))
      hour=int(datetime.now().strftime("%I"))
      minute=int(datetime.now().strftime("%M"))+20
      TimeF=str(datetime.now().strftime("%p"))
      data={'Sr.No.':[1,2,3,4,5,6],
            'Key Name':['Day','Month','Year','Hour','Minute','TimeF'],
            'Value':[Day,Month,Year,hour,minute,TimeF]}
      df = pd.DataFrame(data)
      df.to_excel('D:/KITE/KITE/KITE_DATA/output.xlsx',index=False)
      print("Sucessfully created Excel")

except FileNotFoundError as F:
      print(F)