import pandas as pd
from datetime import datetime, timedelta
import sys
 
try:
      Year=datetime.now().strftime('%Y')
      Month=datetime.now().strftime('%m')
      Day=datetime.now().strftime('%d')
      hour=datetime.now().strftime("%I")
      minute=datetime.now()+timedelta(minutes=20)
      TimeF=str(minute.strftime("%p"))
      minute=minute.strftime("%M")
      data={'Sr. No.':[1,2,3,4,5,6],
            'Key Name':['Day','Month','Year','Hour','Minute','TimeF'],
            'Value':[Day,Month,Year,hour,minute,TimeF]}
      df = pd.DataFrame(data)
      df.to_excel('D:/KITE/KITE/KITE_DATA/output.xlsx',index=False)
      print("Sucessfully created Excel")
 
except FileNotFoundError as F:
      print(F)