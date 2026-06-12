import os

try:
    a=["Actaul.txt","Gold.txt","android_build_number.txt","CANwhitelist_actual.txt","Result.xlsx"]
    l=len(a)
    for i in range(0,l):
        os.remove(a[i])
except Exception as e:
    print(e)