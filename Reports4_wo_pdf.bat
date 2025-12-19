@echo off
setlocal
set "execution_path=Reportsbackup\A15_ExecutionReport_05_12_25_21_17_53"
set "bench_no=10"
set "build_number=A15"
set "file_name=Reports4_wo_pdf"
:wait_for_file
if not exist "D:\KITE_DATA\Execution\%execution_path%\MainDetailedReport.html" (
  ping 127.0.0.1 -n 10 > nul
  goto wait_for_file
)
cd /d "%~dp0"
python %file_name%.py D:\KITE_DATA\Execution\%execution_path% %bench_no% %build_number%
del %file_name%.py
del %file_name%.bat
endlocal