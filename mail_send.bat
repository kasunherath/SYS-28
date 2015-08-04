for /f "skip=1" %%x in ('wmic os get localdatetime') do if not defined MyDate set MyDate=%%x
set today=%MyDate:~6,2%-%MyDate:~4,2%-%MyDate:~0,4%

mailsend1.16.exe -d genesiis.com -smtp 100.100.100.243 -t kasun@genesiis.com,ceo@genesiis.com -f cars_daily_report@genesiis.com -sub "CARS Report_%today%" -M "CARS daily report via 243Xmail" -attach "Report_%today%.csv"

