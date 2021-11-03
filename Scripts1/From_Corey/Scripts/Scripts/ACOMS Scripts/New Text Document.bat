FOR /F "tokens=*" %%A in (c:\users\1376638002e\desktop\jacce.txt) do (wmic /node:"%%A" product get name, version /format:csv > C:\users\1376638002e\desktop\SL_%%A.csv)
pause