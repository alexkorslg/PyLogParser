python -m nuitka --onefile --enable-plugin=tk-inter --windows-disable-console ^
--windows-file-description="PyLogParser" ^
--windows-company-name="PyLogParser" --windows-product-version="0.23.12.1" ^
--include-data-file=".\logo.ico"="logo.ico" ^
--windows-icon-from-ico=".\logo.ico" ^
.\main.py