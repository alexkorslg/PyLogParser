logs parser
<br>
v0.23.12.1

---
### win

win10 schedule usage:

    description here

build script:

    ./build/build.cmd

build command:

>    python -m nuitka --onefile --enable-plugin=tk-inter --windows-disable-console --windows-company-name="PyLogParser" --windows-product-version="0.23.12.1" --include-data-file=".\logo.ico"="logo.ico" --windows-icon-from-ico=".\logo.ico" .\main.py

package will be unpacked to
>C:\Users\USER_NAME\AppData\Local\PyLogParser\app\VERSION\

to kill all processes `python.exe`

    taskkill /IM "python.exe" /F

---
