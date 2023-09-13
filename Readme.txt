[Python Module]
pip install pyinstaller
pip install pywin32
pip install pygetwindow
pip install pyautogui
pip install requests
pip install urllib3
pip install comtypes
pip install uuid
pip install pywifi
pip install git+https://github.com/airgproducts/pybluez2.git@0.46



[Make *.py, *.ico and extra *.exe to *.exe]
pyinstaller --clean -F --icon "Power On-Off Test_S5.ico" "Power On-Off Test_S5.py"



[UPX]
upx --force -9 "Power On-Off Test_S5.exe"
pyinstaller --clean -F --upx-dir "C:\upx" --icon "Power On-Off Test_S5.ico" "Power On-Off Test_S5.py"



[Visual Studio Code Package for Python]
Python
Pylance
Chinese (Traditional) Language Pack for Visual Studio Code
One Dark Pro
https://blog.kyomind.tw/good-vscode-extensions/



[BT tools]
btcom -n "20:64:DE:85:8B:FE"
btcom -b "20:64:DE:85:8B:FE" -r -s110b
btcom -b "20:64:DE:85:8B:FE" -c -s110b



[BT module of pybluez issue]
#Fix Windows Platform SDK issue
https://www.filehorse.com/download-microsoft-windows-sdk/download/

#Fix use_2to3 issue
pip install setuptools==57.5.0

#Fix Irprops.lib issue
cd C:\Program Files (x86)\Windows Kits\10\Lib\10.0.22000.0\um\x64
or
cd C:\Program Files (x86)\Windows Kits\10\Lib\10.0.22621.0\um\x64
mklink IRPROPS.LIB Bthprops.lib
