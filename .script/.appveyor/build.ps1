pyinstaller -F pyment/__main__.py -n $env:EXE_NAME -w -i ./rsrc/img/pyment.ico --hiddenimport xlrd --hiddenimport openpyxl --hiddenimport xlsxwriter
