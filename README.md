# XLSX-Remove-Password

A command-line tool to remove the password from encrypted Excel files (`.xlsx` and `.xls`). This tool decrypts the file so it can be opened without requiring a password.

## Features
- Removes **workbook open password** from Excel files.
- Supports `.xlsx` and `.xls` formats.
- Runs as a standalone `.exe` (no need for Python on the target machine).
- Provides error handling for incorrect passwords and file permissions.

## Prerequisites
If you are running the script as Python, install dependencies:
```sh
pip install msoffcrypto
```

If using the EXE version, no dependencies are needed.

## Usage
### Python Script Usage
```sh
python remove_password.py <input_file> <password> <output_file>
```
Example:
```sh
python remove_password.py encrypted.xlsx mypassword decrypted.xlsx
```

### Windows EXE Usage
If using the compiled `.exe`, run:
```sh
remove_password.exe <input_file> <password> <output_file>
```
Example:
```sh
remove_password.exe encrypted.xlsx mypassword decrypted.xlsx
```

## Creating an Executable (EXE)
To create a standalone `.exe` file, use PyInstaller:
```sh
pip install pyinstaller
pyinstaller --onefile --hidden-import=msoffcrypto remove_password.py
```
The generated executable will be found in the `dist/` folder.

## Error Handling
- **File not found:** Ensure the input file exists.
- **Incorrect password:** Ensure the correct password is provided.
- **Unsupported format:** Only `.xlsx` and `.xls` files are supported.
- **Cannot write output file:** Ensure you have write permissions for the output location.

## License
This project is open-source and free to use.


My scripts:


py remove_password.py C:\Temp\Book1.xlsx P@ssw0rd C:\Temp\Book1unprotected.xlsx

py -m pip install msoffcrypto-tool

py pip install pyinstaller

pyinstaller --onefile remove_password.py

remove_password.exe C:\Temp\Book1.xlsx P@ssw0rd C:\Temp\Book1unprotectedex.xlsx