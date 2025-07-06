import msoffcrypto-tool
import sys
import io
import os

def remove_password(input_file, password, output_file):
    try:
        # Check if input file exists
        if not os.path.isfile(input_file):
            print(f"Error: Input file '{input_file}' not found.")
            sys.exit(1)

        # Check file extension
        if not (input_file.lower().endswith(".xlsx") or input_file.lower().endswith(".xls")):
            print("Error: Unsupported file format. Only .xlsx and .xls files are supported.")
            sys.exit(1)

        # Open the encrypted Excel file
        with open(input_file, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            try:
                office_file.load_key(password)  # Provide the password
            except Exception:
                print("Error: Incorrect password or file not encrypted.")
                sys.exit(1)

            output = io.BytesIO()
            office_file.decrypt(output)

        # Check if output file is writable
        try:
            with open(output_file, "wb") as out:
                out.write(output.getvalue())
        except IOError:
            print(f"Error: Cannot write to output file '{output_file}'. Check permissions.")
            sys.exit(1)

        print(f"Password removed successfully. Output file: {output_file}")

    except Exception as e:
        print(f"Unexpected Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: remove_password.exe <input_file> <password> <output_file>")
        sys.exit(1)

    input_file = sys.argv[1]
    password = sys.argv[2]
    output_file = sys.argv[3]

    remove_password(input_file, password, output_file)
