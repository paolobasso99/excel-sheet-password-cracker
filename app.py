import os
import binhex
import re
import zipfile
import shutil
import glob
import time
from lxml import etree


def main():
    print("Starting ...")
    parser = etree.XMLParser(remove_comments=False)

    # Empty tmp
    if(os.path.isdir('.tmp')):
        print("Empting .tmp...")
        empty_folder('.tmp')

    if os.path.exists("out.xlsm"):
        print("Removing old out.xlsm...")
        os.remove("out.xlsm")

    # password = "macro"
    pattern = b'DPB\=\"(.*?)\"'

    # Check if there are multiple .xlsm files
    xlsmCounter = len(glob.glob1('.', "*.xlsm"))
    if(xlsmCounter > 1):
        print("Only one .xlsm file must be on this folder!")
        time.sleep(3)
        exit()

    # Extract all the contents of zip file in current directory
    os.chdir(".")
    with zipfile.ZipFile(glob.glob("*.xlsx")[0], 'r') as zipObj:
        print("Extracting xlsm...")
        zipObj.extractall('.tmp')

    # Unclock each sheet
    sheets = glob.glob(".tmp/xl/worksheets/*.xml")
    for sheet in sheets:
        print("Reading " + sheet + "...")
        tree = etree.parse(sheet, parser=parser)
        root = tree.getroot()

        output = re.sub(b"(?s)<sheetProtection .*?/>",  b"", etree.tostring(root))

        with open(sheet, 'wb') as f:
            f.write(output)

    # Create out.xlsm
    print("Creating out.xlsm...")
    shutil.make_archive('out', 'zip', '.tmp')
    os.rename("out.zip", "out.xlsx")

    # Success
    print("SUCCESS: The file out.xml has no sheets password")

    # Empty tmp
    if(os.path.isdir('.tmp')):
        print("Empting .tmp...")
        empty_folder('.tmp')

    # Exit
    input("Press Enter to exit...")


def empty_folder(folder):
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))


if __name__ == "__main__":
    main()
