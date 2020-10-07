import os
from openpyxl import load_workbook

path = input("Enter the full path of the .xlsx file you wish to print: ")
directory = os.path.dirname(path)

wb = load_workbook(path)


def dumpWsToText(ws, destination):
    """
    Writes worksheet contents to text file.

    ws: openpyxl Worksheet object
    destination: directory name into which the file will be written
    """
    name = f"{ws.title}.txt"
    path = os.path.join(destination, name)
    with open(path, "w") as newFile:
        for row in ws.values:
            for value in row:
                if value is None:
                    newFile.write((' ' * 19) + " |")
                else:
                    if len(str(value)) >= 15:
                        newFile.write(f"{str(value)[:16]}...".ljust(19) + " |")
                    else:
                        newFile.write(str(value).ljust(19) + " |")
            newFile.write("\n")


if __name__ == "__main__":
    for worksheet in wb.worksheets:
        dumpWsToText(worksheet, directory)
        print(f"Wrote {worksheet.title} to {directory}.")
