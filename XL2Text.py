import os
from openpyxl import load_workbook


def dumpWsToText(ws, destination, colWidth=19):
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
                    fill = colWidth
                    newFile.write((' ' * fill) + " |")
                else:
                    maxW = int(colWidth * 0.80)
                    if len(str(value)) >= maxW:
                        newFile.write(
                            f"{str(value)[:maxW-1]}...".ljust(colWidth-1) + " |"
                        )
                    else:
                        newFile.write(str(value).ljust(colWidth) + " |")
            newFile.write("\n")


if __name__ == "__main__":
    path = input("Enter the full path of the .xlsx file you wish to print: ")
    directory = os.path.dirname(path)

    wb = load_workbook(path)

    specifyColWidth = input(
                        "Would you like to specify a column width? (y/n) "
                      ).lower()

    if specifyColWidth == "y":
        customWidth = int(
                        input(
                          "Please specify width (in characters, default: 19): "
                        )
                      )
    else:
        customWidth = None

    for worksheet in wb.worksheets:
        if customWidth:
            dumpWsToText(worksheet, directory, colWidth=customWidth)
        else:
            dumpWsToText(worksheet, directory)
        print(f"Wrote {worksheet.title} to {directory}.")
