
# Python3 program to convert docx to pdf
# using docx2pdf module
 
# Import the convert method from the
# docx2pdf module
from docx2pdf import convert

# Converting docx present in the same folder
# as the python file
# convert("GFG.docx")
 
# Converting docx present in the same folder
# as the python file
# convert("test.docx")
# convert("test.docx", "Other_Folder/word-test.pdf")

# Notice that the output filename need not be
# the same as the docx
 
# Bulk Conversion
# convert("GeeksForGeeks/")
def main():
    convert(".","word-output/")