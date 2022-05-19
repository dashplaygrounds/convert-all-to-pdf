import os

import excel_to_pdf_convert_all as etp
import pptx_to_pdf_convert_all as ptp
import word_to_pdf_convert_all as wtp


def main():
    print("Job started: Converting all to pdf")
    dir1 = 'excel-output'
    dir2 = 'pptx-output'
    dir3 = 'word-output'
    if not os.path.exists(dir1):
        os.mkdir(dir1)
    if not os.path.exists(dir2):
        os.mkdir(dir2)
    if not os.path.exists(dir3):
        os.mkdir(dir3)
    etp.main()
    print("Done! 1/3")
    ptp.main()
    print("Done! 2/3")
    wtp.main()
    print("Done! 3/3")
    print("Job finished!")

main()
