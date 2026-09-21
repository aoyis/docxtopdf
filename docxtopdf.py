import sys
import os

import comtypes.client
from loguru import logger


logger.info(os.getcwd())
wdFormatPDF = 17
# print(os.listdir)


def convert_in_dir(word):
    for root, _, files in os.walk(os.getcwd()):
        os.chdir(root)
        for file in files:
            if file.endswith(".docx"):
                source = os.path.join(root, file)
                target = os.path.join(root, "".join(file.split(".")[:-1]) + ".pdf")

                # actual conversion
                doc = word.Documents.Open(source)
                doc.SaveAs(target, FileFormat=wdFormatPDF)
                doc.Close()
                    
                logger.success(f"Successfully converted {source} to {target}!")

    logger.success("Successfully converted all docx files to pdf!")
            
# TODO: might need to refactor and add single target file
            
word = comtypes.client.CreateObject('Word.Application')
convert_in_dir(word)

word.Quit()
