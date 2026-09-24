from enum import Enum
from pathlib import Path

import comtypes.client


class WdSaveFormat(Enum):
    wdFormatDocumentDefault = 16  # .docx
    wdFormatPDF = 17  # .pdf


class FileTypeConverter:

    @staticmethod
    def docx2pdf(
        source_file: Path,
        target_file: Path,
    ) -> None:
        com_object = comtypes.client.CreateObject('Word.Application')
        docx = com_object.Documents.Open(source_file)
        docx.SaveAs(target_file, FileFormat=WdSaveFormat.wdFormatPDF.value)
        docx.Close()
        com_object.Quit()
