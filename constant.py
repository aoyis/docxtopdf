from enum import Enum

from convert import FileTypeConverter


class FileType(Enum):
    DOCX = ".docx"
    PDF = ".pdf"


FILE_TYPE_MAPPING = {
    "docx": FileType.DOCX,
    "pdf": FileType.PDF,
}


CONVERSION_MAPPING = {
    (FileType.DOCX, FileType.PDF): FileTypeConverter.docx2pdf,
}
