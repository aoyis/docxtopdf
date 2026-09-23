import os
from pathlib import Path
from typing import Literal, Optional

import comtypes.client
import fire
from loguru import logger


class FileTypeConverter:
    @staticmethod
    def docx2pdf(
        source_file: Path,
        target_file: Path,
    ) -> None:
        com_object = comtypes.client.CreateObject('Word.Application')
        docx = com_object.Documents.Open(source_file)
        wdFormatPDF = 17
        docx.SaveAs(target_file, FileFormat=wdFormatPDF)
        docx.Close()
        com_object.Quit()

    # TODO: might need to refactor and add single target file
    @staticmethod
    def convert(
        source_path: str,
        target_path: str,
        source_type: str = ".docx",
        target_type: str = ".pdf",
    ) -> None:
        ...


    @staticmethod
    def recursive_convert(
        source_type: str,
        target_type: str,
    ) -> None:
        source_path = os.getcwd()
        logger.info(f"Source dir: {source_path}")
        for root, _, files in os.walk(source_path):
            os.chdir(root)
            for file in files:
                if file.endswith(source_type):
                    source_file = os.path.join(root, file)
                    target_file = os.path.join(root, "".join(file.split(".")[:-1]) + target_type)
                    # TODO: match source/target type to method
                    FileTypeConverter.docx2pdf(source_file, target_file)
                    logger.success(f"Successfully converted {source_file} to {target_file}!")
        logger.success("Successfully converted all docx files to pdf!")
                
    @staticmethod
    def run(
        source: Optional[str] = None,
        target: Optional[str] = None,
        source_type: str = ".docx",
        target_type: str = ".pdf",
    ) -> None:
        if source is not None and target is not None:
            FileTypeConverter.convert(
                source,
                target,
                source_type=source_type,
                target_type=target_type,
            )
        else:
            FileTypeConverter.recursive_convert(
                source_type=source_type,
                target_type=target_type,
            )


if __name__ == "__main__":
    fire.Fire(FileTypeConverter.run)
