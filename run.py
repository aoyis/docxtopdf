import os
from typing import Optional

import fire
from loguru import logger

from constant import FileType, FILE_TYPE_MAPPING, CONVERSION_MAPPING


class ConverterRunner:

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
        source_type: FileType,
        target_type: FileType,
    ) -> None:
        source_path = os.getcwd()
        logger.info(f"Source dir: {source_path}")
        for root, _, files in os.walk(source_path):
            os.chdir(root)
            for file in files:
                if file.endswith(source_type.value):
                    source_file = os.path.join(root, file)
                    target_file = os.path.join(root, "".join(file.split(".")[:-1]) + target_type.value)
                    # conversion
                    CONVERSION_MAPPING[(source_type, target_type)](source_file, target_file)
                    logger.success(f"Successfully converted {source_file} to {target_file}!")
        logger.success("Successfully converted all docx files to pdf!")
                
    @staticmethod
    def run(
        source: Optional[str] = None,
        target: Optional[str] = None,
        source_file_type: str = "docx",
        target_file_type: str = "pdf",
    ) -> None:
        source_type = FILE_TYPE_MAPPING[f"{source_file_type}"]
        target_type = FILE_TYPE_MAPPING[f"{target_file_type}"]
        if source is not None and target is not None:
            ConverterRunner.convert(
                source,
                target,
                source_type=source_type,
                target_type=target_type,
            )
        else:
            ConverterRunner.recursive_convert(
                source_type=source_type,
                target_type=target_type,
            )


if __name__ == "__main__":
    fire.Fire(ConverterRunner.run)
