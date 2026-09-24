# File Type Converter
Program written in python to convert files on Windows machine with Microsoft Office installed.
This repo is forked from `baptistegaut/docxtopdf`. 

## How to run
To convert single file:
```
python run.py \
--source source.docx \
--target target.pdf \
--source_file_type docx \
--target_file_type pdf
```
To convert all files under current working directory recursively:
```
python run.py \
--source_file_type docx \
--target_file_type pdf
```
