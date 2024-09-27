# python-translator

Experimental python scripts that translates office files using Amazon Bedrock service

## Installation

```bash
$ pip install -r requirements.txt
```

## pptx translate
```bash
$ python pptx-translator.py --help
usage: Translates pptx files from source language to target language using Amazon Translate service
       [-h] [--terminology TERMINOLOGY]
       source_language_code target_language_code input_file_path

positional arguments:
  source_language_code  The language code for the language of the source text.
                        Example: en
  target_language_code  The language code requested for the language of the
                        target text. Example: pt
  input_file_path       The path of the pptx file that should be translated

optional arguments:
  -h, --help            show this help message and exit
  --terminology TERMINOLOGY
                        The path of the terminology CSV file
```

## xlsx translate
```bash
python xls-translator.py --help
usage: Translates Excel files from source language to target language using Amazon Translate service
       [-h] [--terminology TERMINOLOGY] --region REGION
       source_language_code target_language_code input_file_path

positional arguments:
  source_language_code  Source language code (e.g., en)
  target_language_code  Target language code (e.g., es)
  input_file_path       Path to the input Excel file

optional arguments:
  -h, --help            show this help message and exit
  --terminology TERMINOLOGY
                        Path to the terminology CSV file
  --region REGION       AWS region (e.g., us-west-2)
```
