# Office Document Translator

This Python script uses Amazon Bedrock with Claude 3.5 to translate Excel (.xlsx) and PowerPoint (.pptx) files while preserving their formatting.

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)
- AWS account with access to Amazon Bedrock
- AWS credentials configured on your system

## Installation

1. Clone this repository or download the script files:
   ```
   git clone [repository-url]
   # or download manually
   ```

2. Install the required Python packages:
   ```
   pip install -r requirements.txt
   ```

## AWS Configuration

1. Ensure you have AWS credentials configured with access to Amazon Bedrock. You can set this up by:
   - Using AWS CLI: `aws configure`
   - Or setting environment variables:
     ```
     export AWS_ACCESS_KEY_ID=your_access_key
     export AWS_SECRET_ACCESS_KEY=your_secret_key
     ```

2. Set the AWS region environment variable:
   ```
   export AWS_REGION=your_region  # e.g., us-east-1
   ```

## Usage

The script can translate both Excel and PowerPoint files:

```
python office-translator.py [source_language_code] [target_language_code] [input_file_path]
```

Examples:
```
# Translate Excel file from English to Spanish
python office-translator.py en es document.xlsx

# Translate PowerPoint file from French to German
python office-translator.py fr de presentation.pptx
```

### Parameters:
- `source_language_code`: The language code of the source document (e.g., en, fr, de)
- `target_language_code`: The language code of the desired translation (e.g., es, it, ja)
- `input_file_path`: Path to the Excel (.xlsx) or PowerPoint (.pptx) file

The translated file will be saved in the same directory as the input file, with the target language code appended to the filename.

## Features

- Preserves formatting in Excel files:
  - Font styles
  - Cell alignment
  - Cell colors
  - Borders

- Preserves formatting in PowerPoint files:
  - Text in shapes
  - Text in tables
  - Text in grouped objects

## Common Issues

1. **AWS Region not set**
   ```
   Error: AWS_REGION environment variable is not set
   ```
   Solution: Set the AWS_REGION environment variable as shown in the AWS Configuration section.

2. **Missing dependencies**
   ```
   ModuleNotFoundError: No module named 'package_name'
   ```
   Solution: Ensure you've installed all requirements using `pip install -r requirements.txt`

3. **AWS credentials not found**
   ```
   botocore.exceptions.NoCredentialsError: Unable to locate credentials
   ```
   Solution: Configure AWS credentials using AWS CLI or environment variables as shown in the AWS Configuration section.

## Limitations

- The script only supports .xlsx and .pptx files
- Very large files may take significant time to process