# Office Document Translator

A simple yet powerful tool that uses Amazon Bedrock with Claude 3.7 to translate Excel and PowerPoint files while preserving all formatting.

![Translator Banner](https://via.placeholder.com/800x200?text=Office+Document+Translator)

## 📋 What It Does

This tool translates:
- Excel spreadsheets (.xlsx)
- PowerPoint presentations (.pptx)

While maintaining all original formatting including fonts, colors, borders, and layout.

## 🚀 Quick Start

```bash
# Translate an Excel file from English to French
python office-translator.py en fr document.xlsx

# Translate a PowerPoint from Spanish to German
python office-translator.py es de presentation.pptx
```

## 📦 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/scouturier/python-translator.git
   cd python-translator
   ```

2. Set up a Python virtual environment (recommended, especially on macOS):
   ```bash
   # Create a virtual environment
   python3 -m venv venv
   
   # Activate the virtual environment
   # On macOS/Linux:
   source venv/bin/activate
   
   # On Windows:
   # venv\Scripts\activate
   
   # Your terminal prompt should now show (venv) indicating the environment is active
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. When you're done using the translator, you can deactivate the virtual environment:
   ```bash
   deactivate
   ```

## ⚙️ AWS Setup

Before using the translator, you need:

1. An AWS account with Amazon Bedrock access
2. AWS credentials configured on your system

Set up your AWS environment:

```bash
# Option 1: Using AWS CLI
aws configure

# Option 2: Using environment variables
export AWS_ACCESS_KEY_ID=your_access_key
export AWS_SECRET_ACCESS_KEY=your_secret_key
export AWS_REGION=your_region  # e.g., us-east-1
```

## 📝 Usage Details

### Command Format
```
python office-translator.py [source_language] [target_language] [file_path]
```

### Parameters
- `source_language`: Language code of the source document (en, fr, de, es, etc.)
- `target_language`: Language code for translation (en, fr, de, es, etc.)
- `file_path`: Path to your Excel or PowerPoint file

### Output
The translated file is saved in the same directory with the target language code appended to the filename:
- `document.xlsx` → `document_fr.xlsx`
- `presentation.pptx` → `presentation_de.pptx`

## ✨ Features

### Excel Translation
- Translates all cell content
- Preserves:
  - Font styles and formatting
  - Cell alignment and orientation
  - Background colors and patterns
  - Borders and cell styles
  - Formulas (cell references remain intact)

### PowerPoint Translation
- Translates text in:
  - Regular slides
  - Shapes and text boxes
  - Tables and charts
  - Notes and comments
  - Grouped objects
- Preserves:
  - Slide layouts and designs
  - Animations and transitions
  - Speaker notes
  - All visual elements

## ❓ Troubleshooting

| Problem | Solution |
|---------|----------|
| `AWS_REGION environment variable is not set` | Run: `export AWS_REGION=your_region` |
| `ModuleNotFoundError: No module named 'package_name'` | Run: `pip install -r requirements.txt` |
| `botocore.exceptions.NoCredentialsError: Unable to locate credentials` | Configure AWS credentials with `aws configure` |
| Translation taking too long | Large files may require more time. Consider breaking into smaller files. |

## ⚠️ Limitations

- Only supports .xlsx and .pptx formats (not .xls or .ppt)
- Very large files may take significant time to process
- Character limits apply based on Amazon Bedrock quotas

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgements

- Uses Amazon Bedrock with Claude 3.7 for high-quality translations
- Built with Python, openpyxl, and python-pptx libraries
