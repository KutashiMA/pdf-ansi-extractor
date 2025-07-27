# PDF to Excel Extractor for ANSI Standards

A Python tool for extracting structured data from ANSI (American National Standards Institute) PDF documents and converting them to Excel format for further analysis.

## Features

-   **PDF Text Extraction**: Extracts text content from multi-page PDF documents.
-   **Structured Data Parsing**: Intelligently parses ANSI standards information, including operating names, legal names, websites, document titles, and publishing dates.
-   **Excel Export**: Outputs clean, structured data to a `.xlsx` file.
-   **Modular Design**: Features an object-oriented architecture for easy maintenance and extension.
-   **Error Handling**: Includes robust error handling for file operations and data processing.

## Project Structure

The project is organized into the following directories and files:

```
pdf-ansi-extractor/
├── src/
│   └── pdf_extractor.py    # Main Python script with the extraction logic
├── data/
│   ├── input/              # Place your source PDF file here
│   └── output/             # Directory where the extracted Excel file will be saved
├── requirements.txt        # List of Python dependencies
├── README.md               # This file
└── .gitignore              # Specifies files and folders to be ignored by Git
```

## How to Use

### 1. Clone the Repository

First, clone the repository to your local machine:

```bash
git clone https://github.com/YOUR_USERNAME/pdf-ansi-extractor.git
cd pdf-ansi-extractor
```

### 2. Set Up a Virtual Environment

It is highly recommended to use a virtual environment to manage project dependencies.

```bash
# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows
venv\Scripts\activate

# On macOS/Linux
source venv/bin/activate
```

### 3. Install Dependencies

Install the required Python libraries using the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

### 4. Run the Script

Place the PDF file you want to process into the `data/input/` directory. Then, execute the script from the root of the project folder:

```bash
python src/pdf_extractor.py
```

The script will process the PDF, print its progress to the console, and save the resulting Excel file (`ANSI_Extracted.xlsx`) in the `data/output/` directory.

## Advanced Usage

The `ANSIStandardsExtractor` class can be imported and used in other Python scripts for more advanced workflows.

```python
from src.pdf_extractor import ANSIStandardsExtractor

# Initialize the extractor with custom paths
extractor = ANSIStandardsExtractor(
    pdf_path="data/input/your_file.pdf",
    output_path="data/output/"
)

# Process the PDF and save with a custom filename
df, output_file = extractor.process("custom_output_name.xlsx")

# Access the resulting DataFrame
print(df.head())
print(f"Extracted {len(df)} standards")
```

## Output

The tool generates an Excel file with the following columns:

| Column          | Description                              |
| --------------- | ---------------------------------------- |
| Operating Name  | Short name or acronym of the organization |
| Legal Name      | Full legal name of the organization      |
| Website         | Organization's website URL               |
| Document Name   | The ANSI/INCITS document identifier      |
| Standard Title  | Full title of the standard               |
| Publishing Date | The final action date from the document  |
| American Standard | A flag indicating "American" standard    |

### Example Console Output:

After running the script, you will see console output similar to this:

```
Extracting text from PDF...
Parsing text data...
Creating DataFrame...
Extracted 150 records
DataFrame shape: (150, 7)
Saving to Excel...
Data saved to: data/output/ANSI_Extracted.xlsx

First 5 rows of extracted data:
  Operating Name      Legal Name         Website      Document Name  ...
0           ASSP   (American Soc)  www.assp.org       ANSI/ASSP A10.1  ...
1           ASSP   (American Soc)  www.assp.org       ANSI/ASSP A10.3  ...
2           ASSP   (American Soc)  www.assp.org       ANSI/ASSP A10.4  ...
3           ASSP   (American Soc)  www.assp.org       ANSI/ASSP A10.5  ...
4           ASSP   (American Soc)  www.assp.org       ANSI/ASSP A10.6  ...
```

## Contributing

Contributions are welcome! Please follow these steps:

1.  Fork the repository.
2.  Create a feature branch (`git checkout -b feature/new-feature`).
3.  Commit your changes (`git commit -am 'Add new feature'`).
4.  Push to the branch (`git push origin feature/new-feature`).
5.  Create a Pull Request.

## Author

**Kutashi Muhammed**

-   **GitHub**: `[@KutashiMA](https://github.com/KutashiMA)`
-   **LinkedIn**: `[Muhammed Kutashi](https://www.linkedin.com/in/muhammed-kutashi-645a25243/)`