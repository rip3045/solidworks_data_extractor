# SolidWorks Metadata Extractor

This Python script extracts metadata from SolidWorks `.SLDPRT` part files using the SolidWorks API. The extracted data is saved into a CSV file. 

## Features
- Open SolidWorks part files silently.
- Extract custom properties and metadata.
- Save metadata to a CSV file.

## Installation

To run this script, you need to have Python 3.x installed along with the following dependencies:

1. Install dependencies:

```bash
pip install -r requirements.txt

solidworks-metadata-extractor/
│
├── src/
│   ├── extract_metadata_python.py  # Main script
│
├── data/
│   └── metadata_output.csv  # Generated metadata file (once script runs)
│
├── README.md  # Project documentation
├── .gitignore  # Files to ignore in the repository
└── requirements.txt  # Dependencies

## Dependencies
pywin32 - For COM automation to interface with SolidWorks.
pandas - For data handling and saving the metadata to CSV.
tkinter - For GUI file dialog (if needed).
These dependencies are listed in the requirements.txt.

## Usage
Clone the repository:
git clone https://github.com/your-username/solidworks-metadata-extractor.git
cd solidworks-metadata-extractor

Run the script:
python src/extract_metadata_python.py

## File Browsing Option
You can select a SolidWorks part file using a file browser. Uncomment the following line in the main() function of the script:

