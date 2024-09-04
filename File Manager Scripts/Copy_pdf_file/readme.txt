This Python script, named `copy_pdf_reports.py`, iterates through all directories within the same repository to perform two key tasks:

1. **Create Root Directory**: It first creates a `PdfReports` directory in the root of the repository if it doesn't already exist.
2. **Copy Files**: The script then checks for the existence of an `XMLReport/PdfReports` directory in each subdirectory. If found, it copies the contents into the root `PdfReports` directory, preserving the original subdirectory structure.

This script is designed to streamline the process of collecting PDF reports from various parts of the repository into a central location.