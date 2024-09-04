# Open HTML Reports in Subfolders Script

This Python script searches for subfolders within the current directory whose names start with a number. It checks if those subfolders contain a folder named `HTMLReport`, and if that folder contains a file named `HTMLReport.html`. If the `HTMLReport.html` file exists, the script opens it using Microsoft Edge.

## Features

- Searches subfolders that start with a number (e.g., `12345/`, `67890/`).
- Looks for an `HTMLReport` folder inside those subfolders.
- Opens `HTMLReport.html` in Microsoft Edge if it exists.

## Requirements

- **Python 3.x** must be installed.
- **Microsoft Edge** must be installed on your system.
- Adjust the path to the Edge browser if it's different from the default location.

## Installation

1. Ensure Python 3.x is installed on your system.
2. Ensure Microsoft Edge is installed.
3. Clone or download the script file (`open_html_reports.py`).
4. Optionally, adjust the Edge browser path in the script if it's different from the path provided.

## Usage

1. Place the `open_html_reports.py` script in the root of the directory where your subfolders are located.

2. The directory structure should look something like this:

