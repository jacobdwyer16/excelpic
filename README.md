# ExcelPic: Convert Excel Ranges to Images with Ease

Welcome to `ExcelPic`! This tool simplifies the process of converting parts of Excel files into images, perfect for presentations, documentation, or sharing on social media. Whether you need a snapshot of your latest data analysis or want to include Excel visuals in your reports, ExcelPic has you covered.

## Features:
- **Ease of Use**: Convert Excel ranges or entire sheets to images through a simple command-line interface.
- **Flexible**: Choose specific sheets or ranges to convert, or export the entire sheet.
- **Non-intrusive**: Runs in the background without displaying Excel or interrupting your workflow.
- **Safe**: Closes all COM objects gracefully, ensuring no Excel processes hang in the background.
- **Windows Native Operations**: Based on the `Excel2img` script, this code avoids the reliance on Pillow and screenshots. All image capture is done through scripting and COM operations.
- **Supports single Workbook connection**: To make the best use of I/O resources, a single workbook connection can be passed in and maintained, *or* a string path can be passed in and `ExcelPic` will take care of resource management.

## Prerequisites:
- Python 3.8 or greater (although it's probably fine to use older versions...)
- Windows OS (due to reliance on COM API)
- win32com.client, imgkit libraries installed
- wkhtmltoimg installed and on path for imgkit.

## Installation:
1. Clone this repository or download the source code.
2. Install the required Python libraries:
```bash
pip install pypiwin32 imgkit
```

## Usage:
To use `ExcelPic`, navigate to the directory containing the script and run:

### Python:
1. Converting an entire Excel file to an image:
```python
from excelpic import excelpic

# provide the path to your excel file and the desired output image filename
excel_path = 'path.xlsx'
output_image = 'output_image.png'

# Call the function to perform the conversion
excelpic(excel_path, output_image)
```

2. Converting a specific sheet and range within an Excel file:
```python
from excelpic import excelpic

# Define paths and parameters
excel_path = 'path.xlsx'
output_image = 'output_specific.png'
sheet_name = 'Sheet1'
range_spec = 'B2:F20'
imgkit_params = {"format": "png", "quality": 100, "zoom": 4}

# Perform the conversion with specific sheet and range
excelpic(excel_path, output_image, page = sheet_name, _range = range_spec, imgkit_params = imgkit_params)
```
3. Passing in an existing workbook connection and keeping the connection open.
```python
xlApp = win32.Dispatch("Excel.Application")
wb = xlApp.Workbooks.Open("workbook.xlsx")

# The existing workbook connection can be used
excelpic(wb, "image_location.png")

# The existing workbook connection is still active
wb.save()
wb.close()

xlApp.Quit()
```

### Command Line:
```bash
python excelpic.py excel_filename image_filename [options]
```

## Command Line Arguments

- excel_filename: The path to the Excel file you want to process.
- image_filename: The name and path for the output image file.
- -p, --page: Optional. Specify the Excel sheet by name. If not provided, the first sheet is used.
- -r, --range: Optional. Specify the Excel range in A1 notation. If not provided, the entire used range is selected.

## Examples:
Convert the entire first sheet to an image:

```bash
python excelpic.py example.xlsx example.png
```

Convert a specific range from a specific sheet:
```bash
python excelpic.py example.xlsx example.png -p Sheet1 -r A1:U8
```

## License
Distributed under the MIT License. See LICENSE for more information.

`ExcelPic` is designed to be a straightforward, powerful tool for converting Excel data into more shareable formats. Try it out today and streamline how you share your Excel insights!