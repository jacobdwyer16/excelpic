import argparse
import hashlib
import logging
import os
import re
import uuid
from typing import Any, Dict, Optional, Union

import imgkit
import win32com.client
from pythoncom import CoInitialize, CoUninitialize
from pywintypes import com_error

# COM constants
SOURCE_TYPE = 4
HTML_TYPE = 0
# Constants
LOG_SIZE = 50 * 1024 * 1024
BACKUP_COUNT = 3


class ExcelOpenError(Exception):
    """Exception raised for errors that occur during Excel file operations."""

    pass


class COMError(Exception):
    """Exception raised for errors that occur during COM operations."""

    pass


logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


def setup_logging(
    file_path: str = "application.log",
    level: int = logging.INFO,
    propagate: bool = False,
) -> None:
    """
    Sets up a logger with a rotating file handler.
    """
    handler = logging.FileHandler(file_path)
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    handler.setFormatter(formatter)
    logger.setLevel(level)
    logger.addHandler(handler)
    logger.propagate = propagate
    return


# Table formatting functions for intermediate steps between HTML and image
def extract_charset(html_path: str) -> str:
    """
    Extracts the charset from an HTML file.

    This function reads the first 2048 bytes of the file and searches for a charset declaration.
    If a charset is found, it is returned. If no charset is found, "utf-8" is returned.

    Args:
        file_path (str): The path to the HTML file.

    Returns:
        str: The charset found in the HTML file, or "utf-8" if no charset is declared.
    """
    charset_re = re.compile(rb'<meta.*?charset=["\']*(.+?)["\'>]', flags=re.IGNORECASE)

    with open(html_path, "rb") as file_:
        head = file_.read(2048)

    match = charset_re.search(head)
    if match:
        charset = match.group(1).decode("ascii", errors="ignore")
        return charset
    else:
        return "utf-8"  # Default to utf-8 if no charset is declared


def generate_hashed_filename(extension: str, modifier: Optional[str] = None) -> str:
    """
    Generatese a UUID SHA256 hash string to use for the table.

    Minimizes the odds of collision if excelpic is running in multiple instances on a shared env.

    Args:
        extension (str): Image file type (PNG, JPEG, SVG, etc.)
        modifier (str): Optional additional string to add to HTML file name.

    Returns:
        str: The file name to use when writing the intermediary HTML file
    """
    unique_id = uuid.uuid4()
    unique_id_bytes = str(unique_id).encode("utf-8")
    hash_object = hashlib.sha256(unique_id_bytes)
    hashed_id = hash_object.hexdigest()
    val: str = (
        f"{hashed_id}.{extension}"
        if modifier is None
        else f"{hashed_id}{modifier}.{extension}"
    )
    return val


def clean_html(file_path: str, charset: str) -> None:
    """
    Cleans an HTML file by removing invalid characters.

    This function reads the HTML file, removes any "�" characters, and writes the cleaned HTML back to the file.

    Args:
        file_path (str): The path to the HTML file.
        charset (str): The charset of the HTML file.
    """
    with open(file_path, "r", encoding=charset) as file_:
        html_content = file_.read()

    cleaned_content = html_content.replace("�", "")

    with open(file_path, "w", encoding=charset) as file_:
        file_.write(cleaned_content)


def css_to_remove_borders(file_path: str, charset: str) -> None:
    """
    Modifies the CSS of an HTML file to remove borders.

    This function reads the HTML file, adds CSS to the <body> and <table> tags to remove borders,
    and writes the modified HTML back to the file.

    Args:
        file_path (str): The path to the HTML file.
        charset (str): The charset of the HTML file.
    """
    with open(file_path, "r", encoding=charset) as file:
        html_content = file.read()

    new_css = """
        body {
            margin: 0;          /* Removes default margin */
            width: auto;        /* Allows width to adjust based on content */
            height: auto;       /* Allows height to adjust based on content */
        }
        table {
            width: 100%;        /* Makes table use full width of its container */
        }
    """

    style_pattern = re.compile(r"(<style[^>]*>)(.*?)(</style>)", re.DOTALL)

    if style_pattern.search(html_content):
        html_content = style_pattern.sub(r"\1\2" + new_css + r"\3", html_content, 1)
    else:
        head_pattern = re.compile(r"(</head>)", re.IGNORECASE)
        new_style_tag = f"<style>{new_css}</style>\n</head>"
        html_content = head_pattern.sub(new_style_tag, html_content)

    with open(file_path, "w", encoding=charset) as file:
        file.write(html_content)
    return


# Utilitiy functions for COM Operations
def com_initialize() -> None:
    CoInitialize()


def com_uninitialize() -> None:
    CoUninitialize()


class ExcelWorkbook(object):
    """
    Class that wraps Excel workbook for functions.
    """

    def __init__(self, com_object: Optional[win32com.client.CDispatch] = None) -> None:
        """
        Initializes an ExcelWorkbook instance, optionally wrapping a provided COM object.

        Args:
            com_object (Optional[CDispatch]): A COM Dispatch object that represents an Excel application,
                                            which if provided, will be used to initialize the workbook and app properties.
        """
        self.app: Optional[win32com.client.CDispatch] = (
            com_object.Application if com_object else None
        )
        self.workbook: Optional[win32com.client.CDispatch] = (
            com_object if com_object else None
        )

    def __enter__(self) -> Any:
        """
        Context manager entry handling. Returns the workbook instance itself when entering a context.

        Returns:
            Any: The instance of ExcelWorkbook currently being managed.
        """
        return self

    def __exit__(self, *args: Any) -> None:
        """
        Context manager exit handling. Ensures the workbook is closed when exiting a context.

        Args:
            *args (Any): Standard parameters for context manager exit (exception type, value, traceback).
        """
        self.close()

    @classmethod
    def open(cls, filename: str, read_only: bool = True) -> "ExcelWorkbook":
        """
        Opens an Excel workbook from the provided filename and initializes COM components.

        Args:
            filename (str): Path to the Excel file to be opened.

        Raises:
            IOError: If the file does not exist or an error occurs during opening.
        """
        excel_pathname = os.path.abspath(filename)  # excel requires abspath
        if not os.path.exists(excel_pathname):
            logger.error(f"No such excel file: {filename}")
            raise IOError(f"No such excel file: {filename}")
        com_initialize()
        try:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = 0
            app.DisplayAlerts = 0
            app.AskToUpdateLinks = 0
            workbook = app.Workbooks.Open(excel_pathname, ReadOnly=read_only)
            return_cls: "ExcelWorkbook" = cls(workbook)
            return return_cls
        except com_error as e:
            logger.error(f"COM  error occured while opening {filename}: {e}")
            raise COMError(f"Failed to open {filename}. COM error: {e}")
        except IOError as e:
            logger.error(f"Failed to open {filename}. Error: {e}")
            raise ExcelOpenError(f"Failed to open {filename}. Error: {e}")

    def close(self) -> None:
        """
        Closes the open Excel workbook and cleans up the COM objects, ensuring no resources are left hanging.
        """
        if self.app:
            if self.workbook:
                self.workbook.Close(SaveChanges=False)
                self.workbook = None
            self.app.Quit()
            self.app = None
        com_uninitialize()
        logger.info("Workbook closed and cache cleared.")


def _is_gen_py_object(obj: Any) -> bool:
    """
    Determines if a given object is a generated Python COM object.

    Args:
        obj (Any): The object to check.

    Returns:
        bool: True if the object is a generated Python COM object, False otherwise.
    """
    cls = obj.__class__
    module_name = cls.__module__
    val: bool = bool(module_name.startswith("win32com.gen_py"))
    return val


def _range_and_print(
    excel: ExcelWorkbook,
    fn_image: str,
    imgkit_params: Optional[Dict[str, Union[str, int, float]]],
    page: Optional[int] = None,
    _range: Optional[str] = None,
) -> None:
    """
    Processes an Excel range or entire sheet to export it as an image.

    Args:
        excel (ExcelWorkbook): The Excel workbook object.
        fn_image (str): Filename where the image should be saved.
        page (Optional[int]): Specific page number or name to target. If not specified, the current selection or first page is used.
        _range (Optional[str]): Specific Excel range in A1 notation to target. If not specified, the entire used range is targeted.
    """

    if _range is not None and page is not None and "!" not in _range:
        _range = f"'{page}'!{_range}"
    try:
        if excel.workbook is not None:
            if _range:
                rng = excel.workbook.Application.Range(_range)
            else:
                rng = excel.workbook.Sheets(page).UsedRange

        if _export_range_to_image(rng, excel, fn_image, imgkit_params):
            logger.info("Successfully generated picture ")
        else:
            logger.error("Unable to generate picture successfully.")
    except IOError as e:
        logger.error(f"Failed due to {e}")

    return


def _export_range_to_image(
    rng: win32com.client.CDispatch,
    excel: ExcelWorkbook,
    fn_image: str,
    imgkit_params: Optional[Dict[str, Union[str, int, float]]],
    temp_folder: str = "temporary_files",
) -> bool:
    """
    Exports a specified range from an Excel sheet to an image file using HTML as an intermediate format.

    Args:
        rng (CDispatch): The COM Dispatch object representing the Excel range.
        excel (ExcelWorkbook): The Excel workbook instance.
        fn_image (str): The path where the output image will be saved.
        temp_folder (str): The name of the temporary folder to save the HTML to.
        temp_file_name (str): Base name for temporary HTML file used during conversion.

    Returns:
        bool: True if the image was successfully created, False otherwise.
    """
    try:
        lib_dir = os.path.dirname(os.path.abspath(__file__))

        temp_cache = os.path.join(lib_dir, temp_folder)
        unique_file_name = generate_hashed_filename("html")
        temp_html_path = os.path.join(temp_cache, unique_file_name)

        if not os.path.exists(temp_cache):
            os.makedirs(temp_cache)

        if excel.app:
            pub = excel.app.ActiveWorkbook.PublishObjects.Add(
                SourceType=SOURCE_TYPE,
                Filename=temp_html_path,
                Sheet=rng.Worksheet.Name,
                Source=rng.Address,
                HtmlType=HTML_TYPE,
            )

            pub.Publish(True)

            charset = extract_charset(temp_html_path)
            clean_html(temp_html_path, charset)
            css_to_remove_borders(temp_html_path, charset)

            _imgkit_screenshot(temp_html_path, fn_image, options=imgkit_params)

            try:
                os.remove(temp_html_path)
            except IOError as e:
                logging.warning(f"Failed to delete intermediate HTML file: {e}")

            return True

        return False

    except (OSError, IOError) as e:
        logging.error(f"File system error during export: {e}", exc_info=True)
        return False

    except com_error as e:
        logging.error(f"COM error during export:{e}", exc_info=True)
        return False


def _imgkit_screenshot(
    html_path: str,
    fn_image: str,
    options: Optional[Dict[str, Union[str, int, float]]],
    wkhtmltoimage_path: Optional[str] = None,
) -> bool:
    """
    Converts HTML file to an image using the wkhtmltoimage tool.

    Args:
        html_path (str): Path to the HTML file to be converted.
        fn_image (str): Path where the resulting image should be saved.
        wkhtmltoimage_path (Optional[str]): Path to the wkhtmltoimage executable, if not in PATH.
        options (Optional[dict]): Dictionary of options for image conversion (e.g., format, quality).

    Returns:
        bool: True if the conversion was successful, False otherwise.
    """
    if wkhtmltoimage_path:
        os.environ["Path"] += os.pathsep + f"{wkhtmltoimage_path}"
    if not options:
        options = {"format": "png", "quality": 100, "zoom": 4}

    try:
        imgkit.from_file(html_path, fn_image, options=options)
        return True
    except IOError as e:
        logging.error(f"Error of {e} during generation")
    return False


def excelpic(
    fn_excel: Optional[Union[str, win32com.client.CDispatch]],
    fn_image: str,
    page: Optional[int] = None,
    _range: Optional[str] = None,
    imgkit_params: Optional[Dict[str, Union[str, int, float]]] = None,
) -> None:
    """
    Main function to handle exporting specified parts of an Excel file as images.

    Args:
        fn_excel (Optional[Union[str, CDispatch]]): The filename of the Excel file or an existing COM object to be processed.
        fn_image (str): The filename for the output image.
        page (Optional[int]): Specific sheet to use if not default or inferred.
        _range (Optional[str]): Specific range within the Excel sheet to export.
    """
    if isinstance(fn_excel, str):
        with ExcelWorkbook.open(fn_excel) as excel:
            # Pass in string name of file into a context manager.
            # Enter/Exit will be used and file connection will be closed after img generation.
            _range_and_print(excel, fn_image, imgkit_params, page, _range)
    elif _is_gen_py_object(fn_excel):
        # Directly pass the ExcelFile instance that wraps the win32com.client.CDispatch object
        _range_and_print(ExcelWorkbook(fn_excel), fn_image, imgkit_params, page, _range)

    return


if __name__ == "__main__":
    # This block handles command-line parsing and initiates the export based on provided arguments.
    parser = argparse.ArgumentParser(
        description="Script to export parts of an Excel file as images",
        usage="""{prog} excel_filename image_filename [options]\nExamples:
            {prog} example.xlsx example.png
            {prog} example.xlsx example.png -p Sheet1
            {prog} example.xlsx example.png -r NamedRange
            {prog} example.xlsx example.png -r 'Sheet1!A1:U8'
            {prog} example.xlsx example.png -r 'Sheet2!SheetRange' """,
    )
    parser.add_argument("excel_filename", type=str, help="Excel file name to process.")
    parser.add_argument(
        "image_filename", type=str, help="Output image file name and extension."
    )
    parser.add_argument(
        "-p",
        "--page",
        type=str,
        help="pick a page (sheet) by page name. When not specified (and RANGE either not specified or doesn't imply a page), first page will be selected",
    )
    parser.add_argument(
        "-r",
        "--range",
        metavar="RANGE",
        type=str,
        dest="_range",
        help="pick a range, in Excel notation. When not specified all used cells on a page will be selected",
    )
    args = parser.parse_args()

    excelpic(args.excel_filename, args.image_filename, args.page, args._range)
