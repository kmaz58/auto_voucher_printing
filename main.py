import fitz  # PyMuPDF for PDF processing
import time
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pygetwindow as gw
import pytesseract
from PIL import Image
import io
import re
import threading
import pystray
from pystray import MenuItem as item
from win10toast import ToastNotifier
import logging
import shutil
import subprocess
import os
import sys

# -----------------------
# Logging Configuration
# -----------------------
LOG_FILENAME = "pdf_watcher.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILENAME),
        logging.StreamHandler()  # Also log to console (optional)
    ]
)
logger = logging.getLogger(__name__)

# -----------------------
# Configuration and Globals
# -----------------------
# # Path to the Tesseract executable (used for OCR)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# Path to SumatraPDF (used for silent printing)
sumatra_path = r"./SumatraPDF-3.5.2-64.exe"

# Define folders to watch and their categories
WATCH_DIRS = {
    r"C:\monpetit\Reporting": "a4",
    r"C:\epsilon net\Pylon Platform\Temp\PrintTemp": "a4",
    r"C:\monpetit\elta_vouchers": "a6",
    r"C:\Common_Downloads": "downloads",
    r"C:\monpetit\ACS": "a4"
}

# Create a ToastNotifier instance for Windows notifications
notifier = ToastNotifier()

# Define barcode or ID groupings per courier/label type
lastmile = {'801238128'}
DHL = {'DHL'}
A4 = {'094360202', '094058824'}
A6 = {'099759170', '800635204'}

# Observer state globals
running = False
observer = None

# -----------------------
# Detect Tesseract
# -----------------------

def find_tesseract():
    tesseract_path = shutil.which("tesseract")
    print(tesseract_path)
    common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]


    if tesseract_path:
        print(f"Tesseract found at: {tesseract_path}")
        return tesseract_path

    else:
        common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]

        for p in common_paths:
            if os.path.exists(p):
                print("Tesseract Found: " + str(p))
                return p

        # raise FileNotFoundError("Tesseract not found. Please install it or add to PATH.")

    

# -----------------------
# PDF Processing Functions
# -----------------------

def extract_text_from_pdf(file_path):
    """Extracts text from a PDF using PyMuPDF (fallbacks to OCR if necessary)."""
    try:
        doc = fitz.open(file_path)
        full_text = ""
        for page_num, page in enumerate(doc):
            text = page.get_text()
            if is_elta(text):
                logger.info(f"{file_path} likely contains ELTA VOUCHER. Falling back to OCR for numbers.")
                ocr_text = ocr_page_to_text(page, config="--psm 6 -c tessedit_char_whitelist=0123456789")
                full_text += ocr_text
            elif is_text_garbage(text):
                logger.info(f"{file_path} likely contains no readable text. Falling back to full OCR.")
                ocr_text = ocr_page_to_text(page, config="--psm 6")
                full_text += ocr_text
            elif text.strip():
                logger.info("Just reading available text")
                full_text += text
        return full_text
    except Exception as e:
        logger.error(f"Error extracting text from {file_path}: {e}", exc_info=True)
        return ""

def extract_9digit_numbers(text):
    """Extracts 9-digit numbers from provided text."""
    return re.findall(r'\b\d{9}\b', text)

def extract_mydhl(text):
    """Extracts occurrence(s) of the 'DHL' keyword from text."""
    return re.findall(r'DHL', text)

def ocr_page_to_text(page, config="--psm 6"):
    """Converts a PDF page to text using OCR via PyTesseract."""
    try:
        pix = page.get_pixmap(dpi=300)
        img_data = pix.tobytes("png")
        image = Image.open(io.BytesIO(img_data)).convert("L")
        thresholded = image.point(lambda x: 0 if x < 150 else 255, '1')
        return pytesseract.image_to_string(thresholded, lang='eng', config=config)
    except Exception as e:
        logger.error(f"OCR error on page: {e}", exc_info=True)
        return ""

def decide_action_on_downloads(ids_found, filepath):
    """Determines the action based on extracted ID(s) and sends the file to the appropriate printer."""
    try:
        for id_ in ids_found:
            if id_ in lastmile:
                printer_name = "SATO WS408 Petit"
                selected_pdf = select_pdf_area_LASTMILE(filepath)
                print_pdf(selected_pdf, printer_name)
                os.remove(selected_pdf)
                logger.info("Action: approved Skroutz Lastmile")
                return "approved Skroutz Lastmile", id_
            elif id_ in A4:
                printer_name = "Brother HL-L6250DN series Printer"
                print_pdf(filepath, printer_name)
                logger.info("Action: approved A4")
                return "approved A4", id_
            elif id_ in A6:
                printer_name = "SATO WS408 Petit"
                print_pdf(filepath, printer_name)
                logger.info("Action: approved A6")
                return "approved A6", id_
            elif id_ in DHL:
                printer_name = "SATO WS408 Petit"
                selected_pdf = select_pdf_area_DHL(filepath)
                print_pdf(selected_pdf, printer_name)
                logger.info("Action: approved DHL")
                return "approved DHL", id_
        logger.warning("Action: unknown ID")
        return "unknown", ids_found[0] if ids_found else None
    except Exception as e:
        logger.error(f"Error in decide_action: {e}", exc_info=True)
        return "error", None

def is_text_garbage(text):
    """Determines if text is mostly unreadable (e.g., very few alphanumerical characters)."""
    return len(re.findall(r'[a-zA-Z0-9]', text)) < 5

def is_elta(text):
    """Checks for the substring 'elta' (case-insensitive) in text."""
    return 'elta' in text.lower()

def decide_paper_size(filepath):
    """Analyzes PDF content to decide the paper size and routing for printing."""
    try:
        text = extract_text_from_pdf(filepath)
        if 'MyDHL' not in text:
            numbers = extract_9digit_numbers(text)
            status, matched_id = decide_action_on_downloads(numbers, filepath)
        else:
            text = extract_mydhl(text)
            status, matched_id = decide_action_on_downloads(text, filepath)
        logger.info(f"Status: {status}")
        logger.info(f"Matched ID: {matched_id}")
    except Exception as e:
        logger.error(f"Error in decide_paper_size: {e}", exc_info=True)

def print_pdf(file_path, printer_name):
    """Sends a print command to SumatraPDF for the specified file and printer."""
    try:
        cmd = [
            f'"{sumatra_path}"',
            "-print-to", f'"{printer_name}"',
            f'"{file_path}"'
        ]
        logger.info(f"Print command issued for {file_path} on {printer_name}")
        subprocess.run(" ".join(cmd), shell=True)
    except Exception as e:
        logger.error(f"Error printing {file_path}: {e}", exc_info=True)

def select_pdf_area_LASTMILE(input_pdf):
    """Crops a specific area from the original PDF for Skroutz Lastmile and saves it as a temporary file."""
    try:
        temp_pdf_path = os.path.join("./selected_area.pdf")
        x1 = -20
        y1 = 80
        width = 280
        height = 460

        doc = fitz.open(input_pdf)
        new_doc = fitz.open()

        for page in doc:
            crop_rect = fitz.Rect(x1, y1, x1 + width, y1 + height)
            new_page = new_doc.new_page(width=280, height=400)
            new_page.show_pdf_page(new_page.rect, doc, page.number, clip=crop_rect)

        new_doc.save(temp_pdf_path)
        logger.info(f"Temporary cropped PDF created: {temp_pdf_path}")
        return temp_pdf_path
    except Exception as e:
        logger.error(f"Error in select_pdf_area_LASTMILE: {e}", exc_info=True)
        return input_pdf

def select_pdf_area_DHL(input_pdf):
    """Crops a specific area from the original PDF for DHL and saves it as a temporary file."""
    try:
        temp_pdf_path = os.path.join("./selected_area.pdf")
        x1 = 80
        y1 = 30
        width = 300
        height = 550

        doc = fitz.open(input_pdf)
        new_doc = fitz.open()

        for page in doc:
            crop_rect = fitz.Rect(x1, y1, x1 + width, y1 + height)
            new_page = new_doc.new_page(width=280, height=400)
            new_page.show_pdf_page(new_page.rect, doc, page.number, clip=crop_rect)

        new_doc.save(temp_pdf_path)
        logger.info(f"Temporary cropped PDF created: {temp_pdf_path}")
        return temp_pdf_path
    except Exception as e:
        logger.error(f"Error in select_pdf_area_DHL: {e}", exc_info=True)
        return input_pdf

# -----------------------
# File System Event Handling
# -----------------------
def handle_file_creation(folder, file_path):
    """Handles a new or moved PDF file in a monitored directory."""
    logger.info(f"File found: {file_path}")
    action = WATCH_DIRS.get(folder)
    try:
        if action == "a4":
            printer_name = "Brother HL-L6250DN series Printer"
            time.sleep(3)
            print_pdf(file_path, printer_name)
        elif action == "a6":
            printer_name = "SATO WS408 Petit"
            print_pdf(file_path, printer_name)
        elif action == "downloads":
            decide_paper_size(file_path)
        else:
            logger.warning(f"Unknown action for folder: {folder}")
    except Exception as e:
        logger.error(f"Error on handle_file_creation for {file_path}: {e}", exc_info=True)

class PDFHandler(FileSystemEventHandler):
    """Handles file system events for PDFs."""
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".pdf"):
            folder = os.path.dirname(event.src_path)
            try:
                if folder in WATCH_DIRS:
                    logger.info(event)
                    # print("File Size: " + str(os.path.getsize(event.src_path.lower())))
                    print(event.src_path)
                    handle_file_creation(folder, event.src_path)

                    # if os.path.getsize(event.src_path) == 0 :
                    #     while os.path.getsize(event.src_path) == 0 :
                    #         time.sleep(2)
                    #         print("Waiting")
                    #     handle_file_creation(folder, event.src_path)
                    # else :
                    #     handle_file_creation(folder, event.src_path)
            except Exception as e:
                logger.error(f"Error on_created for {event.src_path}: {e}", exc_info=True)


    def on_moved(self, event):
        if not event.is_directory and event.dest_path.lower().endswith(".pdf"):
            folder = os.path.dirname(event.dest_path)
            if folder in WATCH_DIRS:
                logger.info(event)
                print("File Size: " + str(os.path.getsize(event.dest_path)))
                try:
                    if os.path.getsize(event.dest_path) == 0 :
                        while os.path.getsize(event.dest_path) == 0 :
                            time.sleep(2)
                        handle_file_creation(folder, event.dest_path)
                    else :
                        handle_file_creation(folder, event.dest_path)
                except Exception as e:
                    logger.error(f"Error on_moved for {event.dest_path}: {e}", exc_info=True)


# -----------------------
# Observer Control Functions
# -----------------------
def start_observer():
    """Starts the file system observer for watching PDFs and sends a notification."""
    global observer, running
    if not running:
        logger.info("Starting observer...")
        notifier.show_toast("PDF Watcher", "Observer started", duration=5, threaded=True)
        try:
            observer = Observer()
            for folder in WATCH_DIRS:
                observer.schedule(PDFHandler(), folder, recursive=False)
            observer.start()
            running = True
            logger.info("Observer started successfully.")
        except Exception as e:
            logger.error(f"Error starting observer: {e}", exc_info=True)

def stop_observer():
    """Stops the file system observer."""
    global observer, running
    if running and observer:
        logger.info("Stopping observer...")
        try:
            observer.stop()
            observer.join()
            running = False
            logger.info("Observer stopped successfully.")
        except Exception as e:
            logger.error(f"Error stopping observer: {e}", exc_info=True)

# -----------------------
# Tray Integration
# -----------------------
def get_state_label():
    """Return a label showing the current state of the observer."""
    return "Observer: Running" if running else "Observer: Stopped"

def start_observer_callback(icon, item):
    start_observer()
    return 0

def stop_observer_callback(icon, item):
    stop_observer()
    return 0

def quit_app(icon, item):
    stop_observer()
    icon.stop()
    # Forcefully exit the program
    os._exit(0)

def setup_tray():
    """Creates the system tray icon with dynamic menu displaying current observer state."""
    try:
        icon_image = Image.open("./tray_icon.ico")  # Ensure tray_icon.ico exists.
        menu = pystray.Menu(
            item(lambda icon: get_state_label(), None),
            item("Start", start_observer_callback),
            item("Stop", stop_observer_callback),
            item("Quit", quit_app)
        )
        tray_icon = pystray.Icon("PDF Watcher", icon_image, "PDF Watcher", menu)
        tray_icon.run()
    except Exception as e:
        logger.error(f"Error setting up tray icon: {e}", exc_info=True)

# -----------------------
# Main Execution Block
# -----------------------
if __name__ == "__main__":

    # Check Tesseract installation
    tesseract_path = find_tesseract()
    if not tesseract_path:
        try:
            # Assuming the installer is in the same directory as the .exe
            installer_path = r"tesseract-ocr-w64-setup-5.5.0.20241111.exe"
            print(installer_path)
            print("Tesseract not found. Running installer...")
            subprocess.run(installer_path, shell=True)
            tesseract_path = find_tesseract()
        except Exception:
            raise(Exception)
        # raise Exception("Tesseract installation not found. Please install it.")
    pytesseract.pytesseract.tesseract_cmd = tesseract_path
    print(tesseract_path)
    
    # Start the system tray icon in a background thread.
    tray_thread = threading.Thread(target=setup_tray, daemon=True)
    tray_thread.start()

    logger.info("PDF watcher running in system tray...")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        stop_observer()
        os._exit(0)
