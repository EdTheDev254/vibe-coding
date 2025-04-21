import tkinter as tk
from tkinter import messagebox
import sys
import os
import platform
import subprocess
from PIL import ImageGrab, Image, ImageOps, ImageFilter
import pytesseract
import pyperclip
from typing import Optional, Tuple

# --- Configuration ---

# Tesseract Path Handling:
# Set to None initially. The script will try to find it, prioritizing
# the bundled location ('Tesseract-OCR' subdirectory) if found.
TESSERACT_CMD_PATH: Optional[str] = None

# OCR Configuration
DEFAULT_TESSERACT_CONFIG = '-l eng --oem 3 --psm 6' # English, Default OCR Engine, Assume uniform text block

# Preprocessing options
RESIZE_FACTOR = 2 # How much to scale up the image before OCR
THRESHOLD_VALUE = 150 # Simple binary threshold value (0-255)
USE_ADAPTIVE_THRESHOLD = False # Set to True to try adaptive thresholding
GRAYSCALE = True # Convert to grayscale

# --- Helper Functions ---

def get_application_path() -> str:
    """Gets the base path for the running application (frozen or script)."""
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (e.g., by PyInstaller)
        application_path = os.path.dirname(sys.executable)
    else:
        # If run as a normal Python script
        try:
            # Use __file__ if available
            application_path = os.path.dirname(os.path.realpath(__file__))
        except NameError:
            # Fallback to current working directory if __file__ is not defined
            application_path = os.getcwd()
    return application_path

def find_tesseract_executable() -> Optional[str]:
    """
    Tries to find the tesseract executable, prioritizing bundled location.
    Returns the full path if found outside PATH, or None if relying on PATH.
    Returns None and logs error if completely not found.
    """
    base_path = get_application_path()
    exe_name = "tesseract.exe" if platform.system() == "Windows" else "tesseract"
    print(f"Application base path: {base_path}") # Log for debugging

    # 1. Check explicitly configured path (if TESSERACT_CMD_PATH was manually set)
    if TESSERACT_CMD_PATH and os.path.exists(TESSERACT_CMD_PATH):
         print(f"Using explicitly configured Tesseract path: {TESSERACT_CMD_PATH}")
         return TESSERACT_CMD_PATH

    # 2. Check relative path (for Tesseract bundled via Inno Setup in 'Tesseract-OCR' subdir)
    #    This is the primary method expected when installed via the setup.
    relative_path = os.path.join(base_path, "Tesseract-OCR", exe_name)
    print(f"Checking bundled path: {relative_path}") # Log for debugging
    if os.path.exists(relative_path):
        print(f"Found Tesseract bundled at: {relative_path}")
        return relative_path

    # 3. Check default PyInstaller temp location (less common for the exe itself)
    if getattr(sys, 'frozen', False):
        meipass_path = os.path.join(getattr(sys, '_MEIPASS', base_path), exe_name)
        if os.path.exists(meipass_path):
            print(f"Found Tesseract potentially in MEIPASS: {meipass_path}")
            return meipass_path

    # 4. Check system PATH (pytesseract uses this if cmd not set, but we verify first)
    try:
        # Try running 'tesseract --version' to see if it's in the PATH
        print(f"Checking for '{exe_name}' in system PATH...") # Log for debugging
        subprocess.run([exe_name, "--version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0)
        print(f"Found Tesseract via system PATH: '{exe_name}'. Pytesseract should find it.")
        # Let pytesseract handle finding it via PATH by returning None
        return None
    except (subprocess.CalledProcessError, FileNotFoundError):
        print(f"Tesseract command '{exe_name}' not found or failed in system PATH.")
        pass # Continue to check default locations

    # 5. Check common default install locations (Windows example)
    if platform.system() == "Windows":
        print("Checking common default Windows install locations...") # Log for debugging
        default_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        for path in default_paths:
            if os.path.exists(path):
                 print(f"Found Tesseract at default location: {path}")
                 return path

    # If not found anywhere
    print("Tesseract executable could not be located.")
    return "NOT_FOUND" # Return a specific marker for not found


# --- GUI Class ---

class ScreenshotGUI:
    """GUI application for selecting a screen region and extracting text using OCR."""

    def __init__(self, root_window: tk.Tk):
        """Initialize the GUI window and its components."""
        self.root = root_window
        self.root.title("Screenshot Text Extractor")
        self.root.minsize(350, 250) # Adjusted min size

        # Make window resizable
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)

        self.start_button = tk.Button(self.root, text="Select Region to Capture Text", command=self.start_selection, padx=10, pady=5)
        self.start_button.grid(row=0, column=0, columnspan=2, pady=10, padx=10, sticky="ew")

        self.text_frame = tk.Frame(self.root) # Frame to hold text box and scrollbar
        self.text_frame.grid(row=1, column=0, columnspan=2, pady=(0, 5), padx=10, sticky="nsew")
        self.text_frame.rowconfigure(0, weight=1)
        self.text_frame.columnconfigure(0, weight=1)

        self.text_box = tk.Text(self.text_frame, height=10, width=50, wrap=tk.WORD, borderwidth=1, relief="solid")
        self.text_box.grid(row=0, column=0, sticky="nsew")

        self.scrollbar = tk.Scrollbar(self.text_frame, command=self.text_box.yview)
        self.scrollbar.grid(row=0, column=1, sticky='nsew')
        self.text_box['yscrollcommand'] = self.scrollbar.set

        # Status Bar
        self.status_var = tk.StringVar()
        self.status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=5)
        self.status_bar.grid(row=2, column=0, columnspan=2, sticky='ew')
        self.update_status("Ready.")

        # Selection state variables
        self.selection_window: Optional[tk.Toplevel] = None
        self.canvas: Optional[tk.Canvas] = None
        self.start_x: Optional[int] = None
        self.start_y: Optional[int] = None
        self.rect_id: Optional[int] = None

    def update_status(self, message: str):
        """Updates the status bar text."""
        self.status_var.set(message)
        # self.root.update_idletasks() # Avoid forcing updates too often, can cause flicker

    def start_selection(self):
        """Create a transparent, fullscreen window for region selection."""
        self.root.withdraw()
        # Short delay might help ensure main window is hidden
        self.root.after(100, self._create_selection_window)

    def _create_selection_window(self):
        """Creates the actual selection window."""
        try:
            self.selection_window = tk.Toplevel(self.root)
            self.selection_window.attributes('-fullscreen', True)
            # Slightly less transparent might be easier to see on some screens
            self.selection_window.attributes('-alpha', 0.35)
            self.selection_window.attributes('-topmost', True)
            self.selection_window.focus_force()

            self.canvas = tk.Canvas(self.selection_window, cursor="cross", bg="grey")
            self.canvas.pack(fill=tk.BOTH, expand=True)

            self.canvas.bind("<Button-1>", self.on_mouse_down)
            self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
            self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
            self.canvas.bind("<Escape>", self.cancel_selection)

            self.update_status("Click and drag to select region. Press Esc to cancel.")
        except tk.TclError as e:
             print(f"Error creating selection window: {e}")
             messagebox.showerror("Window Error", "Could not create selection overlay.\nEnsure your graphics environment supports transparency.")
             self.cleanup_selection()
             self.root.deiconify()


    def cleanup_selection(self):
         """Destroys selection window and resets state variables."""
         if self.selection_window:
             try:
                 self.selection_window.destroy()
             except tk.TclError:
                 pass # Window might already be destroyed
         self.selection_window = None
         self.canvas = None
         self.rect_id = None
         self.start_x = None
         self.start_y = None

    def cancel_selection(self, event=None):
         """Cancels the selection process."""
         self.update_status("Selection cancelled.")
         self.cleanup_selection()
         # Bring back the main window
         if self.root.state() == 'withdrawn':
             self.root.deiconify()
         self.root.focus_force()


    def on_mouse_down(self, event):
        """Record starting position of the selection."""
        self.start_x = event.x
        self.start_y = event.y
        if self.canvas:
            self.rect_id = self.canvas.create_rectangle(
                self.start_x, self.start_y,
                self.start_x, self.start_y,
                outline='red', width=2
            )

    def on_mouse_drag(self, event):
        """Update the rectangle coordinates as the mouse moves."""
        if self.rect_id is not None and self.start_x is not None and self.start_y is not None and self.canvas:
            try:
                self.canvas.coords(self.rect_id, self.start_x, self.start_y, event.x, event.y)
            except tk.TclError:
                # Handle cases where canvas might be destroyed during drag
                pass

    def on_mouse_up(self, event):
        """Finalize selection, capture screenshot, and perform OCR."""
        if self.rect_id is None or self.start_x is None or self.start_y is None or not self.selection_window:
            self.cancel_selection()
            return

        try:
            # Get final coordinates from the canvas relative to the screen
            # Use winfo_rootx/y which should be relative to the screen origin
            x1 = self.selection_window.winfo_rootx() + self.start_x
            y1 = self.selection_window.winfo_rooty() + self.start_y
            x2 = self.selection_window.winfo_rootx() + event.x
            y2 = self.selection_window.winfo_rooty() + event.y

            # Ensure correct ordering for bbox
            left = min(x1, x2)
            top = min(y1, y2)
            right = max(x1, x2)
            bottom = max(y1, y2)

            # Prevent zero-size or tiny captures
            if abs(left - right) < 5 or abs(top - bottom) < 5:
                self.update_status("Selection too small, please try again.")
                self.cancel_selection()
                return

            # Important: Destroy selection window *before* taking screenshot
            # to avoid capturing the overlay itself.
            self.cleanup_selection()
            self.root.update() # Process window destruction events

            # Capture the screen region
            self.update_status("Capturing region...")
            screenshot = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
            self.update_status("Processing image...")

            # --- Image Preprocessing ---
            if GRAYSCALE:
                 processed_image = screenshot.convert('L')
            else:
                 processed_image = screenshot.copy()

            if USE_ADAPTIVE_THRESHOLD:
                 processed_image = ImageOps.autocontrast(processed_image, cutoff=10)
                 processed_image = processed_image.point(lambda p: 255 if p > THRESHOLD_VALUE else 0, mode='1')
            else:
                 processed_image = processed_image.point(lambda p: 255 if p > THRESHOLD_VALUE else 0, mode='1')

            if RESIZE_FACTOR > 1:
                try:
                    w, h = processed_image.size
                    processed_image = processed_image.resize((w * RESIZE_FACTOR, h * RESIZE_FACTOR), Image.LANCZOS)
                except Exception as resize_err:
                    print(f"Error resizing image: {resize_err}")
                    # Proceed without resizing if it fails

            # --- OCR ---
            self.update_status("Performing OCR...")
            try:
                text = pytesseract.image_to_string(processed_image, config=DEFAULT_TESSERACT_CONFIG)
            except pytesseract.TesseractNotFoundError:
                 # This error *should* have been caught at startup, but handle as fallback
                 messagebox.showerror("Tesseract Error", "Tesseract executable not found or inaccessible.\nPlease ensure it's installed correctly.")
                 self.update_status("Error: Tesseract not found.")
                 return # Stop processing
            except Exception as ocr_err:
                 # Catch other potential Tesseract errors
                 messagebox.showerror("OCR Error", f"An error occurred during text recognition:\n{ocr_err}")
                 self.update_status("Error during OCR.")
                 print(f"Detailed OCR Error: {ocr_err}")
                 return # Stop processing


            # --- Display & Copy ---
            final_text = text.strip()
            self.text_box.delete(1.0, tk.END)
            self.text_box.insert(tk.END, final_text)

            if final_text:
                 try:
                     pyperclip.copy(final_text)
                     self.update_status("Text extracted and copied to clipboard.")
                 except pyperclip.PyperclipException as clip_err:
                     print(f"Error copying to clipboard: {clip_err}")
                     self.update_status("Text extracted (clipboard error).")
            else:
                 self.update_status("OCR finished, but no text detected.")

        except ImageGrab.GrabError as grab_err:
             print(f"Error grabbing screenshot: {grab_err}")
             messagebox.showerror("Capture Error", f"Failed to capture the screen region:\n{grab_err}")
             self.update_status("Error during capture.")
        except Exception as e:
            error_msg = f"An unexpected error occurred:\n{type(e).__name__}: {e}"
            print(f"Detailed Error: {error_msg}") # Log detailed error
            messagebox.showerror("Error", f"Failed to process screenshot. Check console for details.\nError: {type(e).__name__}")
            self.update_status(f"Error: {type(e).__name__}")
        finally:
            # Ensure cleanup happens and main window reappears
            self.cleanup_selection()
            if self.root.state() == 'withdrawn':
                self.root.deiconify()
            self.root.focus_force()


    def run(self):
        """Start the GUI event loop."""
        self.root.mainloop()

# --- Main Execution ---
if __name__ == "__main__":

    # 1. Find and Configure Tesseract Path
    tesseract_path_result = find_tesseract_executable()

    tesseract_ok = False
    if tesseract_path_result == "NOT_FOUND":
        # Tesseract was not found in bundled location, PATH, or defaults
        # Create a temporary root just for the error message
        root_check = tk.Tk()
        root_check.withdraw()
        messagebox.showerror("Tesseract Configuration Error",
                             "Tesseract OCR executable could not be located.\n\n"
                             "Please ensure:\n"
                             "1. Tesseract was installed correctly by the setup program.\n"
                             "2. (If installed separately) Tesseract is in your system PATH.\n\n"
                             "The application cannot function without Tesseract.")
        root_check.destroy()
        sys.exit("Tesseract not found. Exiting.")

    elif tesseract_path_result is not None:
        # Found in bundled location, explicit config, or default location
        pytesseract.pytesseract.tesseract_cmd = tesseract_path_result
        print(f"Pytesseract configured to use Tesseract at: {tesseract_path_result}")
        tesseract_ok = True
    else:
        # find_tesseract_executable returned None, meaning it *should* be in PATH
        # Let's verify pytesseract can actually use it from PATH
        print("Verifying Tesseract accessibility via system PATH using pytesseract...")
        try:
            version = pytesseract.get_tesseract_version()
            print(f"Pytesseract successfully accessed Tesseract (Version: {version}) via system PATH.")
            tesseract_ok = True
        except pytesseract.TesseractNotFoundError:
             # Even though subprocess found it, pytesseract might have issues
             root_check = tk.Tk()
             root_check.withdraw()
             messagebox.showerror("Tesseract Configuration Error",
                                 "Tesseract was found in your system PATH, but the application could not access it.\n\n"
                                 "Potential issues:\n"
                                 "- Incorrect Tesseract installation/permissions.\n"
                                 "- Environment variable caching (try restarting).\n\n"
                                 "The application cannot function without accessible Tesseract.")
             root_check.destroy()
             sys.exit("Tesseract found but inaccessible via Pytesseract. Exiting.")
        except Exception as e:
             root_check = tk.Tk()
             root_check.withdraw()
             messagebox.showerror("Tesseract Check Error",
                                  f"An unexpected error occurred while verifying Tesseract:\n{e}")
             root_check.destroy()
             sys.exit(f"Error verifying Tesseract: {e}. Exiting.")


    # 2. Proceed only if Tesseract is configured/accessible
    if tesseract_ok:
        # Set up and run the main application window
        main_root = tk.Tk()
        app = ScreenshotGUI(main_root)
        try:
            app.run()
        except KeyboardInterrupt:
            print("\nApplication exited by user.")
        except Exception as final_err:
             print(f"\nUnhandled application error: {final_err}")
             # Try showing a final error message if GUI loop crashes badly
             try:
                  messagebox.showerror("Fatal Application Error", f"A critical error occurred:\n{final_err}")
             except:
                  pass # If Tkinter is completely broken, we can't show message