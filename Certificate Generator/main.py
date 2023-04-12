from tkinter import Tk, filedialog, simpledialog
from generator import generate_certificates
from sort import organize_certificates
import sys

# User selection
root = Tk()
root.withdraw()

control_path = filedialog.askopenfilename(title="Name Control List", filetypes=[("Excel files", "*.xlsx")])

xlsx_file = filedialog.askopenfilename(title="Names", filetypes=[("Excel files", "*.xlsx")])
if not xlsx_file:
    sys.exit()

template_path = filedialog.askopenfilename(title="Template", filetypes=[("PowerPoint files", "*.pptx")])
if not template_path:
    sys.exit()

user_input = simpledialog.askstring(title="Training", prompt="Name of the event/course:")
if not user_input:
    sys.exit()

output_dir = filedialog.askdirectory(title="Folder to save")
if not output_dir:
    sys.exit()

generate_certificates(xlsx_file, template_path, user_input, output_dir)

if control_path:
    organize_certificates(output_dir, control_path)