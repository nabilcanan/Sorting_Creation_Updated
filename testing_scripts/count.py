import tkinter as tk
from tkinter import filedialog


def count_lines_of_code(files):
    loc_by_file = {}  # Dictionary to store lines of code by filename

    for filepath in files:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            loc = sum(1 for _ in f)  # Count lines
            loc_by_file[filepath] = loc

    return loc_by_file


def main():
    # Open file selection dialog
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    files = filedialog.askopenfilenames(title="Select .py files", filetypes=[("Python files", "*.py")])

    if not files:
        print("No files selected.")
        return

    result = count_lines_of_code(files)
    total_lines = 0
    for file, loc in result.items():
        print(f"{file}: {loc} lines of code")
        total_lines += loc

    print(f"\nTotal lines of code in selected files: {total_lines}")


if __name__ == "__main__":
    main()
