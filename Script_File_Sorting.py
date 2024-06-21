import xml.etree.ElementTree as ET
import os
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog


def get_name_from_xml(xml_file):
    try:
        print(f"Parsing XML file: {xml_file}")
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # Print the XML structure for debugging
        print(f"XML structure:\n{ET.tostring(root, encoding='unicode')}")

        # Try to find <NamCE17> first, then <NamEX17>
        name_element = root.find('.//NamCE17')
        if name_element is None:
            name_element = root.find('.//NamEX17')

        if name_element is None:
            raise ValueError(f"No <NamCE17> or <NamEX17> element found in the XML file: {xml_file}")

        name = name_element.text.strip()
        print(f"Extracted name: {name}")
        return name
    except ET.ParseError as e:
        print(f"Error parsing XML file {xml_file}: {e}")
        raise
    except Exception as e:
        print(f"An error occurred while processing XML file {xml_file}: {e}")
        raise


def move_pdf_based_on_name(pdf_file, target_directory_base, name):
    try:
        print(f"Preparing to move PDF file: {pdf_file}")
        if not os.path.isfile(pdf_file):
            raise FileNotFoundError(f"PDF file not found: {pdf_file}")

        current_year = datetime.now().year
        target_path = os.path.join(target_directory_base, "ΠΕΛΑΤΕΣ", name, "ΔΙΑΣΑΦΗΣΕΙΣ", str(current_year))

        if not os.path.exists(target_path):
            print(f"Creating target directory: {target_path}")
            os.makedirs(target_path)

        print(f"Moving PDF to {target_path}")
        shutil.move(pdf_file, os.path.join(target_path, os.path.basename(pdf_file)))
        print(f"Successfully moved PDF to {target_path}")
    except Exception as e:
        print(f"An error occurred while moving the PDF file: {e}")
        raise


def process_files_in_directory(directory, target_directory_base):
    if not os.path.exists(directory):
        print(f"Directory does not exist: {directory}")
        return

    files = os.listdir(directory)
    if not files:
        print(f"No files found in directory: {directory}")
        return

    print(f"Files found in directory: {files}")

    for filename in files:
        if filename.lower().endswith(".xml"):
            xml_file = os.path.join(directory, filename)
            pdf_file = os.path.join(directory, filename.rsplit('.', 1)[0] + ".pdf")

            print(f"Processing file: {filename}")

            try:
                name = get_name_from_xml(xml_file)
                move_pdf_based_on_name(pdf_file, target_directory_base, name)
            except FileNotFoundError as e:
                print(f"PDF file not found for {filename}: {e}")
            except Exception as e:
                print(f"Failed to process {filename}: {e}")


def main():
    print("Opening file dialog...")
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    directory = filedialog.askdirectory(title="Select the directory containing XML and PDF files")

    if not directory:
        print("No directory selected. Exiting.")
        return

    print(f"Selected directory: {directory}")
    target_directory_base = "D:/"

    process_files_in_directory(directory, target_directory_base)


if __name__ == "__main__":
    main()
