import win32com.client
import pythoncom
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import time

def extract_metadata(file_path):
    # Initialize COM object
    start_time = time.time()
    swApp = win32com.client.Dispatch("SldWorks.Application")
    swApp.Visible = False  # Run SolidWorks in the background
    end_time = time.time()
    print(f"Opening SldWorks.Application took {end_time - start_time:.2f} seconds")

    # Ensure the file path is in the correct string format
    file_path = str(file_path)

    # Open document with correct parameter types
    try:
        print(f"Opening file: {file_path}")
        
        # Initialize Errors and Warnings as VARIANT ByRef integers
        Errors = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        Warnings = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        
        # Constants for document type and open options
        swDocPART = 1
        swOpenDocOptions_Silent = 1  # Open silently (no dialogs)

        # Time the document opening
        start_time = time.time()
        doc = swApp.OpenDoc6(file_path, swDocPART, swOpenDocOptions_Silent, "", Errors, Warnings)
        end_time = time.time()
        print(f"OpenDoc6 took {end_time - start_time:.2f} seconds")
        
        if not doc:
            print(f"Failed to open document. Errors: {Errors.value}")
            return None
    except Exception as e:
        print(f"Error opening file: {e}")
        return None

    # Time the title retrieval
    start_time = time.time()
    title = doc.GetTitle
    end_time = time.time()
    print(f"GetTitle took {end_time - start_time:.2f} seconds")

    # Time the custom properties retrieval
    start_time = time.time()
    extension = doc.Extension
    prop_manager = extension.CustomPropertyManager("")
    prop_names = prop_manager.GetNames
    end_time = time.time()
    print(f"GetNames took {end_time - start_time:.2f} seconds")

    properties = {}
    if prop_names:
        for name in prop_names:
            # Prepare VARIANT variables for output parameters
            ValueOut = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, '')
            ResolvedValueOut = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, '')

            # Time the Get4 method for each property
            start_time = time.time()
            success = prop_manager.Get4(name, False, ValueOut, ResolvedValueOut)
            end_time = time.time()
            print(f"Get4 for property '{name}' took {end_time - start_time:.2f} seconds")

            if success:
                properties[name] = ResolvedValueOut.value
            else:
                properties[name] = None

    # Time the document closing
    start_time = time.time()
    swApp.CloseDoc(title)
    end_time = time.time()
    print(f"CloseDoc took {end_time - start_time:.2f} seconds")

    # Quit SolidWorks
    swApp.ExitApp()
    swApp = None

    # Create DataFrame
    data = {"Title": title, **properties}
    df = pd.DataFrame([data])
    
    return df

def browse_file():
    """
    Open a file dialog to select a SolidWorks part (.SLDPRT) file.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    file_path = filedialog.askopenfilename(
        title="Select a SolidWorks Part File",
        filetypes=[("SolidWorks Part Files", "*.SLDPRT")]
    )
    return file_path

def main():
    # file_path = browse_file()  # Uncomment this line if you want to use file dialog
    file_path = r'C:\Users\glen.merritt\Documents\topology_test.SLDPRT'  # Testing with hardcoded file path
    if file_path:
        df = extract_metadata(file_path)
        
        if df is not None:
            print(df.to_string(index=False))
            df.to_csv("metadata_output.csv", index=False)
        else:
            print("Failed to extract metadata.")
    else:
        print("No file selected.")

if __name__ == "__main__":
    main()
