import pandas as pd
import openpyxl
import math
import wx
import platform

# This script splits the contents of an excel file into worksheets, by a predefined number of rows.
# It asks for an excel file, then makes a copy of it. Only processes the first sheet of a file.

# Define the number of rows per sheet
ROWS_PER_SHEET = 30000

def data_breakup(filename: str):
    # Read the Excel file
    data = pd.read_excel(filename)
    
    # Calculate the number of sheets required
    no_of_sheets = math.ceil(len(data) / ROWS_PER_SHEET)
    
    # Print the number of blocks (sheets) and the length of the data
    #print(f"Blocks: {no_of_sheets}")
    #print(f"File length: {len(data)}")
    
    # Define the output filename
    output_filename = f"{filename.rsplit('.', 1)[0]}_split.xlsx"
    
    # Create a Pandas Excel writer using openpyxl as the engine
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        # Split the data into blocks and save to different sheets in the same file
        for x in range(no_of_sheets):
            start_row = x * ROWS_PER_SHEET
            end_row = start_row + ROWS_PER_SHEET
            
            # Use iloc to select the correct rows
            new_data = data.iloc[start_row:end_row]
            
            # Save the new block to a new sheet in the same Excel file
            sheet_name = f"Sheet{x+1}"
            new_data.to_excel(writer, sheet_name=sheet_name, index=False)
            
    #print(f"Data has been split and saved to {output_filename}")
    return output_filename

class MyApp(wx.App):
    def OnInit(self):
        # Create a file dialog
        with wx.FileDialog(None, "Selecciona un archivo de Excel", wildcard="Excel files (*.xlsx;*.xls)|*.xlsx;*.xls",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return False  # The user cancelled the dialog

            # Get the selected file path
            pathname = fileDialog.GetPath()
            output_filename = data_breakup(pathname)  # Process the selected file

            # Show a message box with the result
            wx.MessageBox(f"Los datos se han dividido y guardado en {output_filename}", "Info", wx.OK | wx.ICON_INFORMATION)
        
        return True
    
    if platform.system() == 'Darwin':  # macOS
        def applicationSupportsSecureRestorableState(self):
            return True

def main():
    # Create an instance of the application
    app = MyApp(False)
    app.MainLoop()

# Run the main function
if __name__ == "__main__":
    main()
