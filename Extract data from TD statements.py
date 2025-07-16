#region imports and constants
import pdfplumber
import pandas as pd
import pathlib
import tkinter as tk
from tkinter import filedialog
root = tk.Tk()
root.withdraw()  # Hide the root window
source_folder = filedialog.askdirectory(title="Select the folder where the PDF statement files are located")
STATEMENT_DIR = pathlib.Path(source_folder)
pathlib.Path.mkdir(STATEMENT_DIR / "parsed", exist_ok=True)
OUTPUT_PATH = pathlib.Path(STATEMENT_DIR.__str__() + "/parsed")
TABLE_COLUMNS = ["Description", "Withdrawals", "Deposits", "Date", "Balance", "Overflow"]
DEBUG_ON = True
# Adjust the EXPLICIT_LINES constant to match the vertical lines in the PDF.
# The debug_tablefinder will output images that show the lines when you run the script with DEBUG_ON set to True.
EXPLICIT_LINES = [58,195,300,400,432,540]
#endregion
for f in STATEMENT_DIR.glob("TD*.pdf"):
    doc = f
    pp = pdfplumber.open(doc)
    for pg in pp.pages:
        # dicttext = pg.extract_text_lines(layout=False, strip=True, return_chars=True)
        startingbalance_line = pg.search(r'STARTING *BALANCE', regex=True, case=True)
        balanceforward_heading = pg.search(r'BALANCE *FORWARD',regex=True,case=True)
        description_heading = pg.search("DESCRIPTION", regex=False, case=False)
        closingbalance_line = pg.search(r'CLOSING *BALANCE', regex=True,case=True)
        top_of_bottom_table = pg.search(r'Account */ *Transaction *Type', regex=True,case=False)
        if startingbalance_line or balanceforward_heading or description_heading or closingbalance_line:
            if top_of_bottom_table:
                bottom = top_of_bottom_table[0].get('bottom') - 25
            else:
                bottom = pg.height
            '''
            if closingbalance_line:
                bottom = closingbalance_line[0].get('bottom')
            else:
                bottom = pg.height
            '''
            top = description_heading[0].get('top')
            top = top - 10
            if startingbalance_line:
                x0 = startingbalance_line[0].get('x0')
            else:
                x0 = balanceforward_heading[0].get('x0')
            x0 = x0 - 5
            x1 = pg.width
            if DEBUG_ON == True:
                croppedpage = pg.crop((0 ,top, pg.width, bottom))
                croppedpageimage = croppedpage.to_image(resolution=150)
                # croppedpageimage.show()
                croppedpageimage.save(OUTPUT_PATH.__str__() + "/"+ f.stem + pg.page_number.__str__() + "crop_image.png")
                debugtable = croppedpageimage.debug_tablefinder(({
                    "vertical_strategy": "explicit", 
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines":EXPLICIT_LINES}))
                debugtable.save(OUTPUT_PATH.__str__() + "/"+ f.stem + pg.page_number.__str__() + "debug_table.png")
            try:
                if len(tablelist) != 0:
                    moredata = croppedpage.extract_table({
                    "vertical_strategy": "explicit", 
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines":EXPLICIT_LINES})
                    tablelist.extend(moredata)
                else:
                    tablelist = croppedpage.extract_table({
                    "vertical_strategy": "explicit", 
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines":EXPLICIT_LINES})    
            except:
                tablelist = croppedpage.extract_table({
                "vertical_strategy": "explicit", 
                "horizontal_strategy": "text",
                "explicit_vertical_lines":EXPLICIT_LINES})
        #if page has a relevant table ^
    #for each page
    data_dict = {i: row for i, row in enumerate(tablelist)}
    df = pd.DataFrame.from_dict(data_dict, orient='index')
    # df = pd.DataFrame(tablelist[1:], columns=TABLE_COLUMNS, engine='python')
    tablelist.clear()
    df.to_excel(OUTPUT_PATH.__str__() + "/parsed" + f.stem + ".xlsx")
#for each file
