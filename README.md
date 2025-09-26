📌 Insanity CLI Tool  

Insanity CLI is a Python command-line tool to explore, browse, and export hierarchical data from Excel workbooks.  

It works with multiple sheets, treating each sheet as a main category, and automatically detects columns for Subcategory, Sub-subcategory, Item (or column X), and Client (or column Y).  

You can navigate categories interactively or export selected data directly to a new Excel file.  

🛠️ Requirements  

- Python 3.x installed.  

- Virtual environment recommended:  

  - python3 -m venv venv  

  - source venv/bin/activate # macOS/Linux  

- Packages: pandas, openpyxl (install with pip install pandas openpyxl).  

- Place insanity.py and your Excel file(s) on the Desktop for easy access.  

🏃 Running the Program  

Open a terminal and navigate to the Desktop (cd ~/Desktop).  

Run the script:  

python3 insanity.py  

Enter the path to your Excel file when prompted, e.g.:  

~/Desktop/dummy_categories.xlsx  

💻 Commands  

- insanity → Type this to list all main categories (sheet names).  

- Type a main category → Shows its subcategories.  

- Type a subcategory → Shows its sub-subcategories (if any).  

- Type a sub-subcategory → Shows all items in that sub-subcategory.  

- fix the insanity → Type this to export a list of categories, subcategories, or sub-subcategories to a new Excel file.  

- back → Return to the previous level while browsing.  

- bye → Exit the program gracefully.  

⚠️ Important: Do not type fix the insanity while navigating a category; this won't work. Return to the main menu using back or "insanity".  

📤 Exporting Items  

- After typing fix the insanity, enter a comma-separated list of categories, subcategories, or sub-subcategories.  

- The program will export all matching items to insane_workbook.xlsx.  

- Columns are auto-detected and renamed: Title → Item, Authors → Client.  

- Always type names exactly as they appear in your Excel file.  

- After exporting or finishing browsing, type bye to exit.  
