# Convert all to pdf
Note: The outputs are in excel-output/, pptx-output/, and word-output/ directories as they are processed in batch mode.

Instructions

1. Make a python environment and activate env  
> $ python -m venv env  

For windows:  
> $ . env/Scripts/activate - using gitbash  
> $ env\Scripts\activate.bat - using cmd  

For mac/linux:  
> $ . env/bin/activate  

2. Pip install libraries  

For excel to pdf:  
> $ pip install pywin32, comtypes  

For powerpoint to pdf:  
> $ pip install comtypes  

For word to pdf:  
> $ pip install docx2pdf  

To make executable:  
> $ pip install pyinstaller  


3. Convert to executable  
> $ pyinstaller --onefile main.py  

4. Run executable by opening cmd  
> $ start main.exe