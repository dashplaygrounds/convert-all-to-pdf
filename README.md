# Convert all to pdf

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
> $ pip install win32com, comtypes  

For powerpoint to pdf:  
> $ pip install comtypes  

For word to pdf:  
> $ pip install docx2pdf  

To make executable:  
> $ pip install pyinstaller  


3. Convert to executable  
> $ pyinstaller --onefile main.py  
