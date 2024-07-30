Invoice Generation

```~~~~~~~~~~
	- Invoice generation is a python project which generates the invoices based on the records in a excel file.
	- In this porject we are generating invoices based on the recordes in a excel file.
	- The project reads the excel file and a PDF in two different folders and mearge them as one pdf document.

Prerequisites:
	1) Ananconda with python 3.9+ versions (You can download from here https://www.anaconda.com/distribution/#download-section)
	2) Visual C++ Redistributable for Visual Studio 2015 (You can download from here https://www.microsoft.com/en-in/download/details.aspx?id=48145 or search in google as follow "Visual C++ Redistributable for Visual Studio 2015")
	3) Any code editors (Notepad++ or Visual Studio Code is recommanded)

Installation:
	1)Download wkhtmltopdf.exe VERSION 0.12.6 file form https://wkhtmltopdf.org/downloads.html based on your system Architectures.

		OS/Distribution				Supported on				Architectures
		Windows				Installer (Vista or later)			64-bit	32-bit
							7z Archive (XP/2003 or later)		64-bit	32-bit

	2) Find the downloaded file could be named as "wkhtmltox-0.12.6-1.msvc2015-win64.exe".
	3) Double click on the ".exe" file to install.
	4) Note the installed .exe path.
		Example :- C:\Program Files\wkhtmltopdf\bin
	5)Copy the path and repalce it in line number 786 of app.py which is as below:

		Replace the "pdfkit_html_pdf_config_filepath" variable value with the file path of yours.

		pdfkit_html_pdf_config_filepath = <'File path'>(should be in single quotes)

		Example:
		pdfkit_html_pdf_config_filepath = 'C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'


Steup:
	1) Project file:
		- Obtain the project folder.
	2) Anaconda Environment setup:
		step-1: Open Anaconda prompt from start menu.
		step-2: Execute the below command in anaconda prompt to create Anaconda Virtual enivronment.
			    Syntax: conda create -n ENV_NAME python=3.9
				Note: Replace ENV_NAME with your preferred name to create virtual environment.
		step-3: Activate the Anaconda virtual environment.
				Syntax: conda activate ENV_NAME
				Note: Replace ENV_NAME with the name you have given while creating anaconda environment
	3) Package Installation:
		Dependent Python package installations.(The direct Installation of dependency package through pip or conda may fail some times, so we have mentioned all possible methods of installation procedures)
		- You can install the dependency packages using "pip" or "conda" commands through anaconda prompt.
		- Recommanded to use pip to install dependencies.
		- Execute the below sytax to install following packages.
		  (Installation of pathlib, pandas,pdfkit,argparse,tqdm,num2words and PyPDF2 packages)
		pip:
			syntax: pip install pathlib==1.0.1 pandas==2.0.1 PyPDF2==3.0.1 pdfkit==1.0.0 argparse==1.4.0 tqdm==4.65.0 num2words==0.5.12
   	Note:
		1) All the mention syntax/commands should be executed in Anaconda Prompt only.
		2) In case, while execution if any syntax/command fails, we need to start the installation procedure from the beginning.

Usage:

	Method-1:

		1) Move to the project directory.
			syntax: cd <Project parent folder> (Place the complete path till project folder)
		2) Activate Anaconda environment in anaconda prompt (Neglect this step if you already activated anaconda environment)
			syntax: conda activate <ENV_Name>
		2) Opening VSCode.
			syntax: code .
		3) Run the project.
			syntax:
				(i) python app.py -f "<fileName.xlsx>" -l "<fileName.pdf>"
										(or)
			    (ii) python app.py -filename "<fileName.xlsx>" -lut "<fileName.pdf>"

			Example: python app.py -f "billz.xlsx" -l "LUT Copies_2023-24.pdf"


Contributing:
	Any contribution to the project is welcome.

License:
	__authors__ = "OTSi AI-Ml TEAM"
	__version__ = "1.0.0"
	__copyright__ = "Copyright (c) 2023-2024 Object Technology Solutions India Private Ltd"
	__license__ = "Enterprise Edition"
```
"# invoice_generator" 
