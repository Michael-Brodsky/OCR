This package contains a demonstration version of DOCUFILE_TS.accdb, a Microsoft Access 
application that reads and catalogs files (scanned documents, photos, etc.) in a
database. Source files must be convertible to plain text and contain a unique, human-
readable "key", such as a number, date, or string that uniquely identifies the file. 
DOCUFILE_TS reads each file according to user-defined rules, stores the information in 
a database and moves the file to a specified permanent storage location. 

To install and run DOCUFILE_TS, follow this procedure:

	1) EXTRACT: 
		Extract the contents of the DOCUFILE_TS.zip folder to your computer.
	2) OPEN: 
		Open DOCUFILE_TS.accdb. The application may require you to enable  
	   	active content, depending on your Microsoft Trust Center settings. If 
	   	so, you'll be prompted to "Enable active content and run Setup." 
	   	First, close this dialog then enable active content by clicking "Open" 
		and/or "Enable" in the Microsoft Trust Center dialogs.
	4) SETUP: 
		The application will prompt you to run the Setup procedure. If 
	   	required, this will download and install several dependencies (see 
	   	dependencies\readme.txt), configure it to run in the current 
	   	environment and setup the demo program. The installer requires an 
	   	active internet connection.
	5) RUN: 
		The demo program reads pdf files from the \samples folder, catalogs 
	   	them and moves any successful files to the \catalog folder. To run the 
	   	demo, open the "Admin" form in the Navigation Pane and click the green 
	   	"Play" button. The application displays status messages as the program 
	   	runs and how many files were successfully cataloged when finished. The 
		program can be stopped with the red "Stop" button.
	6) VIEW: 
		Cataloged files can be viewed by opening the "Catalog" form. The form 
	   	has several fields showing data extracted from the source files. The 
	   	"Transaction" field is a hyperlink that opens the cataloged file in 
	   	its associated application.

Process settings can be viewed and maintained using the "Settings" form. To open the 
form, double-click in the "Processes" list of the Admin form. The Settings form has 
four collapsible sections:

Source Files - contains information on where and what types of files to search for.

Catalog - contains information about the catalog database, where to permanently 
	  store source files and how to store the file path.

OCR Data Rules - lists user-defined rules that specify what and how to extract 
	         data from source files and store it in the catalog.

File Conversion - contains the OCR engine parameters used to convert source files 
		- into readable text so they can be cataloged.

The default settings are for the demo program. Changing these settings may prevent 
the application from successfully cataloging the sample files.

The demo comes with a sample table "tblCatalogDemo", and form "Catalog", for storing
and viewing cataloged files. These can be changed in the "Catalog" section of the 
Settings form, allowing the user to select another database and table. To select an
external database, enter the its full path or connection string. To use the current
database, leave it blank. When a database is selected, the "SaveTo Table" drop-down 
will list the available database tables. Once a table is selected, the 
"SaveTo Field" drop-down will list the available fields that can be used to store 
the cataloged file's permanent location, specified in the "SaveTo Path" setting 
above it. If "SaveTo Path" is left blank, cataloged source files will remain in the 
same location they were found. Regardless, the "SaveTo Table" and "SaveTo Field" 
must be specified. If a different database and/or table are used, the "Catalog" 
form will have to be redesigned or replaced with one compatible with the new 
database and table.

The sample source files resemble sales receipts. The demo extracts several pieces of
information and stores the information, along with the file's permanent location in 
the catalog table. The "OCR Data Rules" section of the Settings form lists the rules
used to extract and store data from source files. They can be edited by double-
clicking the list. The rules use Regular Expressions and pattern matching logic to 
extract data and specify where and how to store it.

NOTE: The application file, DOCUFILE_TS.accdb, is relocatable, but the folders  
installed with it are not. They must remain in the same location where the 
applicationwas originally installed, otherwise the application may stop working.
 
