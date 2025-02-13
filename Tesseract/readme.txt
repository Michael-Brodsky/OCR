This package contains a demonstration version of DOCUFILE_TS.accdb, a Microsoft Access 
application that reads and catalogs files (scanned documents, photos, etc.) in a
database. Source files must be convertible to plain text and contain a unique, human-
readable "key", such as a number, date, or string that uniquely identifies the file. 
DOCUFILE_TS reads each file according to user-defined rules, stores the information in 
a database and moves the file to a specified storage location. 

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
form, double-click in the "Processes" list or click the gear icon on the Admin form.
The form has several collapsible sections that specify where and what types of files 
to search for, the catalog database, file storage path, data extraction and storage 
rules, and parameters to convert source files into readable text.

The default settings are for the demo program. Changing these settings may prevent 
the application from successfully cataloging the sample source files.

The sample files resemble sales receipts. The demo extracts several pieces of data 
and stores them, along with the file's permanent location, in the sample table 
"tblCatalogDemo". Cataloged files can be viewed using the sample "Catalog" form 
described above.

NOTE: The application file, DOCUFILE_TS.accdb, is relocatable, but the folders  
installed with it are not. They must remain in the same location where the 
application was originally installed, otherwise the application may stop working.
 
