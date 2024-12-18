This package contains a demonstration version of DOCUFILE_TS.accdb, a Microsoft Access 
application that can read and catalog image files (scanned documents, photos, etc.) in 
a database. Image files must contain a unique, readable "key", such as a receipt or 
invoice number, person's name, date or any other human-readable information that 
uniquely identifies the file. DOCUFILE_TS reads each file according to user-defined 
rules and, if a "key" is found, stores the information in a database and moves the
file to a specified location for permanent storage.

To install and run DOCUFILE_TS, follow this procedure:

	1) Extract the contents of the DOCUFILE_TS.zip folder to your computer.
	2) Open DOCUFILE_TS.accdb. The application may require you to enable active 
	   content, depending on your Microsoft Trust Center settings. The application
	   will prompt you to "Enable active content and run Setup when prompted." 
	   This message window should first be closed and then active content enable
	   by clicking "Open" and/or "Enable" in the Microsoft Trust Center message 
	   dialogs.
	4) The application will then prompt you to run the Setup procedure. If 
	   required, this will download and install several dependencies (see 
	   dependencies\readme.txt), configure it to run in the current environment 
	   and setup the demo program.
	5) The demo program reads pdf files from the \samples folder, catalogs them in
	   a table and moves any successful files to the \catalog folder. To run the	
	   demo, open the "Admin" form in the Navigation Pane and click the green 
	   "Play" button. The application will display status messages as the program
	   runs and how many files were successfully cataloged when finished. The 
	   demo can be stopped with the red "Stop" button.
	6) Cataloged files can be viewed by opening the "Catalog" form. This simple
	   form contains two fields, "ID" and "Catalog Date". The ID field is a 
	   hyperlink that will open the cataloged file in its associated application.

Program settings can be viewed and maintained using the "Settings" form. To open the 
form, click the "Tools" button at the bottom right corner of the "Source Files" box 
in the Admin window. Another Tools button on the Settings form toggles between Basic
and Advanced settings. The default settings are for the demo program. Changing these
settings may prevent the program from successfully cataloging the sample files. 
(The program will still run, but may not be able to read or catalog the files).

The demo comes with a sample table "tblCatalogDemo", and form "Catalog", for storing
and viewing cataloged files. These can be changed in the "Catalog" section of the 
Advanced Settings form, allowing the user to select another database and table. To 
select an external database, enter the full path to the database file in the 
"Database" box. To use the current database, leave it blank. When a database is 
selected, the "Table" drop-down will list the available tables in that database. 
Once a table is selected, the "Key Field" and "File Field" drop-downs will list the 
available fields that can be used to store the file's "key" and path to its storage 
location. The Key Field must be of an appropriate data type to store key values, 
usually Text, Long Integer or Date, depending on the key's format. A Text field 
should be able to store any key since its value comes from the OCR text found in the 
source file and is already in the proper format. The File Field must be of type
Text, as it stores a file path as a text string. Catalog tables may have any number
of fields, but the application requires at a minimum the Key and File fields to
function.

NOTE: The application file, DOCUFILE_TS.accdb, is relocatable but the folders  
installed with it are not and must remain in the same location where the application
was originally installed, otherwise the application may stop working.
 
