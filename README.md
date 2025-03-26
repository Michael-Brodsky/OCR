# OCR
Optical Character Recognition Database File Storage Solutions (PaperShredder)

This repository contains several Microsoft Access database solutions for storing digital copies 
of paper documents in a database. The library and application databases can read photos and  
scanned documents, extract identifying information and store that information along with links 
to the source documents. Clients that access the database can provide users with means to search 
for and view the documents.

To read and extract information, the library databases convert source files to text using 
optical character recognition (OCR) technology. Currently, two OCR engines are used: OCR-
Tesseract and Microsoft OneNote. Tesseract is an open-source OCR engine and one of the most 
reliable and effective available. It is hosted on Github: https://github.com/tesseract-ocr/tesseract. 
MS OneNote is part of the MS Office suite and has OCR capabilities built in. Depending on source 
file format and quality, both technologies are nearly 100% successful at converting images to 
text.

Tesseract has the advantage of being much faster than OneNote, slightly more reliable at OCR 
conversion and, most importantly, can run silently in the background without disturbing the 
user's desktop. Its main disadvantage is that it can only convert a few image file formats 
into text, thus requiring installation of additional converters. The repository libraries use  
ImageMagick, an open-source image converter, to convert Adobe pdf files into a format readable 
by Tesseract. ImageMagick requires the support of GPL GhostScript, which must also be installed 
on the host computer.

OneNote has the advantage of being part of the MS Office suite and thus requires no further
licensing or installation of third-party software. It's main disadvantages are its speed and 
that it cannot run silently in the background, essentially taking over the user's desktop. This 
is especially troublesome when cataloging Adobe pdf files as Adobe has removed the silent 
printing option from Acrobat and so Acrobat and OneNote contiuously switch back and forth from 
being the top-level window, even if the user clicks on another window. This prevents users 
from performing any work while cataloging is in progress. One solution is to run the 
applications on different monitors in a multiple-monitor setup. This is only a "visual" 
improvement as OneNote and Adobe continually get the focus, including the mousepointer, and 
users still cannot perform any other tasks during the process.

Currently, the library and application databases have only been tested on Windows 10 and 11, 
64-bit implementations. The apps require that a setup procedure be run prior to use. The
setup checks for and installs any dependencies, creates a simple file system and configures 
the app to run in host computer's environment. The apps also come with sample source files 
and a demo program that reads and catalogs them. Contact the author for a self-extracting
.zip file that contains all necessary components to install and setup the applications.

Any database server that MS Access is able to establish an ODBC connection with can be 
used for back-end data storage, including other Access databases, SQL Server, Azure and
dBase. The demo program includes a simple local table that illustrates how files are 
cataloged and a form to view and manage the catalog.
