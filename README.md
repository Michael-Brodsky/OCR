# OCR
Optical Character Recognition Database File Storage Solutions

This repository contains several Microsoft Access database solutions for reading and cataloging
electronic documents in a database. Documents, such as scans and photos are converted to OCR 
text, parsed for a file "key" (a unique string, number or datetime that identifies the file) 
and stored in a database. The repository includes libraries that perform the work and simple 
user interfaces to demonstrate their use.

The libraries use either Tesseract or Microsoft OneNote as the OCR engine. Tesseract is an 
open-source OCR engine and one of the most reliable and effective available. It is hosted on 
Github: https://github.com/tesseract-ocr/tesseract. MS OneNote is part of the MS Office suite 
and has OCR conversion capabilities built in. Depending on source file format and quality 
OneNote is also a highly effective OCR engine.

Tesseract has the advantage of being slightly more reliable at OCR conversion and, most 
importantly, can run silently in the background without disturbing the user's desktop. Its
main disadvantage is that it can only convert a few image file formats into OCR text, thus 
requiring installation of additional converters. The repository libraries use ImageMagick,
an open-source image converter, to convert Adobe pdf files into a format readable by 
Tesseract. ImageMagick however requires the support of GPL GhostScript, which must also be
installed on the host computer.

OneNote has the advantage of being part of the MS Office suite and thus requires no further
licensing or installation of third-party software. It's main disadvantage is that it cannot
run silently and takes over the user's desktop. This is especially troublesome when cataloging
Adobe pdf files as Adobe has removed the silent printing option from Acrobat and so Acrobat
and OneNote contiuously switch back and forth from being the top-level window, even if the
user clicks on another window. This prevents users from performing any work while the 
cataloging process is in progress. One solution is to run the applications on different 
monitors in a multiple-monitor setup. This however is only a "visual" improvement as OneNote 
and Adobe continually get the focus, including the mousepointer, and users are still not 
able to perform any other tasks while the process runs.

Currently, the libraries and UI applications have only been tested on Windows 10 and 11, 64
bit implementations. The UI apps require that a setup procedure be run prior to use. The
setup checks for and installs any dependencies, creates a simple file system and configures 
the app to run in host computer's environment. The UI apps also come with sample source 
files and a demo program that reads and catalogs them.
