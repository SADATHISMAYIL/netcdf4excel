# NetCDF add-in written in Visual Basic for MS Excel #

## Context ##

NetCDF (network Common Data Form) is a set of software libraries and machine-independent data formats that support the creation, access, and sharing of array-oriented scientific data. NetCDF is provided by Unidata.

## Objective ##

This web page intends to simplify the use of NetCDF data in Excel, providing a ready to use solution for manipulating this type of data.

For developers, the open-source (GPL V3 license) can be downloaded directly or checked out with Mercurial.

## The NetCDF add-in ##

This add-in is written in VBA 6.0 (so it won't work with Office 2010 64 bits) and is designed for Excel 2007 running with the Microsoft Windows operating system.

It allows to open NetCDF data (classical format, versions 3.x) with Excel for read / write access. Nevertheless, please remind that you will always be limited by Excel's own memory limitations, i.e. trying to open huge data files may result in crashing Excel...

### Supported structures and Types ###

Supported data structures
  * Values
  * Vectors with dimension length up to 1,048,576
  * Numerical Matrices with up to N dimensions : at least dimension N (or N-1) should be less than 16,384 AND the product of the other dimensions should be less than 1,048,576 divided by dimension N (or N-1)
  * Character Matrices with up to 4 dimensions (the last one has to be the - constant - number of characters of the strings)


Supported data types
  * NC\_BYTE (the equivalent of the C type « signed char »)
  * NC\_SHORT
  * NC\_INT
  * NC\_FLOAT
  * NC\_DOUBLE
  * NC\_CHAR (be careful : NetCDF character strings are stored in fixed length vectors)

Supported platforms
  * Windows XP 32 bits
  * Windows Vista 32 bits
  * Windows Seven 32 and 64 bits
  * Windows 8 32 and 64 bits

## Installing the software ##

Just download the add-in, execute the installer, accept the license agreement and that's all. You should then have a short-cut on your desktop allowing you to launch an empty Excel workbook including the add-in.


### Reading an entire NetCDF file ###

In the NetCDF menu, select « Open File » , then browse to the desired nc file. DON'T TRY TO USE THE NATIVE EXCEL "OPEN FILE" MENU, IT JUST WON'T WORK!

### Reading a specific variable or a slice of variable in a NetCDF file ###

In the NetCDF menu, select « Read Variable » , browse to the desired nc file and select the variable you want to read. It is possible to select a slice of data to be read.

### Writing a NetCDF file ###

Important notice : if you make changes to the data in Excel, you should be careful to strictly respect the structure of the Excel workbook, as described in the "NC\_Info" tab.

In the NetCDF menu, select « Save File » , then browse to the desired location.


### Known limitations and problems ###

  * Rounding errors may occur with data stored as floating point values.
  * Character strings : in Excel, strings are written in the cells of the spreadsheet with double quotes and are of a fixed length (strings are completed with blanks if necessary).


### Download NetCDF4Excel ###

Due to Google Code limitations on downloads, the Excel add-in is hosted on an external server (see link on the left side of this page).


### Contributors ###

After a few years alone on this project, I'm now looking for a few good men and women to help me with maintaining this add-in and porting it to other operating systems and Excel versions.