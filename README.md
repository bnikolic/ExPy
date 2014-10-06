ExPy
====

In-process Python for use within Excel.

Currently tested against Python 2.7 and Excel 32bit. Works with, for
example, the Windows 32 bit Anaconda distribution.

BSD-3 type license, see LICENSE.

For more information see
http://www.bnikolic.co.uk/expy/expy.html. Build of current latest
version is available at: http://www.bnikolic.co.uk/soft/ExPy-0.4.zip

Inquiries to webs@bnikolic.co.uk

Description
-----------

ExPy allows the use of Python in Excel by loading Python into the same
process as excel and exposing function using the Excel Addin
interface. This is a relatively high-performance and robust interface
mechanism. ExcelDNA works in a similar way (providing interfacing to
.Net languages) as well many other application-specific Excel addins.

Alternatives ways of interfacing are via COM (e.g., excelpython,
https://github.com/ericremoreynolds/excelpython ) or client/server
architectures (e.g., I believe xlloop).

Installation - Binaries
-----------------------

Binaries only work against Python 2.7.x versions 32bit and Excel 32
bit.

1. Ensure you have installed Python version 2.7.x and it is in the
   PATH environment variable. The recommended way of doing this is by
   installing the Anaconda distribution: http://continuum.io/downloads

2. Unpack the zip file

3. Open the expy.xll file and confirm you want to give it privileges
   to run










