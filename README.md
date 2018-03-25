# OpenXLSX
OpenXLSX is a C++ library for reading, writing, creating and modyfing Microsoft Excel® files, with the .xlsx format.

## Motivation
Many programming languages have the ability to modify Excel files, either natively or in the form of open source libraries. This includes Python, Java and C#. For C++, however, things are more scattered. While there are some libraries, they are generally less mature and have a smaller featureset than for other languages. Here is a summary of the main C++ libraries for Excel files:

### libxls
The libxls library (https://sourceforge.net/projects/libxls/) is a C library for reading files in the legacy Excel file format, .xls. It cannot be used for writing or modifying Excel files.

### xlsib
The xlslib library (https://sourceforge.net/projects/xlslib/) is a C/C++ library for creating files in the legacy Excel file format, .xls. It cannot be used for reading or modifying Excel files.

### libxlsxwriter
The libxlsxwriter library (https://libxlsxwriter.github.io) is a C library for creating .xlsx files. It cannot be used for reading or modifying Excel files.

### LibXL
The LibXL library (http://www.libxl.com) can read, write, create and modify Excel files, in both the .xls and .xlsx formats. It is the most feature complete library available and has interfaces for C, C++, C# and Delphi. It is only available for purchase, however.

### QtXlsx
Of the open source libraries, the QtXlsx library (https://github.com/dbzhang800/QtXlsxWriter) is the most feature complete. It is, however, based on the Qt framework. While I'm a big fan of Qt for application programming purposes, I don't believe it is the best option for lower-level libraries.

Because there are no open source library that fully fitted my needs, I decided to develop the OpenXLSX library.

## Ambition
The ambition is that OpenXLSX should be able to read, write, create and modify Excel files (data as well as formatting), and do so with as few dependencies as possible. Currently, OpenXLSX depends on the following 3rd party libraries:

 - Boost
 - RapidXML (included)
 - LibZip (and, in turn, zlib)
 - LibZip++ (included)
 
 ## Current Status
 OpenXLSX is still work i progress. The following is a list of features which have been implemented and should be working properly:
 
  - Create/open/save files
  - Read/write/modify cell contents
  - Copy cells and cell ranges
  - Copy worksheets
  
  Features related to formatting, plots and figures have not yet been implemented
  
  ## License
  Use of OpenXLSX and the included third party libraries are granted under the following licenses:
  
  ### RapidXML (The MIT License)
  Copyright (c) 2006, 2007 Marcin Kalicinski
  
  Permission is hereby granted, free of charge, to any person obtaining a copy 
  of this software and associated documentation files (the "Software"), to deal 
  in the Software without restriction, including without limitation the rights 
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies 
  of the Software, and to permit persons to whom the Software is furnished to do so, 
  subject to the following conditions:
  
  The above copyright notice and this permission notice shall be included in all 
  copies or substantial portions of the Software.
  
  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
  THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
  IN THE SOFTWARE.
  
  ### LibZip++ (BSD License 2.0)
  Copyright (C) 2013 Cédric Tabin
  
  The author can be contacted on http://www.astorm.ch/blog/index.php?contact
  
  Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
  
  1. Redistributions of source code must retain the above copyright
  notice, this list of conditions and the following disclaimer.

  2. Redistributions in binary form must reproduce the above copyright
  notice, this list of conditions and the following disclaimer in the
  documentation and/or other materials provided with the distribution.
  
  3. The names of the authors may not be used to endorse or promote products
  derived from this software without specific prior written permission.
  
  THIS SOFTWARE IS PROVIDED BY THE AUTHORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
  
  ### OpenXLSX (BSD License 2.0)
  Copyright (c) 2018, Kenneth Troldal Balslev
  
  All rights reserved.
  
  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
        notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
        notice, this list of conditions and the following disclaimer in the
        documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
        names of any contributors may be used to endorse or promote products
        derived from this software without specific prior written permission.
  
  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.