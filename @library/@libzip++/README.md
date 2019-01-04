
libzippp
--------

libzippp is a simple basic C++ wrapper around the libzip library.
It is meant to be a portable and easy-to-use library for ZIP handling.

Compilation has been tested with:
- GCC 4.9.3 (GNU/Linux Gentoo)
- GCC 6.2.1 (GNU/Linux Debian) 
- GCC 6.3.0 (GNU/Linux Debian) 
- GCC 7.2.0 (GNU/Linux Debian) 
- MS Visual Studio 2012 (Windows 7)

Underlying libraries:
- [ZLib](http://zlib.net) 1.2.11
- [libzip](http://www.nih.at/libzip) 1.5.1

# Compilation

### LINUX

0. Make sure you have the following commands: ```g++``` ```make``` ```tar``` ```wget```
1. Download and compile the libraries (zlib and libzip) with the command: ```make libraries```
2. Then create the static and shared libraries of libzippp: ```make```
3. You may want to run the tests (optional): ```make tests```
4. Now you just have to include the src folder in your include path and
  link against *libzippp.a* or *libzippp.so* (do not forget to also link 
  against libzip libraries in *lib/libzip-1.5.1/lib/.libs/*).
  An example of compilation with g++:
  
```Shell
g++ -I./lib/libzip-1.5.1/lib -I./src \
    main.cpp libzippp.a \
    lib/libzip-1.5.1/lib/.libs/libzip.a \
    lib/zlib-1.2.11/libz.a
```

#### Encryption

Since version 1.5, libzip uses an underlying cryptographic library (OpenSSL, GNUTLS or CommonCrypto) that
is necessary for static compilation. By default, libzippp will use `-lssl -lcrypto` (OpenSSL) as default flags
to compile the tests. This can be changed by using `make CRYPTO_FLAGS=-lsome_lib tests`

See [here](https://github.com/nih-at/libzip/blob/master/INSTALL.md) for more information.

### WINDOWS

#### From Stage 1 - Use prepared environment


0. Make sure you have cmake 3.10 (*cmake.exe* must be in the PATH) and MS Visual 
  Studio 2012. The dev command prompt path (defined in *compile.bat*) should be:
  ```
  <MSVS11>\Common7\Tools\VsDevCmd.bat
  ```

1. Download the *libzip-\<version\>-windows-ready_to_compile.zip* file from the release 
  and extract it somewhere on your system. This will create a prepared structure, so *libzippp* can 
  be compiled along with the needed libraries.

2. Simply execute the *compile.bat* file. This will compile *zlib*, *libzip* and
 finally *libzippp*.

3. You'll have a *dist* folder containing the *release* and *debug* folders 
  where you can now execute the libzippp tests.

#### From Stage 0 - DIY

0. Make sure you have cmake 3.10 (*cmake.exe* must be in the PATH) and MS Visual 
  Studio 2012. The dev command prompt path (defined in *compile.bat*) should be:
  ```
  <MSVS11>\Common7\Tools\VsDevCmd.bat
  ```
  
1. Download [libzip](http://www.nih.at/libzip/libzip-1.5.1.tar.gz) and [zlib](http://zlib.net/zlib1211.zip) sources and extract them in the 'lib' folder.
  You should end up with the following structure:
  ```
  libzippp/compile.bat
  libzippp/lib/zlib-1.2.11
  libzippp/lib/libzip-1.5.1
  ```

2. Execute the *compile.bat* (simply double-click on it). The compilation should 
  go without error.

3. You'll have a *dist* folder containing the *release* and *debug* folders 
  where you can now execute the libzippp tests.

4. You can either use *libzippp.dll* and *libzippp.lib* to link dynamically the 
  library or simply use *libzippp_static.lib* to link it statically. Unless you 
  also link zlib and libzippp statically, you'll need the dll packaged with 
  your executable.

# Usage 

The API is meant to be very straight forward. Some french explanations
can be found [here](http://www.astorm.ch/blog).

How to list and read files in an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

ZipArchive zf("archive.zip");
zf.open(ZipArchive::READ_ONLY);

vector<ZipEntry> entries = zf.getEntries();
vector<ZipEntry>::iterator it;
for(it=entries.begin() ; it!=entries.end(); ++it) {
  ZipEntry entry = *it;
  string name = entry.getName();
  int size = entry.getSize();

  //the length of binaryData will be size
  void* binaryData = entry.readAsBinary();

  //the length of textData will be size
  string textData = entry.readAsText();

  //...
}

zf.close();
```

How to read a specific entry from an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

ZipArchive zf("archive.zip");
zf.open(ZipArchive::READ_ONLY);

//raw access
char* data = (char*)zf.readEntry("myFile.txt", true);
ZipEntry entry1 = zf.getEntry("myFile.txt");
string str1(data, entry1.getSize());

//text access
ZipEntry entry2 = zf.getEntry("myFile.txt");
string str2 = entry2.readAsText();

zf.close();
```

How to read a large entry from an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

ZipArchive zf("archive.zip");
zf.open(ZipArchive::READ_ONLY);

ZipEntry largeEntry = z1.getEntry("largeentry");
std::ofstream ofUnzippedFile("largeFileContent.data");
largeEntry.readContent(ofUnzippedFile);
ofUnzippedFile.close();

zf.close();
```

How to add data to an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

ZipArchive zf("archive.zip");
zf.open(ZipArchive::WRITE);
zf.addEntry("folder/subdir/");

const char* textData = "Hello,World!";
zf.addData("helloworld.txt", textData, 12);

zf.close();
```

How to remove data from an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

ZipArchive zf("archive.zip");
zf.open(ZipArchive::WRITE);
zf.deleteEntry("myFile.txt");
zf.deleteEntry("myDir/subDir/");
zf.close();
```

# Issues

### LINUX

You might already have libzip compiled elsewhere on your system. Hence, you
don't need to run 'make libzip'. Instead, just put the libzip location when
you compile libzipp:

```Shell
make LIBZIP=path/to/libzip
```

Under Debian, you'll have to install the package *zlib1g-dev* in order to compile
if you don't want to install zlib manually.

### WINDOWS

By default, MS Visual Studio 2012 is installed under the following path:

```
C:\Program Files (x86)\Microsoft Visual Studio 11.0\
```

Be aware that non-virtual-only classes are shared within the DLL of libzippp.
Hence you'll need to use the same compiler for libzippp and the pieces of code
that will use it. To avoid this issue, you'll have to link the library statically.

More information [here](http://www.codeproject.com/Articles/28969/HowTo-Export-C-classes-from-a-DLL).

