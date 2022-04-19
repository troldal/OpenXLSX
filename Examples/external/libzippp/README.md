[![Build Status](https://travis-ci.com/ctabin/libzippp.svg?branch=master)](https://travis-ci.com/ctabin/libzippp)

# LIBZIPPP

libzippp is a simple basic C++ wrapper around the libzip library.
It is meant to be a portable and easy-to-use library for ZIP handling.

Compilation has been tested with:
- GCC 8 (Travis CI)
- GCC 10.2.1 (GNU/Linux Debian) 
- MS Visual Studio 2012 (Windows 7)

Underlying libraries:
- [ZLib](http://zlib.net) 1.2.11
- [libzip](http://www.nih.at/libzip) 1.8.0
- [BZip2](https://www.sourceware.org/bzip2/) 1.0.8 (optional)

## Integration

libzippp has been ported  to [vcpkg](https://github.com/microsoft/vcpkg) and thus can be
very easily integrated by running:
```
./vcpkg install libzippp
```

## Compilation

### Install Prerequisites

This library requires at least C++ 11 to be compiled.

- Linux
  - Install the development packages for zlib and libzip (e.g. `zlib1g-dev`, `libzip-dev`, `liblzma-dev`, `libbz2-dev`).
  - It is possible to use the Makefile by executing `make libraries`.

- Windows:
  - Use precompiled libraries from *libzippp-\<version\>-windows-ready_to_compile.zip*.
  - Install from source via CMake (similar to workflow below).

- All Operating systems
  - If it is intended to be used with encryption it is necessary to compile libzip with any encryption and to enable it in libzippp through the cmake flag `LIBZIPPP_ENABLE_ENCRYPTION`.

### Compile libzippp

#### Quick start

```sh
mkdir build
cd build
cmake ..
make
make install
```

#### Step by step

- Make sure you have a compiler (MSVC, g++, ...) and CMake installed
- Switch to the source folder
- Create a build folder and switch to it, e.g.: `mkdir build && cd build`
- Configure the build with cmake:
  - Commandline: `cmake .. -DCMAKE_BUILD_TYPE=Release`
  - Within the CMake GUI, set source and build folder accordingly
	- Click `Add Cache Entry` to add `CMAKE_BUILD_TYPE` if not building with MSVC
	- Click `Configure` & `Generate`
  - If CMake can't find zlib and/or libzip you need to set `CMAKE_PREFIX_PATH` to the directories where you installed those into
  (either via `-DCMAKE_PREFIX_PATH=<...>` or via the GUI)
    - Example: `-DCMAKE_PREFIX_PATH=/home/user/libzip-1.8.0:/home/user/zlib-1.2.11`
- Compile as usual
  - Linux: `make && make install`
  - Windows: Open generated project in MSVC. Build the `INSTALL` target to install.

#### CMake variables of interest

Set via commandline as `cmake -DNAME=VALUE <other opts>` or via CMake GUI or CCMake `Add Cache Entry`.

- `LIBZIPPP_INSTALL`: Enable/Disable installation of libzippp. Default is OFF when using via `add_subdirectory`, else ON
- `LIBZIPPP_INSTALL_HEADERS`: Enable/Disable installation of libzippp headers. Default is OFF when using via `add_subdirectory`, else ON
- `LIBZIPPP_BUILD_TESTS`: Enable/Disable building libzippp tests. Default is OFF when using via `add_subdirectory`, else ON
- `LIBZIPPP_ENABLE_ENCRYPTION`: Enable/Disable building libzippp with encryption capabilities. Default is OFF.
- `CMAKE_INSTALL_PREFIX`: Where to install the project to
- `CMAKE_BUILD_TYPE`: Set to Release or Debug to build with or without optimizations
- `CMAKE_PREFIX_PATH`: Colon-separated list of prefix paths (paths containing `lib` and `include` folders) for installed libs to be used by this
- `BUILD_SHARED_LIBS`: Set to ON or OFF to build shared or static libs, uses platform default if not set

### Referencing libzippp

Once installed libzippp can be used from any CMake project with ease:   
Given that it was installed (via `CMAKE_INSTALL_PREFIX`) into a standard location or its install prefix is passed into your projects
`CMAKE_PREFIX_PATH` you can simply call `find_package(libzippp 3.0 REQUIRED)` and link against `libzippp::libzippp`.

When not using CMake to consume libzippp you have to pass its include directory to your compiler and link against `libzippp.{a,so}`.
Do not forget to also link against libzip libraries e.g. in *lib/libzip-1.8.0/lib/.libs/*).
An example of compilation with g++:
  
```shell
g++ -I./src \
    -I./lib/libzip-1.8.0/lib I./lib/libzip-1.8.0/build \
    main.cpp libzippp.a \
    lib/libzip-1.8.0/lib/.libs/libzip.a \
    lib/zlib-1.2.11/libz.a
```

### Encryption

Since version 1.5, libzip uses an underlying cryptographic library (OpenSSL, GNUTLS or CommonCrypto) that
is necessary for static compilation. By default, libzippp will use `-lssl -lcrypto` (OpenSSL) as default flags
to compile the tests. This can be changed by using `make CRYPTO_FLAGS="-lsome_lib" LIBZIP_CMAKE="" tests`.

Since libzip `cmake`'s file detects automatically the cryptographic library to use, by default all the allowed
libraries but OpenSSL are explicitely disabled in the `LIBZIP_CMAKE` variable in the Makefile.

See [here](https://github.com/nih-at/libzip/blob/master/INSTALL.md) for more information.

### WINDOWS - Alternative way

The easiest way is to download zlib, libzip and libzippp sources and use CMake GUI to build each library in order:

- Open CMake GUI
- Point `Source` to the libraries source folder, `Build` to a new folder `build` inside it
- Run `Generate`
- Open the generated solution in MSVC and build & install it
- Repeat for the next library

But there is also a prepared batch file to help automate this.
It may need some adjusting though.

#### From Stage 1 - Use prepared environment

0. Make sure you have cmake 3.20 (*cmake.exe* must be in the PATH) and MS Visual Studio.

1. Download the *libzippp-\<version\>-windows-ready_to_compile.zip* file from the release 
  and extract it somewhere on your system. This will create a prepared structure, so *libzippp* can 
  be compiled along with the needed libraries.

2. Simply execute the *compile.bat* file. This will compile *zlib*, *libzip* and
 finally *libzippp*.

3. You'll have a *dist* folder containing the *release* and *debug* folders 
  where you can now execute the libzippp tests.

#### From Stage 0 - DIY

0. Make sure you have cmake 3.10 (*cmake.exe* must be in the PATH) and MS Visual Studio 2012.
  
1. Download [libzip](http://www.nih.at/libzip/libzip-1.8.0.tar.gz) and [zlib](http://zlib.net/zlib1211.zip) sources and extract them in the 'lib' folder.
  You should end up with the following structure:
  ```
  libzippp/compile.bat
  libzippp/lib/zlib-1.2.11
  libzippp/lib/libzip-1.8.0
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

### List and read files in an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

int main(int argc, char** argv) {
  ZipArchive zf("archive.zip");
  zf.open(ZipArchive::ReadOnly);

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
  
  return 0;
}
```

You can also create an archive directly from a buffer:
```C++
#include "libzippp.h"
using namespace libzippp;

int main(int argc, char** argv) {
  char* buffer = someData;
  uint32_t bufferSize = sizeOfBuffer;

  ZipArchive* zf = ZipArchive::fromBuffer(buffer, bufferSize);
  if(zf!=nullptr) {
    /* work with zf */
    zf->close();
    delete zf;
  }
  
  return 0;
}
```

### Read a specific entry from an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

int main(int argc, char** argv) {
  ZipArchive zf("archive.zip");
  zf.open(ZipArchive::ReadOnly);

  //raw access
  char* data = (char*)zf.readEntry("myFile.txt", true);
  ZipEntry entry1 = zf.getEntry("myFile.txt");
  string str1(data, entry1.getSize());

  //text access
  ZipEntry entry2 = zf.getEntry("myFile.txt");
  string str2 = entry2.readAsText();

  zf.close();
  
  return 0;
}
```

### Read a large entry from an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

int main(int argc, char** argv) {
  ZipArchive zf("archive.zip");
  zf.open(ZipArchive::ReadOnly);

  ZipEntry largeEntry = z1.getEntry("largeentry");
  std::ofstream ofUnzippedFile("largeFileContent.data");
  largeEntry.readContent(ofUnzippedFile);
  ofUnzippedFile.close();

  zf.close();

  return 0;
}
```

### Add data to an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

int main(int argc, char** argv) {
  ZipArchive zf("archive.zip");
  zf.open(ZipArchive::Write);
  zf.addEntry("folder/subdir/");

  const char* textData = "Hello,World!";
  zf.addData("helloworld.txt", textData, 12);

  zf.close();

  return 0;
}
```

### Remove data from an archive:

```C++
#include "libzippp.h"
using namespace libzippp;

int main(int argc, char** argv) {
  ZipArchive zf("archive.zip");
  zf.open(ZipArchive::Write);
  zf.deleteEntry("myFile.txt");
  zf.deleteEntry("myDir/subDir/");
  zf.close();
  
  return 0;
}
```

### Progression of committed changes

```C++
#include "libzippp.h"
using namespace libzippp;

class SimpleProgressListener : public ZipProgressListener {
public:
    SimpleProgressListener(void) {}
    virtual ~SimpleProgressListener(void) {}

    void progression(double p) {
        cout << "-- Progression: " << p << endl;
    }
};

int main(int argc, char** argv) {
  ZipArchive zf("archive.zip");
  /* add/modify/delete entries in the archive */

  //register the listener
  SimpleProgressListener spl;
  zf.addProgressListener(&spl);

  //adjust how often the listener will be invoked
  zf.setProgressPrecision(0.1);

  //listener will be invoked
  zf.close();

  return 0;
}
```

### In-memory archives

```C++
#include "libzippp.h"
using namespace libzippp;

int main(int argc, char** argv) {
  char* buffer = new char[4096];

  ZipArchive* z1 = ZipArchive::fromBuffer(buffer, 4096, ZipArchive::New);
  /* add content to the archive */
  
  //will update the content of the buffer
  z1->close();

  //length of the buffer content
  int bufferContentLength = z1->getBufferLength();

  //read again from the archive:
  ZipArchive* z2 = ZipArchive::fromBuffer(buffer, bufferContentLength);
  /* read/modify the archive */

  delete z1;
  delete z2;
  delete buffer;

  return 0;
}
```

## Known issues

### LINUX

You might already have libzip compiled elsewhere on your system. Hence, you
don't need to run 'make libzip'. Instead, just put the libzip location when
you compile libzippp:

```shell
make LIBZIP=path/to/libzip
```

Under Debian, you'll have to install the package `zlib1g-dev` in order to compile
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

### Static linkage

Extra explanations can be found [here](http://hostagebrain.blogspot.com/search/label/zlib).

