# Zippy
Zippy is a simple C++ wrapper around the 'miniz' zip library. It allows for creating, reading and modifying zip archives.

The underlying 'miniz' library is included and Zippy is therefore completely self-contained; it does not rely on zlib, minizip, libzip or any other external libraries.

## Usage
Here is a short example, showing how to use Zippy:

```cpp
#include <iostream>
#include <Zippy/Zippy.h>

using namespace std;

static const string str = "MISSION CONTROL I wouldn't worry too much about the computer. First of all, there is still a chance that he is right, despite your tests, and if it should happen again, we suggest eliminating this possibility by allowing the unit to remain in place and seeing whether or not it actually fails. If the computer should turn out to be wrong, the situation is still not alarming. The type of obsessional error he may be guilty of is not unknown among the latest generation of HAL 9000 computers. It has almost always revolved around a single detail, such as the one you have described, and it has never interfered with the integrity or reliability of the computer's performance in other areas. No one is certain of the cause of this kind of malfunctioning. It may be over-programming, but it could also be any number of reasons. In any event, it is somewhat analogous to human neurotic behavior. Does this answer your query?  Zero-five-three-Zero, MC, transmission concluded.";

int main() {

    // ===== Creating empty archive
    Zippy::ZipArchive arch;
    arch.Create("./TestArchive.zip");

    // ===== Adding 10 entries to the archive
    for (int i = 0; i <= 9; ++i)
        arch.AddEntry(to_string(i) + ".txt", "this is test " + to_string(i) + ": " + str);

    // ===== Delete the first entry
    arch.DeleteEntry("0.txt");

    // ===== Save and close the archive
    arch.Save();
    arch.Close();

    // ===== Reopen and check contents
    arch.Open("./TestArchive.zip");
    cout << "Number of entries in archive: " << arch.GetNumEntries() << endl;
    cout << "Content of \"9.txt\": " << endl << arch.GetEntry("9.txt").GetDataAsString();

    return 0;
}
```

## Background
For my OpenXLSX library, I have been looking around for a simple C++ library for handling .zip files. Unfortunately, that proved very difficult to find, so I ended up relying on a combination of zlib libzip and libzippp (https://github.com/ctabin/libzippp). However, handling several external dependencies turned out to be very cumbersome, so I began looking for alternatives.

The 'miniz' library (https://github.com/richgel999/miniz) has a proven track record as a .zip library contained in a single source file. Unfortunately, it's written in C and is not very well documented. A project called 'miniz-cpp' (https://github.com/tfussell/miniz-cpp) wraps miniz in C++ classes in order to make it more accessible. It was developed as part of the XLNT library (https://github.com/tfussell/xlnt). Nevertheless, it did not fulfill the requirements I had, so I began developing Zippy instead.

## Features
The ambition for Zippy is to enable reading, writing and modification of existing zip archives, as well as creation of new archives, by means of an easy-to-use C++ interface.

In addition, some features are not directly available in 'miniz' (e.g. deletion of archive entries), so instead workarounds for these have been implemented directly in Zippy.

## Current Status
The library already works well, but additional work is required to ensure robustness. 

It is my hope that, in the future, this can become a header-only library.
 
## Compatibility
The library has been tested on:
- MacOS (GCC and Clang)

In the future, it will be tested on Linux and Windows as well.
