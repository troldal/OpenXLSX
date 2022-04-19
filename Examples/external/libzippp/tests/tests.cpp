
/*
  libzippp.h -- exported declarations.
  Copyright (C) 2013 Cédric Tabin

  This file is part of libzippp, a library that wraps libzip for manipulating easily
  ZIP files in C++.
  The author can be contacted on http://www.astorm.ch/blog/index.php?contact

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions
  are met:
  1. Redistributions of source code must retain the above copyright
     notice, this list of conditions and the following disclaimer.
  2. Redistributions in binary form must reproduce the above copyright
     notice, this list of conditions and the following disclaimer in
     the documentation and/or other materials provided with the
     distribution.
  3. The names of the authors may not be used to endorse or promote
     products derived from this software without specific prior
     written permission.
 
  THIS SOFTWARE IS PROVIDED BY THE AUTHORS ``AS IS'' AND ANY EXPRESS
  OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
  ARE DISCLAIMED.  IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
  DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE
  GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER
  IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
  OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
  IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

#include <assert.h>
#include <string.h>
#include <iostream>
#include <iterator>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <string>

#include "libzippp.h"

using namespace std;
using namespace libzippp;

class SimpleProgressListener : public ZipProgressListener {
public:
    SimpleProgressListener(void) : firstValue(-1), lastValue(-1) {}
    virtual ~SimpleProgressListener(void) {}

    double firstValue;
    double lastValue;

    void progression(double p) {
        cout << "-- Progression: " << p << endl;
        if(firstValue<0) { firstValue = p; }
        lastValue = p;
    }
};

void test1() {
    cout << "Running test 1...";
    
    ZipArchive z1("test.zip");
    assert(!z1.isOpen());
    assert(!z1.isMutable());
    z1.open(ZipArchive::Write);
    assert(z1.isOpen());
    assert(z1.isMutable());
    bool result = z1.addEntry("folder/subfolder/finalfolder/");
    assert(result);
    assert(z1.close() == LIBZIPPP_OK);
    assert(!z1.isOpen());
    assert(!z1.isMutable());
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    assert(z2.isOpen());
    assert(!z2.isMutable());
    int nbEntries = z2.getNbEntries();
    assert(nbEntries==3);
    assert(z2.hasEntry("folder/"));
    assert(z2.hasEntry("folder/subfolder/"));
    assert(z2.hasEntry("folder/subfolder/finalfolder/"));
    assert(z2.close() == LIBZIPPP_OK);
    assert(z2.unlink());
    
    cout << " done." << endl;
}

void test2() {
    cout << "Running test 2...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata", txtFile, len);
    
    ZipEntry z1e1 = z1.getEntry("somedata");
    z1e1.setCompressionEnabled(true);
    
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    assert(z2.getNbEntries()==1);
    assert(z2.hasEntry("somedata"));
    
    ZipEntry entry = z2.getEntry("somedata");
    assert(!entry.isNull());
    assert(entry.isCompressionEnabled());
    
    string data = entry.readAsText();
    int clen = data.size();
    assert(clen==len);
    assert(strncmp(txtFile, data.c_str(), len)==0);
    
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test3() {
    cout << "Running test 3...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolder/data.txt", txtFile, len);
    assert(z1.addEntry("somedata/"));
    assert(z1.addEntry("in/"));
    assert(z1.addEntry("in/subfolder/"));
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    assert(z2.getNbEntries()==6);
    assert(z2.hasEntry("somedata/in/subfolder/data.txt"));
    
    ZipEntry entry = z2.getEntry("somedata/in/subfolder/data.txt");
    assert(!entry.isNull());
    assert(!entry.isCompressionEnabled());
    
    string data = entry.readAsText();
    int clen = data.size();
    assert(clen==len);
    assert(strncmp(txtFile, data.c_str(), len)==0);
    
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test4() {
    cout << "Running test 4...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/test.txt", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    assert(z2.getNbEntries()==2);
    
    ZipEntry d = z2.getEntry("somedata/test.txt");
    assert(!d.isNull() && d.isFile());
    assert(z2.deleteEntry(d)==1);
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test5() {
    cout << "Running test 5...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolders/test.txt", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    assert(z2.getNbEntries()==4);
    
    ZipEntry d = z2.getEntry("somedata/in/");
    assert(!d.isNull() && d.isDirectory());
    assert(z2.deleteEntry(d)==3);
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test6() {
    cout << "Running test 6...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolders/test.txt", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    assert(z2.getNbEntries()==4);
    
    ZipEntry d = z2.getEntry("somedata/in/");
    assert(!d.isNull() && d.isDirectory());
    assert(z2.renameEntry(d, "somedata/out/")==3);
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test7() {
    cout << "Running test 7...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolders/test.txt", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    assert(z2.getNbEntries()==4);
    
    ZipEntry d = z2.getEntry("somedata/in/");
    assert(!d.isNull() && d.isDirectory());
    assert(z2.renameEntry(d, "somedata/in/subfolder/")==3);
    z2.close();
    
    ZipArchive z3("test.zip");
    z3.open(ZipArchive::ReadOnly);
    assert(z3.getNbEntries()==5);
    z3.close();
    z3.unlink();
    
    cout << " done." << endl;
}

void test8() {
    cout << "Running test 8...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolders/test.txt", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    assert(z2.getNbEntries()==4);
    
    ZipEntry d = z2.getEntry("somedata/in/");
    assert(!d.isNull() && d.isDirectory());
    assert(z2.renameEntry(d, "newdata/out/subfolders/")==3);
    z2.close();
    
    ZipArchive z3("test.zip");
    z3.open(ZipArchive::ReadOnly);
    assert(z3.getNbEntries()==6);
    z3.close();
    z3.unlink();
    
    cout << " done." << endl;
}

void test9() {
    cout << "Running test 9...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolders/test.txt", txtFile, len);
    z1.addData("somedata/out/subfolders/other.txt", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    assert(z2.getNbEntries()==7);
    
    ZipEntry d = z2.getEntry("somedata/in/");
    assert(!d.isNull() && d.isDirectory());
    assert(z2.renameEntry(d, "root/")==3);
    z2.close();
    
    ZipArchive z3("test.zip");
    z3.open(ZipArchive::ReadOnly);
    assert(z3.getNbEntries()==7);
    z3.close();
    z3.unlink();
    
    cout << " done." << endl;
}

void test10() {
    cout << "Running test 10...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolders/test.txt", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    assert(z2.getNbEntries()==4);
    
    ZipEntry d = z2.getEntry("somedata/in/");
    assert(!d.isNull() && d.isDirectory());
    assert(z2.renameEntry(d, "newdata/out/subfolders/")==3);
    z2.discard();
    
    ZipArchive z3("test.zip");
    z3.open(ZipArchive::ReadOnly);
    assert(z3.getNbEntries()==4);
    z3.close();
    z3.unlink();
    
    cout << " done." << endl;
}

void test11() {
    cout << "Running test 11...";
    
    string c = "a basic comment";
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addEntry("test/");
    z1.setComment(c);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    string str = z2.getComment();
    assert(str == c);
    z2.close();
    
    ZipArchive z3("test.zip");
    z3.open(ZipArchive::Write);
    z3.removeComment();
    z3.close();
    
    ZipArchive z4("test.zip");
    z4.open(ZipArchive::ReadOnly);
    string str2 = z4.getComment();
    assert(str2.empty());
    z4.close();
    z4.unlink();
    
    cout << " done." << endl;
}

void test12() {
    cout << "Running test 12...";
    
    string c = "a basic comment";
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addEntry("test/");
    z1.addData("file/data.txt", c.c_str(), c.length());
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    z2.addEntry("content/new/");
    z2.addData("newfile.txt", c.c_str(), c.length());
    assert(z2.getNbEntries(ZipArchive::Current)==6);
    assert(z2.getNbEntries(ZipArchive::Original)==3);
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test13() {
    cout << "Running test 13...";
    
    string c = "some example of text";
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addEntry("test/");
    z1.addData("file/data.txt", c.c_str(), c.length());
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    z2.renameEntry(z2.getEntry("file/data.txt"), "file/subdir/data.txt");
    assert(z2.getNbEntries(ZipArchive::Current)==4);
    assert(z2.getNbEntries(ZipArchive::Original)==3);
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test14() {
    cout << "Running test 14...";
    
    string c = "some example of text";
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addEntry("test/");
    z1.addData("file/data.txt", c.c_str(), c.length());
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::Write);
    z2.renameEntry(z2.getEntry("file/data.txt"), "content/data/file.txt");
    assert(z2.getNbEntries(ZipArchive::Current)==5);
    assert(z2.getNbEntries(ZipArchive::Original)==3);
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test15() {
    cout << "Running test 15...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata/in/subfolder/data.txt", txtFile, len);
    assert(z1.addEntry("somedata/"));
    assert(z1.addEntry("in/"));
    assert(z1.addEntry("in/subfolder/"));
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    assert(z2.getNbEntries()==6);
    assert(z2.hasEntry("somedata/in/subfolder/data.txt"));
    
    char* data = (char*)z2.readEntry("somedata/in/subfolder/data.txt", true);
    int clen = strlen(data);
    assert(clen==len);
    assert(strncmp(txtFile, data, len)==0);
    
    delete[] data;
    
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test16() {
    cout << "Running test 16...";
    
    string c = "some example of text";
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addEntry("dir/");
    z1.addData("file.txt", c.c_str(), c.length());
    
    ZipEntry e1 = z1.getEntry("dir/");
    ZipEntry e2 = z1.getEntry("file.txt");
    assert(e1.setComment("commentDir"));
    assert(e2.setComment("commentFile"));
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    ZipEntry e12 = z2.getEntry("dir/");
    ZipEntry e22 = z2.getEntry("file.txt");
    assert(e12.getComment() == "commentDir");
    assert(e22.getComment() == "commentFile");
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test17() {
    cout << "Running test 17...";
    
    ZipArchive z1("test.zip");
    assert(z1.open(ZipArchive::ReadOnly) == false);
    void* nil1 = z1.readEntry("an/absent/file.txt", true);
    void* nil2 = z1.readEntry("an/absent/file.txt", true);
    assert(nil1 == NULL);
    assert(nil2 == NULL);
    z1.close();
    z1.unlink();
    
    cout << " done." << endl;
}

void test18() {
    cout << "Running test 18...";
    
    string entry1("a/földér/with/sûbâènts/");
    string entry2("fïle/iñ/sûbfôlder/dàtÄ.txt");
    string text("File wiôth sömè tîxt");
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    assert(z1.addEntry(entry1));
    assert(z1.addData(entry2, text.c_str(), text.length()));
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    ZipEntry e1 = z2.getEntry(entry1);
    assert(!e1.isNull());
    assert(e1.getName() == entry1);
    
    ZipEntry e2 = z2.getEntry(entry2);
    assert(!e2.isNull());
    assert(e2.getName() == entry2);
    assert(e2.readAsText() == text);
    
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test19() {
    cout << "Running test 19...";
    
    const char* txtFile = "this is some data";
    int len = strlen(txtFile);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata", txtFile, len);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    assert(z2.getNbEntries()==1);
    assert(z2.hasEntry("somedata"));
    
    ZipEntry entry = z2.getEntry("somedata");
    assert(!entry.isNull());
    
    string data = entry.readAsText(ZipArchive::Current, 4);
    int clen = data.size();
    assert(clen==4);
    assert(strncmp("this", data.c_str(), 4)==0);
    
    string data2 = entry.readAsText(ZipArchive::Current, 1);
    int clen2 = data2.size();
    assert(clen2==1);
    assert(strncmp("t", data2.c_str(), 1)==0);
    
    string data3 = entry.readAsText(ZipArchive::Current, 999);
    int clen3 = data3.size();
    assert(clen3==len);
    assert(strncmp("this is some data", data3.c_str(), len)==0);
    
    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test20() {
    cout << "Running test 20...";

    const char* txtFile = "this is some data";   // 17 Bytes
    const char* txtFile2 = "this is some data!"; // 18 Bytes
    int len = strlen(txtFile);
    int len2 = strlen(txtFile2);
    
    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata", txtFile, len);
    z1.addData("somedata2", txtFile2, len2);
    z1.close();
    
    ZipArchive z2("test.zip");
    z2.open(ZipArchive::ReadOnly);
    assert(z2.getNbEntries()==2);
    assert(z2.hasEntry("somedata"));
    assert(z2.hasEntry("somedata2"));
    
    ZipEntry entry = z2.getEntry("somedata");
    ZipEntry entry2 = z2.getEntry("somedata2");
    assert(!entry.isNull());
    assert(!entry2.isNull());
    
    {
		// Extract somedata with chunk of 2 bytes, which is not divisible by the file size (17 Bytes)
		std::ofstream ofUnzippedFile("somedata.txt");
		assert(static_cast<bool>(ofUnzippedFile));
		assert(entry.readContent(ofUnzippedFile, ZipArchive::Current, 2) == 0);
		ofUnzippedFile.close();

		std::ifstream ifUnzippedFile("somedata.txt");
		assert(static_cast<bool>(ifUnzippedFile));
		std::string strSomedataText((std::istreambuf_iterator<char>(ifUnzippedFile)), std::istreambuf_iterator<char>());
		assert(strSomedataText.compare(txtFile) == 0);
		ifUnzippedFile.close();
		assert(remove("somedata.txt") == 0);
    }

    {
		// Extract somedata with chunk of 0 bytes (will be defaulted to 512KB).
		std::ofstream ofUnzippedFile("somedata.txt");
		assert(static_cast<bool>(ofUnzippedFile));
		assert(entry.readContent(ofUnzippedFile, ZipArchive::Current, 0) == 0);
		ofUnzippedFile.close();

		std::ifstream ifUnzippedFile("somedata.txt");
		assert(static_cast<bool>(ifUnzippedFile));
		std::string strSomedataText((std::istreambuf_iterator<char>(ifUnzippedFile)), std::istreambuf_iterator<char>());
		assert(strSomedataText.compare(txtFile) == 0);
		ifUnzippedFile.close();
		assert(remove("somedata.txt") == 0);
    }

    {
		// Extract somedata2 with a chunk which is divisible by the size of the file (18 Bytes) to check that the modulo branch
		// is not accessed !
		std::ofstream ofUnzippedFile("somedata2.txt");
		assert(static_cast<bool>(ofUnzippedFile));
		assert(entry2.readContent(ofUnzippedFile, ZipArchive::Current, 2) == 0);
		ofUnzippedFile.close();

		std::ifstream ifUnzippedFile("somedata2.txt");
		assert(static_cast<bool>(ifUnzippedFile));
		std::string strSomedataText((std::istreambuf_iterator<char>(ifUnzippedFile)), std::istreambuf_iterator<char>());
		assert(strSomedataText.compare(txtFile2) == 0);
		ifUnzippedFile.close();
		assert(remove("somedata2.txt") == 0);
    }

    z2.close();
    z2.unlink();
    
    cout << " done." << endl;
}

void test21() {
    cout << "Running test 21...";
    const char* txtFile = "this is some data";   // 17 Bytes
    const char* txtFile2 = "this is some data!"; // 18 Bytes
    int len = strlen(txtFile);
    int len2 = strlen(txtFile2);

    ZipArchive z1("test.zip");
    z1.open(ZipArchive::Write);
    z1.addData("somedata", txtFile, len);
    z1.addData("somedata2", txtFile2, len2);
    z1.close();

    std::ifstream ifs("test.zip", std::ios::binary);
    ifs.seekg(0, std::ifstream::end);
    libzippp_uint32 bufferSize = (libzippp_uint32)ifs.tellg();
    char* buffer = (char*)malloc(bufferSize);
    ifs.seekg(std::ifstream::beg);
    ifs.read(buffer, bufferSize);
    ifs.close();
    
    z1.unlink();

    ZipArchive* z2 = ZipArchive::fromBuffer(buffer, bufferSize);
    assert(z2->getNbEntries() == 2);
    assert(z2->hasEntry("somedata"));
    assert(z2->hasEntry("somedata2"));

    ZipEntry entry = z2->getEntry("somedata");
    ZipEntry entry2 = z2->getEntry("somedata2");
    assert(!entry.isNull());
    assert(!entry2.isNull());

    z2->close();
    delete z2;

    free(buffer);

    cout << " done." << endl;
}

void test22() {
    cout << "Running test 22..." << endl;
    const char* content1 = "This is some text that is a little bit longer, so I can test how the progression callback is invoked.";
    const char* content2 = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus sed metus mollis, facilisis orci vitae, eleifend velit. Quisque et dolor vel nisl gravida vulputate. In iaculis viverra vehicula. Donec viverra euismod odio, sit amet tincidunt nisl aliquam sed. Integer lacinia augue vitae odio varius, at convallis diam molestie. Sed ac nisl at arcu convallis ultricies. Etiam eu metus interdum libero semper vulputate. Proin gravida malesuada justo vel ultrices. Etiam tellus ligula, maximus sed efficitur vitae, iaculis at turpis."
                           "Aliquam eu finibus orci. Quisque maximus enim quis imperdiet vulputate. In vitae velit vel diam scelerisque sollicitudin ac et libero. Donec elit nisi, feugiat ut augue semper, cursus egestas libero. Nullam sed euismod ante. Integer gravida risus nulla, quis vestibulum lacus elementum a. Duis quis vulputate est. Ut elit ipsum, aliquet sit amet gravida et, porttitor quis metus. Vivamus vulputate sed ex ac vulputate. Donec venenatis auctor nulla, quis tempus lorem elementum vel. Suspendisse potenti. In sodales arcu enim, vitae imperdiet quam condimentum sagittis."
                           "Fusce rutrum enim massa, eget ultricies nisi iaculis eu. Quisque erat metus, tempus at volutpat nec, interdum in ligula. Curabitur ullamcorper risus non lobortis vehicula. Proin leo sapien, congue vel metus quis, consectetur lacinia mi. Nunc suscipit erat ipsum, varius commodo risus finibus non. Nam viverra vulputate massa vel pulvinar. Quisque nec quam at lectus sagittis eleifend. Fusce vehicula lectus orci, eu rutrum risus finibus ac. Aenean sit amet mi in velit aliquet faucibus. Donec sit amet diam eget nisl rhoncus posuere eu sodales metus. Integer condimentum placerat neque vitae feugiat. Suspendisse metus velit, faucibus nec hendrerit cursus, pellentesque ut nibh. Duis accumsan mollis elit eget molestie. Donec sit amet congue nibh, quis iaculis lectus. Curabitur placerat sem ex, quis elementum leo sollicitudin sed. Etiam pulvinar turpis vitae ante consequat, vitae eleifend risus dapibus."
                           "Fusce sollicitudin lorem consequat viverra blandit. Praesent feugiat eleifend nibh at eleifend. Etiam quis augue id tortor volutpat placerat. Curabitur id dolor aliquet, consequat eros nec, aliquet velit. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Nulla ut nunc ex. Quisque finibus tincidunt sem, sit amet cursus enim hendrerit ut. Donec non interdum augue. Nunc vehicula viverra sem, at scelerisque sapien venenatis vel. Aliquam pretium, enim vel sodales condimentum, quam erat vehicula arcu, et tincidunt tortor justo aliquet lacus. Nulla dignissim pharetra nibh, at varius arcu tincidunt et. Vestibulum viverra velit tristique risus elementum, eu facilisis sapien sollicitudin. Nullam posuere imperdiet nibh sit amet malesuada. Cras ut dolor blandit, interdum lorem auctor, rhoncus dui. Sed mollis, mauris consequat malesuada pulvinar, dui elit pretium est, id euismod est quam in ex."
                           "Ut euismod nisi in leo tempor fringilla. Praesent vel elit et dui facilisis sollicitudin sed non lacus. Etiam tempor ante a tortor pharetra sollicitudin. Nulla facilisi. Sed tincidunt justo urna, sed porttitor dolor aliquam ac. Duis id enim congue, consequat mauris eget, placerat elit. Maecenas eu leo quis tellus eleifend mattis. Suspendisse porttitor suscipit nunc non facilisis. Pellentesque mollis sapien nec purus consectetur interdum. Nunc efficitur neque rhoncus gravida vestibulum. Nullam at aliquet lorem.";
    int len1 = strlen(content1);
    int len2 = strlen(content2);

    ZipArchive z1("test.zip");
    assert(z1.getProgressPrecision() == LIBZIPPP_DEFAULT_PROGRESSION_PRECISION);
    
    SimpleProgressListener spl;
    z1.setProgressPrecision(0);
    z1.addProgressListener(&spl);
    assert(z1.getProgressListeners().size() == 1);
    
    z1.open(ZipArchive::Write);
    z1.addData("somedata", content1, len1);
    z1.addData("somedata2", content2, len2);
    z1.addData("somedata3", content2, len2);
    z1.addData("somedata4", content1, len1);
    z1.addData("somedata5", content1, len1);
    z1.addData("somedata6", content2, len2);
    z1.close();

    assert(z1.getProgressListeners().size() == 1);
    z1.removeProgressListener(&spl);
    assert(z1.getProgressListeners().size() == 0);

    assert(spl.firstValue==0);
    assert(spl.lastValue=1);

    z1.unlink();

    cout << "complete." << endl;
}

void test23() {
    cout << "Running test 23...";
    const char* txtFile = "this is some data";   // 17 Bytes
    const char* txtFile2 = "this is some data!"; // 18 Bytes
    int len = strlen(txtFile);
    int len2 = strlen(txtFile2);

    char* buffer = new char[4096];
    for(int i=0 ; i<4096; ++i) {
        buffer[i] = '\0';
    }

    ZipArchive* z1 = ZipArchive::fromBuffer(buffer, 4096, ZipArchive::New);
    assert(z1!=nullptr);
    z1->addData("somedata", txtFile, len);
    z1->addData("somedata2", txtFile2, len2);
    int rst = z1->close();
    assert(rst==LIBZIPPP_OK);

    /*cout << endl;
    cout << "Buffer data: " << endl;
    for(int i=0 ; i<4096; ++i) {
        char c = buffer[i];
        if(c == '\0') { c = '0'; }
        
        cout << c;
        if((i+1)%8==0) cout << " ";
        if((i+1)%64==0) cout << endl;
    }*/

    int newLength = z1->getBufferLength();
    delete z1;
    
    ZipArchive* z2 = ZipArchive::fromBuffer(buffer, newLength, ZipArchive::ReadOnly, true);
    assert(z2!=nullptr);
    assert(z2->getNbEntries() == 2);
    assert(z2->hasEntry("somedata"));
    assert(z2->hasEntry("somedata2"));
    delete z2;
    
    ZipArchive* z3 = new ZipArchive("within.zip");
    z3->open(ZipArchive::New);
    z3->addData("inside.zip", buffer, newLength);
    z3->close();
    z3->unlink();
    
    delete z3;
    
    delete[] buffer;

    cout << " done." << endl;
    
    /*char buffer[4096] = {};
    for(int i=0 ; i<4096; ++i) {
        buffer[i] = '\0';
    }
    
    zip_error_t error;
    zip_source_t *zs = zip_source_buffer_create(buffer, sizeof(buffer), 0, &error);
    assert(zs!=nullptr);
    
    zip_source_keep(zs);
    
    zip_t * zip = zip_open_from_source(zs, ZIP_TRUNCATE, &error);
    assert(zip!=nullptr);
    
    zip_add_dir(zip, "mydir");
    
    zip_close(zip);
    
    zip_source_open(zs);
    zip_source_read(zs, buffer, sizeof(buffer));
    zip_source_close(zs);
    
    cout << "Buffer data: " << endl;
    for(int i=0 ; i<4096; ++i) {
        char c = buffer[i];
        cout << (int)c;
        if((i+1)%8==0) cout << " ";
        if((i+1)%64==0) cout << endl;
    }*/
}

int main(int argc, char** argv) {
    test1();  test2();  test3();  test4();  test5();
    test6();  test7();  test8();  test9();  test10();
    test11(); test12(); test13(); test14(); test15();
    test16(); test17(); test18(); test19(); test20();
    test21(); test22(); test23();
    return 0;
}


