//
//  Copyright (c) 2015 Artyom Beilis (Tonkikh)
//  Copyright (c) 2019 Alexander Grund
//
//  Distributed under the Boost Software License, Version 1.0. (See
//  accompanying file LICENSE or copy at
//  http://www.boost.org/LICENSE_1_0.txt)
//

#include <nowide/fstream.hpp>

#include <nowide/convert.hpp>
#include <nowide/cstdio.hpp>
#include <fstream>
#include <iostream>
#include <string>

#include "test.hpp"

namespace nw = nowide;

struct remove_file_at_exit
{
    const char* filename;
    remove_file_at_exit(const char* filename) : filename(filename)
    {}
    ~remove_file_at_exit() noexcept(false)
    {
        TEST(nw::remove(filename) == 0);
    }
};

void make_empty_file(const char* filepath)
{
    nw::ofstream f(filepath, std::ios_base::out | std::ios::trunc);
    TEST(f);
}

bool file_exists(const char* filepath)
{
    FILE* f = nw::fopen(filepath, "r");
    if(f)
    {
        std::fclose(f);
        return true;
    } else
        return false;
}

std::string read_file(const char* filepath, bool binary_mode = false)
{
    FILE* f = nw::fopen(filepath, binary_mode ? "rb" : "r");
    TEST(f);
    std::string content;
    int c;
    while((c = std::fgetc(f)) != EOF)
        content.push_back(static_cast<char>(c));
    std::fclose(f);
    return content;
}

void test_with_different_buffer_sizes(const char* filepath)
{
    /* Important part of the standard for mixing input with output:
       However, output shall not be directly followed by input without an intervening call to the fflush function
       or to a file positioning function (fseek, fsetpos, or rewind),
       and input shall not be directly followed by output without an intervening call to a file positioning function,
       unless the input operation encounters end-of-file.
    */
    for(int i = -1; i < 16; i++)
    {
        std::cout << "Buffer size = " << i << std::endl;
        char buf[16];
        nw::fstream f;
        // Different conditions when setbuf might be called: Usually before opening a file is OK
        if(i >= 0)
            f.rdbuf()->pubsetbuf((i == 0) ? NULL : buf, i);
        f.open(filepath, std::ios::in | std::ios::out | std::ios::trunc | std::ios::binary);
        TEST(f);
        remove_file_at_exit _(filepath);

        // Add 'abcdefg'
        TEST(f.put('a'));
        TEST(f.put('b'));
        TEST(f.put('c'));
        TEST(f.put('d'));
        TEST(f.put('e'));
        TEST(f.put('f'));
        TEST(f.put('g'));
        // Read first char
        TEST(f.seekg(0));
        TEST(f.get() == 'a');
        TEST(f.gcount() == 1u);
        // Skip next char
        TEST(f.seekg(1, std::ios::cur));
        TEST(f.get() == 'c');
        TEST(f.gcount() == 1u);
        // Go back 1 char
        TEST(f.seekg(-1, std::ios::cur));
        TEST(f.get() == 'c');
        TEST(f.gcount() == 1u);

        // Test switching between read->write->read
        // case 1) overwrite, flush, read
        TEST(f.seekg(1));
        TEST(f.put('B'));
        TEST(f.flush()); // Flush when changing out->in
        TEST(f.get() == 'c');
        TEST(f.gcount() == 1u);
        TEST(f.seekg(1));
        TEST(f.get() == 'B');
        TEST(f.gcount() == 1u);
        // case 2) overwrite, seek, read
        TEST(f.seekg(2));
        TEST(f.put('C'));
        TEST(f.seekg(3)); // Seek when changing out->in
        TEST(f.get() == 'd');
        TEST(f.gcount() == 1u);

        // Check that sequence from start equals expected
        TEST(f.seekg(0));
        TEST(f.get() == 'a');
        TEST(f.get() == 'B');
        TEST(f.get() == 'C');
        TEST(f.get() == 'd');
        TEST(f.get() == 'e');

        // Putback after flush is implementation defined
        // Boost.Nowide: Works
#if NOWIDE_USE_FILEBUF_REPLACEMENT
        TEST(f << std::flush);
        TEST(f.putback('e'));
        TEST(f.putback('d'));
        TEST(f.get() == 'd');
        TEST(f.get() == 'e');
#endif
        // Rest of sequence
        TEST(f.get() == 'f');
        TEST(f.get() == 'g');
        TEST(f.get() == EOF);

        // Put back until front of file is reached
        f.clear();
        TEST(f.seekg(1));
        TEST(f.get() == 'B');
        TEST(f.putback('B'));
        // Putting back multiple chars is not possible on all implementations after a seek/flush
#if NOWIDE_USE_FILEBUF_REPLACEMENT
        TEST(f.putback('a'));
        TEST(!f.putback('x')); // At beginning of file -> No putback possible
        // Get characters that were putback to avoid MSVC bug https://github.com/microsoft/STL/issues/342
        f.clear();
        TEST(f.get() == 'a');
#endif
        TEST(f.get() == 'B');
        f.close();
    }
}

void test_close(const char* filepath)
{
    const std::string filepath2 = std::string(filepath) + ".2";
    // Make sure file does not exist yet
    TEST(!file_exists(filepath2.c_str()) || nw::remove(filepath2.c_str()) == 0);
    TEST(!file_exists(filepath2.c_str()));
    nw::filebuf buf;
    TEST(buf.open(filepath, std::ios_base::out) == &buf);
    TEST(buf.is_open());
    // Opening when already open fails
    TEST(buf.open(filepath2.c_str(), std::ios_base::out) == NULL);
    // Still open
    TEST(buf.is_open());
    TEST(buf.close() == &buf);
    // Failed opening did not create file
    TEST(!file_exists(filepath2.c_str()));
    // But it should work now:
    TEST(buf.open(filepath2.c_str(), std::ios_base::out) == &buf);
    TEST(buf.close() == &buf);
    TEST(file_exists(filepath2.c_str()));
    TEST(nw::remove(filepath) == 0);
    TEST(nw::remove(filepath2.c_str()) == 0);
}

template<typename IFStream, typename OFStream>
void test_flush(const char* filepath)
{
    OFStream fo(filepath, std::ios_base::out | std::ios::trunc);
    TEST(fo);
    std::string curValue;
    for(int repeat = 0; repeat < 2; repeat++)
    {
        for(size_t len = 1; len <= 1024; len *= 2)
        {
            char c = static_cast<char>(len % 13 + repeat + 'a'); // semi-random char
            std::string input(len, c);
            fo << input;
            curValue += input;
            TEST(fo.flush());
            std::string s;
            // Note: Flush on read area is implementation defined, so check whole file instead
            IFStream fi(filepath);
            TEST(fi >> s);
            // coverity[tainted_data]
            TEST(s == curValue);
        }
    }
}

void test_ofstream_creates_file(const char* filename)
{
    TEST(!file_exists(filename) || nw::remove(filename) == 0);
    TEST(!file_exists(filename));
    // Ctor
    {
        nw::ofstream fo(filename);
        TEST(fo);
    }
    {
        TEST(file_exists(filename));
        remove_file_at_exit _(filename);
        TEST(read_file(filename).empty());
    }
    // Open
    {
        nw::ofstream fo;
        fo.open(filename);
        TEST(fo);
    }
    {
        TEST(file_exists(filename));
        remove_file_at_exit _(filename);
        TEST(read_file(filename).empty());
    }
}

// Create filename file with content "test\n"
void test_ofstream_write(const char* filename)
{
    // char* ctor
    {
        nw::ofstream fo(filename);
        TEST(fo << "test" << 2 << std::endl);
    }
    // char* open
    TEST(read_file(filename) == "test2\n");
    TEST(nw::remove(filename) == 0);
    {
        nw::ofstream fo;
        fo.open(filename);
        TEST(fo << "test" << 2 << std::endl);
    }
    TEST(read_file(filename) == "test2\n");
    TEST(nw::remove(filename) == 0);
    // string ctor
    {
        std::string name = filename;
        nw::ofstream fo(name);
        TEST(fo << "test" << 2 << std::endl);
    }
    TEST(read_file(filename) == "test2\n");
    TEST(nw::remove(filename) == 0);
    // string open
    {
        nw::ofstream fo;
        fo.open(std::string(filename));
        TEST(fo << "test" << 2 << std::endl);
    }
    TEST(read_file(filename) == "test2\n");
    TEST(nw::remove(filename) == 0);
    // Binary mode
    {
        nw::ofstream fo(filename, std::ios::binary);
        TEST(fo << "test" << 2 << std::endl);
    }
    TEST(read_file(filename, true) == "test2\n");
    TEST(nw::remove(filename) == 0);
    // At end
    {
        {
            nw::ofstream fo(filename);
            TEST(fo << "test" << 2 << std::endl);
        }
        nw::ofstream fo(filename, std::ios::ate | std::ios::in);
        fo << "second" << 2 << std::endl;
    }
    TEST(read_file(filename) == "test2\nsecond2\n");
    TEST(nw::remove(filename) == 0);
}

void test_ifstream_open_read(const char* filename)
{
    // Create test file
    {
        nw::ofstream fo(filename);
        TEST(fo << "test" << std::endl);
    }

    // char* Ctor
    {
        nw::ifstream fi(filename);
        TEST(fi);
        std::string tmp;
        TEST(fi >> tmp);
        TEST(tmp == "test");
    }
    // char* open
    {
        nw::ifstream fi;
        fi.open(filename);
        TEST(fi);
        std::string tmp;
        TEST(fi >> tmp);
        TEST(tmp == "test");
    }
    // string ctor
    {
        std::string name = filename;
        nw::ifstream fi(name);
        TEST(fi);
        std::string tmp;
        TEST(fi >> tmp);
        TEST(tmp == "test");
    }
    // string open
    {
        nw::ifstream fi;
        fi.open(std::string(filename));
        TEST(fi);
        std::string tmp;
        TEST(fi >> tmp);
        TEST(tmp == "test");
    }
    // Binary mode
    {
        nw::ifstream fi(filename, std::ios::binary);
        TEST(fi);
        std::string tmp;
        TEST(fi >> tmp);
        TEST(tmp == "test");
    }
    // At end
    {
        // Need binary file or position check might be throw off by newline conversion
        {
            nw::ofstream fo(filename, nw::fstream::binary);
            TEST(fo << "test");
        }
        nw::ifstream fi(filename, nw::fstream::ate | nw::fstream::binary);
        TEST(fi);
        TEST(fi.tellg() == std::streampos(4));
        fi.seekg(-2, std::ios_base::cur);
        std::string tmp;
        TEST(fi >> tmp);
        TEST(tmp == "st");
    }
    // Fail on non-existing file
    TEST(nw::remove(filename) == 0);
    {
        nw::ifstream fi(filename);
        TEST(!fi);
    }
}

void test_fstream(const char* filename)
{
    const std::string sFilename = filename;
    TEST(!file_exists(filename) || nw::remove(filename) == 0);
    TEST(!file_exists(filename));
    // Fail on non-existing file
    {
        nw::fstream f(filename);
        TEST(!f);
        nw::fstream f2(sFilename);
        TEST(!f2);
    }
    {
        nw::fstream f;
        f.open(filename);
        TEST(!f);
        f.open(sFilename);
        TEST(!f);
    }
    TEST(!file_exists(filename));
    // Create empty file (Ctor)
    {
        nw::fstream f(filename, std::ios::out);
        TEST(f);
    }
    TEST(read_file(filename).empty());
    // Char* ctor
    {
        nw::fstream f(filename);
        TEST(f);
        TEST(f << "test");
        std::string tmp;
        TEST(f.seekg(0));
        TEST(f >> tmp);
        TEST(tmp == "test");
    }
    TEST(read_file(filename) == "test");
    // String ctor
    {
        nw::fstream f(sFilename);
        TEST(f);
        TEST(f << "string_ctor");
        std::string tmp;
        TEST(f.seekg(0));
        TEST(f >> tmp);
        TEST(tmp == "string_ctor");
    }
    TEST(read_file(filename) == "string_ctor");
    TEST(nw::remove(filename) == 0);
    // Create empty file (open)
    {
        nw::fstream f;
        f.open(filename, std::ios::out);
        TEST(f);
    }
    remove_file_at_exit _(filename);

    TEST(read_file(filename).empty());
    // Open
    {
        nw::fstream f;
        f.open(filename);
        TEST(f);
        TEST(f << "test");
        std::string tmp;
        TEST(f.seekg(0));
        TEST(f >> tmp);
        TEST(tmp == "test");
    }
    TEST(read_file(filename) == "test");
    // Ctor existing file
    {
        nw::fstream f(filename);
        TEST(f);
        std::string tmp;
        TEST(f >> tmp);
        TEST(tmp == "test");
        TEST(f.eof());
        f.clear();
        TEST(f << "second");
    }
    TEST(read_file(filename) == "testsecond");
    // Trunc & binary
    {
        nw::fstream f(filename, std::ios::in | std::ios::out | std::ios::trunc | std::ios::binary);
        TEST(f);
        TEST(f << "test2");
        std::string tmp;
        TEST(f.seekg(0));
        TEST(f >> tmp);
        TEST(tmp == "test2");
    }
    TEST(read_file(filename) == "test2");
    // Reading in write mode fails (existing file!)
    {
        nw::fstream f(filename, std::ios::out);
        std::string tmp;
        TEST(!(f >> tmp));
        f.clear();
        TEST(f << "foo");
        TEST(f.seekg(0));
        TEST(!(f >> tmp));
    }
    TEST(read_file(filename) == "foo");
    // Writing in read mode fails (existing file!)
    {
        nw::fstream f(filename, std::ios::in);
        TEST(!(f << "bar"));
        f.clear();
        std::string tmp;
        TEST(f >> tmp);
        TEST(tmp == "foo");
    }
    TEST(read_file(filename) == "foo");
}

template<typename T>
bool is_open(T& stream)
{
    // There are const and non const versions of is_open, so test both
    TEST(stream.is_open() == const_cast<const T&>(stream).is_open());
    return stream.is_open();
}

template<typename T>
void do_test_is_open(const char* filename)
{
    T f;
    TEST(!is_open(f));
    f.open(filename);
    TEST(f);
    TEST(is_open(f));
    f.close();
    TEST(f);
    TEST(!is_open(f));
}

void test_is_open(const char* filename)
{
    // Note the order: Output before input so file exists
    do_test_is_open<nw::ofstream>(filename);
    remove_file_at_exit _(filename);
    do_test_is_open<nw::ifstream>(filename);
    do_test_is_open<nw::fstream>(filename);
}

// Reproducer for https://github.com/boostorg/nowide/issues/126
void test_getline_and_tellg(const char* filename)
{
    {
        nw::ofstream f(filename);
        f << "Line 1" << std::endl;
        f << "Line 2" << std::endl;
        f << "Line 3" << std::endl;
    }
    remove_file_at_exit _(filename);
    nw::fstream f;
    // Open file in text mode, to read
    f.open(filename, std::ios_base::in);
    TEST(f);
    std::string line1, line2, line3;
    TEST(getline(f, line1));
    TEST(line1 == "Line 1");
    const auto tg = f.tellg(); // This may cause issues
    TEST(tg > 0u);
    TEST(getline(f, line2));
    TEST(line2 == "Line 2");
    TEST(getline(f, line3));
    TEST(line3 == "Line 3");
}

// Test that a sync after a peek does not swallow newlines
// This can happen because peek reads a char which needs to be "unread" on sync which may loose a converted newline
void test_peek_sync_get(const char* filename)
{
    {
        nw::ofstream f(filename);
        f << "Line 1" << std::endl;
        f << "Line 2" << std::endl;
    }
    remove_file_at_exit _(filename);
    nw::ifstream f(filename);
    TEST(f);
    while(f)
    {
        const int curChar = f.peek();
        if(curChar == std::char_traits<char>::eof())
            break;
        f.sync();
        TEST(f.get() == char(curChar));
    }
}

void test_swap(const char* filename, const char* filename2)
{
    {
        nw::ofstream f(filename);
        f << "02468" << std::endl;
        f.close();
        f.open(filename2);
        f << "13579" << std::endl;
    }
    remove_file_at_exit _(filename);
    remove_file_at_exit _2(filename2);

    nw::ifstream f1(filename);
    nw::ifstream f2(filename2);
    TEST(f1);
    TEST(f2);
    while(f1 && f2)
    {
        const int curChar1 = f1.peek();
        const int curChar2 = f2.peek();
        f1.swap(f2);
        TEST(f1.peek() == curChar2);
        TEST(f2.peek() == curChar1);
        if(curChar1 == std::char_traits<char>::eof() || curChar2 == std::char_traits<char>::eof())
            break;
        TEST(f1.get() == char(curChar2));
        f1.swap(f2);
        TEST(f1.get() == char(curChar1));
    }
}

void test_main(int, char** argv, char**)
{
    const std::string exampleFilename = std::string(argv[0]) + "-\xd7\xa9-\xd0\xbc-\xce\xbd.txt";
    const std::string exampleFilename2 = std::string(argv[0]) + "-\xd7\xa9-\xd0\xbc-\xce\xbd 2.txt";

    std::cout << "Testing fstream" << std::endl;
    test_ofstream_creates_file(exampleFilename.c_str());
    test_ofstream_write(exampleFilename.c_str());
    test_ifstream_open_read(exampleFilename.c_str());
    test_fstream(exampleFilename.c_str());
    test_is_open(exampleFilename.c_str());

    std::cout << "Complex IO" << std::endl;
    test_with_different_buffer_sizes(exampleFilename.c_str());

    std::cout << "filebuf::close" << std::endl;
    test_close(exampleFilename.c_str());

    std::cout << "Flush - Sanity Check" << std::endl;
    test_flush<std::ifstream, std::ofstream>(exampleFilename.c_str());
    std::cout << "Flush - Test" << std::endl;
    test_flush<nw::ifstream, nw::ofstream>(exampleFilename.c_str());

    std::cout << "Regression tests" << std::endl;
    test_getline_and_tellg(exampleFilename.c_str());
    test_peek_sync_get(exampleFilename.c_str());
    test_swap(exampleFilename.c_str(), exampleFilename2.c_str());
}
