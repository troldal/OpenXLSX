//
//  Copyright (c) 2019 Alexander Grund
//
//  Distributed under the Boost Software License, Version 1.0. (See
//  accompanying file LICENSE or copy at
//  http://www.boost.org/LICENSE_1_0.txt)
//

#include <nowide/config.hpp>

#include <nowide/fstream.hpp>

#include <nowide/cstdio.hpp>
#include <iterator>
#include <utility>

#include "test.hpp"

namespace nw = nowide;

void create_file(const std::string& filename, const std::string& contents)
{
    nw::ofstream f(filename, std::ios::trunc);
    TEST(f << contents);
}

template<typename T>
std::string get_file_contents(T& stream)
{
    TEST(stream);
    std::string s((std::istreambuf_iterator<char>(stream)), std::istreambuf_iterator<char>());
    return s;
}

std::string get_file_contents(const std::string& filename)
{
    nw::ifstream f(filename);
    return get_file_contents(f);
}

nw::ifstream make_ifstream(const std::string& filename)
{
    nw::ifstream f(filename);
    TEST(f);
    std::string s;
    TEST(f >> s);
    TEST(s == "Hello");
    return f;
}

void test_ifstream(const std::string& filename)
{
    create_file(filename, "Hello\nWorld");
    // Move construct
    {
        nw::ifstream f = make_ifstream(filename);
        TEST(f);
        std::string s;
        TEST(f >> s);
        TEST(s == "World");
    }
    // Move assign
    {
        nw::ifstream f;
        {
            nw::ifstream f2 = make_ifstream(filename);
            f = std::move(f2);
        }
        TEST(f);
        std::string s;
        TEST(f >> s);
        TEST(s == "World");
    }
    // Swap
    {
        nw::ifstream f;
        {
            nw::ifstream f2 = make_ifstream(filename);
            swap(f, f2);
            TEST(!f2.is_open());
        }
        TEST(f);
        std::string s;
        TEST(f >> s);
        TEST(s == "World");
    }
    TEST(nw::remove(filename.c_str()) == 0);
}

nw::ofstream make_ofstream(const std::string& filename)
{
    nw::ofstream f(filename);
    TEST(f);
    TEST(f << "Hello");
    return f;
}

void test_ofstream(const std::string& filename)
{
    // Move construct
    {
        nw::ofstream f = make_ofstream(filename);
        TEST(f);
        TEST(f << " world");
        f.close();
        TEST(get_file_contents(filename) == "Hello world");
    }
    // Move assign
    {
        nw::ofstream f;
        {
            nw::ofstream f2 = make_ofstream(filename);
            f = std::move(f2);
        }
        TEST(f);
        TEST(f << " world");
        f.close();
        TEST(get_file_contents(filename) == "Hello world");
    }
    // Swap
    {
        nw::ofstream f;
        {
            nw::ofstream f2 = make_ofstream(filename);
            swap(f, f2);
            TEST(!f2.is_open());
        }
        TEST(f);
        TEST(f << " world");
        f.close();
        TEST(get_file_contents(filename) == "Hello world");
    }
    TEST(nw::remove(filename.c_str()) == 0);
}

nw::fstream make_fstream(const std::string& filename)
{
    create_file(filename, "");
    nw::fstream f(filename);
    TEST(f << "Hello");
    return f;
}

void test_fstream(const std::string& filename)
{
    // Move construct
    {
        nw::fstream f = make_fstream(filename);
        TEST(f);
        TEST(f << " world");
        f.seekg(0);
        TEST(get_file_contents(f) == "Hello world");
    }
    // Move assign
    {
        nw::fstream f;
        {
            nw::fstream f2 = make_fstream(filename);
            f = std::move(f2);
        }
        TEST(f);
        TEST(f << " world");
        f.seekg(0);
        TEST(get_file_contents(f) == "Hello world");
    }
    // Swap
    {
        nw::fstream f;
        {
            nw::fstream f2 = make_fstream(filename);
            swap(f, f2);
            TEST(!f2.is_open());
        }
        TEST(f);
        TEST(f << " world");
        f.seekg(0);
        TEST(get_file_contents(f) == "Hello world");
    }
    TEST(nw::remove(filename.c_str()) == 0);
}

void test_main(int, char** argv, char**)
{
    const std::string exampleFilename = std::string(argv[0]) + "-\xd7\xa9-\xd0\xbc-\xce\xbd.txt";
    test_ifstream(exampleFilename);
    test_ofstream(exampleFilename);
    test_fstream(exampleFilename);
}
