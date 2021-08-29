//
//  Copyright (c) 2015 Artyom Beilis (Tonkikh)
//
//  Distributed under the Boost Software License, Version 1.0. (See
//  accompanying file LICENSE or copy at
//  http://www.boost.org/LICENSE_1_0.txt)
//

#include <nowide/iostream.hpp>

#include <nowide/utf/utf.hpp>
#include <limits>
#include <string>

#include "test.hpp"

bool isValidUTF8(const std::string& s)
{
    using namespace nowide::utf;
    for(std::string::const_iterator it = s.begin(); it != s.end();)
    {
        code_point c = utf_traits<char>::decode(it, s.end());
        if(!is_valid_codepoint(c))
            return false;
    }
    return true;
}

void test_main(int argc, char** argv, char**)
{
    const char* example = "Basic letters: \xd7\xa9-\xd0\xbc-\xce\xbd\n"
                          "East Asian Letters: \xe5\x92\x8c\xe5\xb9\xb3\n"
                          "Non-BMP letters: \xf0\x9d\x84\x9e\n"
                          "Invalid UTF-8: `\xFF' `\xd7\xFF' `\xe5\xFF\x8c' `\xf0\x9d\x84\xFF' \n"
                          "\n";

    // If we are using the standard rdbuf we can only put back 1 char
    if(nowide::cin.rdbuf() == std::cin.rdbuf())
    {
        std::cout << "Using std::cin" << std::endl;
        int maxval = 15000;
        for(int i = 0; i < maxval; i++)
        {
            char c = i % 96 + ' ';
            TEST(nowide::cin.putback(c));
            int ci = i % 96 + ' ';
            TEST(nowide::cin.get() == ci);
        }
    } else
    {
        int maxval = 15000;
        for(int i = 0; i < maxval; i++)
        {
            char c = i % 96 + ' ';
            TEST(nowide::cin.putback(c));
        }
        for(int i = maxval - 1; i >= 0; i--)
        {
            int c = i % 96 + ' ';
            TEST(nowide::cin.get() == c);
        }
    }
    nowide::cout << "Normal I/O:" << std::endl;
    nowide::cout << example << std::endl;
    nowide::cerr << example << std::endl;

    nowide::cout << "Flushing each character:" << std::endl;

    for(const char* s = example; *s; s++)
    {
        nowide::cout << *s << std::flush;
        TEST(nowide::cout);
    }

    TEST(nowide::cout);
    TEST(nowide::cerr);
    if(argc == 2 && argv[1] == std::string("-i"))
    {
        nowide::cout << "Input 2 strings" << std::endl;
        std::string v1, v2;
        nowide::cin >> v1 >> v2;
        TEST(nowide::cin);
        TEST(isValidUTF8(v1));
        TEST(isValidUTF8(v2));
        nowide::cout << "First:  " << v1 << std::endl;
        nowide::cout << "Second: " << v2 << std::endl;
        TEST(nowide::cout);

        // Check sync
        nowide::cout << "Input 2 strings\n";
        nowide::cout.flush();
        TEST(nowide::cin >> v1);
        nowide::cin.sync();
        nowide::cout << "First:  " << v1 << std::endl;
        nowide::cout << "2nd string should be ignored. Input 1 more + [ENTER]" << std::endl;
        // And check getline not getting the CR
        TEST(std::getline(nowide::cin, v1));
        TEST(!v1.empty() && v1[v1.size() - 1u] != '\r');
        nowide::cout << "Value:  " << v1 << std::endl;

        nowide::cout << "Press ENTER to exit";
        nowide::cin.clear();
        nowide::cin.ignore(std::numeric_limits<std::streamsize>::max(), '\n');
        nowide::cin.get();
    }
}
