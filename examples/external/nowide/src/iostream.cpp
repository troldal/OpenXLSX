//
//  Copyright (c) 2012 Artyom Beilis (Tonkikh)
//
//  Distributed under the Boost Software License, Version 1.0. (See
//  accompanying file LICENSE or copy at
//  http://www.boost.org/LICENSE_1_0.txt)
//
#define NOWIDE_SOURCE
#include <nowide/iostream.hpp>

#ifndef NOWIDE_WINDOWS

namespace nowide {
    /// Avoid empty compilation unit warnings
    /// \internal
    NOWIDE_DECL void dummy_exported_function()
    {}
} // namespace nowide

#else

#include <nowide/convert.hpp>
#include <cassert>
#include <cstring>
#include <iostream>
#include <vector>

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>

namespace nowide {
    namespace detail {

        namespace {
            bool is_atty_handle(HANDLE h)
            {
                if(h)
                {
                    DWORD dummy;
                    return GetConsoleMode(h, &dummy) != FALSE;
                }
                return false;
            }
        } // namespace

        class console_output_buffer : public std::streambuf
        {
        public:
            console_output_buffer(HANDLE h) : handle_(h)
            {}

        protected:
            int sync()
            {
                return overflow(traits_type::eof());
            }
            int overflow(int c)
            {
                if(!handle_)
                    return -1;
                int n = static_cast<int>(pptr() - pbase());
                int r = 0;

                if(n > 0 && (r = write(pbase(), n)) < 0)
                    return -1;
                if(r < n)
                {
                    std::memmove(pbase(), pbase() + r, n - r);
                }
                setp(buffer_, buffer_ + buffer_size);
                pbump(n - r);
                if(c != traits_type::eof())
                    sputc(traits_type::to_char_type(c));
                return 0;
            }

        private:
            using decoder = utf::utf_traits<char>;
            using encoder = utf::utf_traits<wchar_t>;

            int write(const char* p, int n)
            {
                namespace uf = utf;
                const char* b = p;
                const char* e = p + n;
                DWORD size = 0;
                if(n > buffer_size)
                    return -1;
                wchar_t* out = wbuffer_;
                uf::code_point c;
                size_t decoded = 0;
                while((c = decoder::decode(p, e)) != uf::incomplete)
                {
                    if(c == uf::illegal)
                        c = NOWIDE_REPLACEMENT_CHARACTER;
                    assert(out - wbuffer_ + encoder::width(c) <= static_cast<int>(wbuffer_size));
                    out = encoder::encode(c, out);
                    decoded = p - b;
                }
                if(!WriteConsoleW(handle_, wbuffer_, static_cast<DWORD>(out - wbuffer_), &size, 0))
                    return -1;
                return static_cast<int>(decoded);
            }

            static const int buffer_size = 1024;
            static const int wbuffer_size = buffer_size * encoder::max_width;
            char buffer_[buffer_size];
            wchar_t wbuffer_[wbuffer_size];
            HANDLE handle_;
        };

        class console_input_buffer : public std::streambuf
        {
        public:
            console_input_buffer(HANDLE h) : handle_(h), wsize_(0), was_newline_(true)
            {}

        protected:
            int sync()
            {
                if(FlushConsoleInputBuffer(handle_) == 0)
                    return -1;
                wsize_ = 0;
                was_newline_ = true;
                pback_buffer_.clear();
                setg(0, 0, 0);
                return 0;
            }
            int pbackfail(int c)
            {
                if(c == traits_type::eof())
                    return traits_type::eof();

                if(gptr() != eback())
                {
                    gbump(-1);
                    *gptr() = traits_type::to_char_type(c);
                    return 0;
                }

                char* pnext;
                if(pback_buffer_.empty())
                {
                    pback_buffer_.resize(4);
                    pnext = &pback_buffer_[0] + pback_buffer_.size() - 1u;
                } else
                {
                    size_t n = pback_buffer_.size();
                    pback_buffer_.resize(n * 2);
                    std::memcpy(&pback_buffer_[n], &pback_buffer_[0], n);
                    pnext = &pback_buffer_[0] + n - 1;
                }

                char* pFirst = &pback_buffer_[0];
                char* pLast = pFirst + pback_buffer_.size();
                setg(pFirst, pnext, pLast);
                *gptr() = traits_type::to_char_type(c);

                return 0;
            }

            int underflow()
            {
                if(!handle_)
                    return -1;
                if(!pback_buffer_.empty())
                    pback_buffer_.clear();

                size_t n = read();
                setg(buffer_, buffer_, buffer_ + n);
                if(n == 0)
                    return traits_type::eof();
                return std::char_traits<char>::to_int_type(*gptr());
            }

        private:
            using decoder = utf::utf_traits<wchar_t>;
            using encoder = utf::utf_traits<char>;

            size_t read()
            {
                DWORD read_wchars = 0;
                const size_t n = wbuffer_size - wsize_;
                if(!ReadConsoleW(handle_, wbuffer_ + wsize_, static_cast<DWORD>(n), &read_wchars, 0))
                    return 0;
                wsize_ += read_wchars;
                char* out = buffer_;
                const wchar_t* cur_input_ptr = wbuffer_;
                const wchar_t* const end_input_ptr = wbuffer_ + wsize_;
                while(cur_input_ptr != end_input_ptr)
                {
                    const wchar_t* const prev_input_ptr = cur_input_ptr;
                    utf::code_point c = decoder::decode(cur_input_ptr, end_input_ptr);
                    // If incomplete restore to beginning of incomplete char to use on next buffer
                    if(c == utf::incomplete)
                    {
                        cur_input_ptr = prev_input_ptr;
                        break;
                    }
                    if(c == utf::illegal)
                        c = NOWIDE_REPLACEMENT_CHARACTER;
                    assert(out - buffer_ + encoder::width(c) <= static_cast<int>(buffer_size));
                    // Skip \r chars as std::cin does
                    if(c != '\r')
                        out = encoder::encode(c, out);
                }

                wsize_ = end_input_ptr - cur_input_ptr;
                if(wsize_ > 0)
                    std::memmove(wbuffer_, end_input_ptr - wsize_, sizeof(wchar_t) * wsize_);

                // A CTRL+Z at the start of the line should be treated as EOF
                if(was_newline_ && out > buffer_ && buffer_[0] == '\x1a')
                {
                    sync();
                    return 0;
                }
                was_newline_ = out == buffer_ || out[-1] == '\n';

                return out - buffer_;
            }

            static const size_t wbuffer_size = 1024;
            static const size_t buffer_size = wbuffer_size * encoder::max_width;
            char buffer_[buffer_size];
            wchar_t wbuffer_[wbuffer_size];
            HANDLE handle_;
            size_t wsize_;
            std::vector<char> pback_buffer_;
            bool was_newline_;
        };

        winconsole_ostream::winconsole_ostream(int fd, winconsole_ostream* tieStream) : std::ostream(0)
        {
            HANDLE h = 0;
            switch(fd)
            {
            case 1: h = GetStdHandle(STD_OUTPUT_HANDLE); break;
            case 2: h = GetStdHandle(STD_ERROR_HANDLE); break;
            }
            if(is_atty_handle(h))
            {
                d.reset(new console_output_buffer(h));
                std::ostream::rdbuf(d.get());
            } else
            {
                std::ostream::rdbuf(fd == 1 ? std::cout.rdbuf() : std::cerr.rdbuf());
            }
            if(tieStream)
                tie(tieStream);
        }
        winconsole_ostream::~winconsole_ostream()
        {
            try
            {
                flush();
            } catch(...)
            {}
        }

        winconsole_istream::winconsole_istream(winconsole_ostream* tieStream) : std::istream(0)
        {
            HANDLE h = GetStdHandle(STD_INPUT_HANDLE);
            if(is_atty_handle(h))
            {
                d.reset(new console_input_buffer(h));
                std::istream::rdbuf(d.get());
            } else
            {
                std::istream::rdbuf(std::cin.rdbuf());
            }
            if(tieStream)
                tie(tieStream);
        }

        winconsole_istream::~winconsole_istream()
        {}

    } // namespace detail

    detail::winconsole_ostream cout(1, NULL);
    detail::winconsole_istream cin(&cout);
    detail::winconsole_ostream cerr(2, &cout);
    detail::winconsole_ostream clog(2, &cout);
} // namespace nowide

#endif
