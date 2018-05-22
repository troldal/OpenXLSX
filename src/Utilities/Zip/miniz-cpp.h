/*
  This is free and unencumbered software released into the public domain.

  Anyone is free to copy, modify, publish, use, compile, sell, or
  distribute this software, either in source code form or as a compiled
  binary, for any purpose, commercial or non-commercial, and by any
  means.

  In jurisdictions that recognize copyright laws, the author or authors
  of this software dedicate any and all copyright interest in the
  software to the public domain. We make this dedication for the benefit
  of the public at large and to the detriment of our heirs and
  successors. We intend this dedication to be an overt act of
  relinquishment in perpetuity of all present and future rights to this
  software under copyright law.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
  IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
  OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
  ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
  OTHER DEALINGS IN THE SOFTWARE.

  For more information, please refer to <http://unlicense.org/>
*/

#ifndef OPENXLSX_ALL_MINIZ_CPP_H
#define OPENXLSX_ALL_MINIZ_CPP_H

#include <string>
#include <vector>
#include <iostream>
#include <ostream>
#include <fstream>
#include <sstream>
#include "miniz.h"

namespace miniz_cpp {
    namespace detail {

#ifdef _WIN32
        inline char directory_separator = '\\';
        inline char alt_directory_separator = '/';
#else
        inline char directory_separator = '/';
        inline char alt_directory_separator = '\\';
#endif

        inline std::string join_path(const std::vector<std::string> &parts)
        {
            std::string joined;
            std::size_t i = 0;
            for(auto part : parts)
            {
                joined.append(part);

                if(i++ != parts.size() - 1)
                {
                    joined.append(1, '/');
                }
            }
            return joined;
        }

        inline std::vector<std::string> split_path(const std::string &path, char delim = directory_separator)
        {
            std::vector<std::string> split;
            std::string::size_type previous_index = 0;
            auto separator_index = path.find(delim);

            while(separator_index != std::string::npos)
            {
                auto part = path.substr(previous_index, separator_index - previous_index);
                if(part != "..")
                {
                    split.push_back(part);
                }
                else
                {
                    split.pop_back();
                }
                previous_index = separator_index + 1;
                separator_index = path.find(delim, previous_index);
            }

            split.push_back(path.substr(previous_index));

            if(split.size() == 1 && delim == directory_separator)
            {
                auto alternative = split_path(path, alt_directory_separator);
                if(alternative.size() > 1)
                {
                    return alternative;
                }
            }

            return split;
        }

        inline uint32_t crc32buf(const char *buf, std::size_t len)
        {
            uint32_t oldcrc32 = 0xFFFFFFFF;

            uint32_t crc_32_tab[] = { /* CRC polynomial 0xedb88320 */
                0x00000000, 0x77073096, 0xee0e612c, 0x990951ba, 0x076dc419, 0x706af48f,
                0xe963a535, 0x9e6495a3, 0x0edb8832, 0x79dcb8a4, 0xe0d5e91e, 0x97d2d988,
                0x09b64c2b, 0x7eb17cbd, 0xe7b82d07, 0x90bf1d91, 0x1db71064, 0x6ab020f2,
                0xf3b97148, 0x84be41de, 0x1adad47d, 0x6ddde4eb, 0xf4d4b551, 0x83d385c7,
                0x136c9856, 0x646ba8c0, 0xfd62f97a, 0x8a65c9ec, 0x14015c4f, 0x63066cd9,
                0xfa0f3d63, 0x8d080df5, 0x3b6e20c8, 0x4c69105e, 0xd56041e4, 0xa2677172,
                0x3c03e4d1, 0x4b04d447, 0xd20d85fd, 0xa50ab56b, 0x35b5a8fa, 0x42b2986c,
                0xdbbbc9d6, 0xacbcf940, 0x32d86ce3, 0x45df5c75, 0xdcd60dcf, 0xabd13d59,
                0x26d930ac, 0x51de003a, 0xc8d75180, 0xbfd06116, 0x21b4f4b5, 0x56b3c423,
                0xcfba9599, 0xb8bda50f, 0x2802b89e, 0x5f058808, 0xc60cd9b2, 0xb10be924,
                0x2f6f7c87, 0x58684c11, 0xc1611dab, 0xb6662d3d, 0x76dc4190, 0x01db7106,
                0x98d220bc, 0xefd5102a, 0x71b18589, 0x06b6b51f, 0x9fbfe4a5, 0xe8b8d433,
                0x7807c9a2, 0x0f00f934, 0x9609a88e, 0xe10e9818, 0x7f6a0dbb, 0x086d3d2d,
                0x91646c97, 0xe6635c01, 0x6b6b51f4, 0x1c6c6162, 0x856530d8, 0xf262004e,
                0x6c0695ed, 0x1b01a57b, 0x8208f4c1, 0xf50fc457, 0x65b0d9c6, 0x12b7e950,
                0x8bbeb8ea, 0xfcb9887c, 0x62dd1ddf, 0x15da2d49, 0x8cd37cf3, 0xfbd44c65,
                0x4db26158, 0x3ab551ce, 0xa3bc0074, 0xd4bb30e2, 0x4adfa541, 0x3dd895d7,
                0xa4d1c46d, 0xd3d6f4fb, 0x4369e96a, 0x346ed9fc, 0xad678846, 0xda60b8d0,
                0x44042d73, 0x33031de5, 0xaa0a4c5f, 0xdd0d7cc9, 0x5005713c, 0x270241aa,
                0xbe0b1010, 0xc90c2086, 0x5768b525, 0x206f85b3, 0xb966d409, 0xce61e49f,
                0x5edef90e, 0x29d9c998, 0xb0d09822, 0xc7d7a8b4, 0x59b33d17, 0x2eb40d81,
                0xb7bd5c3b, 0xc0ba6cad, 0xedb88320, 0x9abfb3b6, 0x03b6e20c, 0x74b1d29a,
                0xead54739, 0x9dd277af, 0x04db2615, 0x73dc1683, 0xe3630b12, 0x94643b84,
                0x0d6d6a3e, 0x7a6a5aa8, 0xe40ecf0b, 0x9309ff9d, 0x0a00ae27, 0x7d079eb1,
                0xf00f9344, 0x8708a3d2, 0x1e01f268, 0x6906c2fe, 0xf762575d, 0x806567cb,
                0x196c3671, 0x6e6b06e7, 0xfed41b76, 0x89d32be0, 0x10da7a5a, 0x67dd4acc,
                0xf9b9df6f, 0x8ebeeff9, 0x17b7be43, 0x60b08ed5, 0xd6d6a3e8, 0xa1d1937e,
                0x38d8c2c4, 0x4fdff252, 0xd1bb67f1, 0xa6bc5767, 0x3fb506dd, 0x48b2364b,
                0xd80d2bda, 0xaf0a1b4c, 0x36034af6, 0x41047a60, 0xdf60efc3, 0xa867df55,
                0x316e8eef, 0x4669be79, 0xcb61b38c, 0xbc66831a, 0x256fd2a0, 0x5268e236,
                0xcc0c7795, 0xbb0b4703, 0x220216b9, 0x5505262f, 0xc5ba3bbe, 0xb2bd0b28,
                0x2bb45a92, 0x5cb36a04, 0xc2d7ffa7, 0xb5d0cf31, 0x2cd99e8b, 0x5bdeae1d,
                0x9b64c2b0, 0xec63f226, 0x756aa39c, 0x026d930a, 0x9c0906a9, 0xeb0e363f,
                0x72076785, 0x05005713, 0x95bf4a82, 0xe2b87a14, 0x7bb12bae, 0x0cb61b38,
                0x92d28e9b, 0xe5d5be0d, 0x7cdcefb7, 0x0bdbdf21, 0x86d3d2d4, 0xf1d4e242,
                0x68ddb3f8, 0x1fda836e, 0x81be16cd, 0xf6b9265b, 0x6fb077e1, 0x18b74777,
                0x88085ae6, 0xff0f6a70, 0x66063bca, 0x11010b5c, 0x8f659eff, 0xf862ae69,
                0x616bffd3, 0x166ccf45, 0xa00ae278, 0xd70dd2ee, 0x4e048354, 0x3903b3c2,
                0xa7672661, 0xd06016f7, 0x4969474d, 0x3e6e77db, 0xaed16a4a, 0xd9d65adc,
                0x40df0b66, 0x37d83bf0, 0xa9bcae53, 0xdebb9ec5, 0x47b2cf7f, 0x30b5ffe9,
                0xbdbdf21c, 0xcabac28a, 0x53b39330, 0x24b4a3a6, 0xbad03605, 0xcdd70693,
                0x54de5729, 0x23d967bf, 0xb3667a2e, 0xc4614ab8, 0x5d681b02, 0x2a6f2b94,
                0xb40bbe37, 0xc30c8ea1, 0x5a05df1b, 0x2d02ef8d
            };

#define UPDC32(octet,crc) (crc_32_tab[((crc)\
^ static_cast<uint8_t>(octet)) & 0xff] ^ ((crc) >> 8))

            for ( ; len; --len, ++buf)
            {
                oldcrc32 = UPDC32(*buf, oldcrc32);
            }

            return ~oldcrc32;
        }

        inline tm safe_localtime(const time_t &t)
        {
#ifdef _WIN32
            tm time;
    localtime_s(&time, &t);
    return time;
#else
            tm *time = localtime(&t);
            assert(time != nullptr);
            return *time;
#endif
        }

        inline std::size_t write_callback(void *opaque, std::uint64_t file_ofs, const void *pBuf, std::size_t n)
        {
            auto buffer = static_cast<std::vector<char> *>(opaque);

            if(file_ofs + n > buffer->size())
            {
                auto new_size = static_cast<std::vector<char>::size_type>(file_ofs + n);
                buffer->resize(new_size);
            }

            for(std::size_t i = 0; i < n; i++)
            {
                (*buffer)[static_cast<std::size_t>(file_ofs + i)] = (static_cast<const char *>(pBuf))[i];
            }

            return n;
        }

    } // namespace detail

    struct zip_info
    {
        std::string filename;

        struct
        {
            int year = 1980;
            int month = 0;
            int day = 0;
            int hours = 0;
            int minutes = 0;
            int seconds = 0;
        } date_time;

        std::string comment;
        std::string extra;
        uint16_t create_system = 0;
        uint16_t create_version = 0;
        uint16_t extract_version = 0;
        uint16_t flag_bits = 0;
        std::size_t volume = 0;
        uint32_t internal_attr = 0;
        uint32_t external_attr = 0;
        std::size_t header_offset = 0;
        uint32_t crc = 0;
        std::size_t compress_size = 0;
        std::size_t file_size = 0;
    };

    class zip_file
    {
    public:
        inline zip_file() : archive_(new mz_zip_archive())
        {
            reset();
        }

        inline zip_file(const std::string &filename) : zip_file()
        {
            load(filename);
        }

        inline zip_file(std::istream &stream) : zip_file()
        {
            load(stream);
        }

        inline zip_file(const std::vector<unsigned char> &bytes) : zip_file()
        {
            load(bytes);
        }

        inline ~zip_file()
        {
            reset();
        }

        inline void load(std::istream &stream)
        {
            reset();
            buffer_.assign(std::istreambuf_iterator<char>(stream), std::istreambuf_iterator<char>());
            remove_comment();
            start_read();
        }

        inline void load(const std::string &filename)
        {
            filename_ = filename;
            std::ifstream stream(filename, std::ios::binary);
            load(stream);
        }

        inline void load(const std::vector<unsigned char> &bytes)
        {
            reset();
            buffer_.assign(bytes.begin(), bytes.end());
            remove_comment();
            start_read();
        }

        inline void save(const std::string &filename)
        {
            filename_ = filename;
            std::ofstream stream(filename, std::ios::binary);
            save(stream);
        }

        inline void save(std::ostream &stream)
        {
            if(archive_->m_zip_mode == MZ_ZIP_MODE_WRITING)
            {
                mz_zip_writer_finalize_archive(archive_.get());
            }

            if(archive_->m_zip_mode == MZ_ZIP_MODE_WRITING_HAS_BEEN_FINALIZED)
            {
                mz_zip_writer_end(archive_.get());
            }

            if(archive_->m_zip_mode == MZ_ZIP_MODE_INVALID)
            {
                start_read();
            }

            append_comment();
            stream.write(buffer_.data(), static_cast<long>(buffer_.size()));
        }

        inline void save(std::vector<unsigned char> &bytes)
        {
            if(archive_->m_zip_mode == MZ_ZIP_MODE_WRITING)
            {
                mz_zip_writer_finalize_archive(archive_.get());
            }

            if(archive_->m_zip_mode == MZ_ZIP_MODE_WRITING_HAS_BEEN_FINALIZED)
            {
                mz_zip_writer_end(archive_.get());
            }

            if(archive_->m_zip_mode == MZ_ZIP_MODE_INVALID)
            {
                start_read();
            }

            append_comment();
            bytes.assign(buffer_.begin(), buffer_.end());
        }

        inline void reset()
        {
            switch(archive_->m_zip_mode)
            {
                case MZ_ZIP_MODE_READING:
                    mz_zip_reader_end(archive_.get());
                    break;
                case MZ_ZIP_MODE_WRITING:
                    mz_zip_writer_finalize_archive(archive_.get());
                    mz_zip_writer_end(archive_.get());
                    break;
                case MZ_ZIP_MODE_WRITING_HAS_BEEN_FINALIZED:
                    mz_zip_writer_end(archive_.get());
                    break;
                case MZ_ZIP_MODE_INVALID:
                    break;
            }

            if(archive_->m_zip_mode != MZ_ZIP_MODE_INVALID)
            {
                throw std::runtime_error("");
            }

            buffer_.clear();
            comment.clear();

            start_write();
            mz_zip_writer_finalize_archive(archive_.get());
            mz_zip_writer_end(archive_.get());
        }

        inline bool has_file(const std::string &name)
        {
            if(archive_->m_zip_mode != MZ_ZIP_MODE_READING)
            {
                start_read();
            }

            int index = mz_zip_reader_locate_file(archive_.get(), name.c_str(), nullptr, 0);

            return index != -1;
        }

        inline bool has_file(const zip_info &name)
        {
            return has_file(name.filename);
        }

        inline zip_info getinfo(const std::string &name)
        {
            if(archive_->m_zip_mode != MZ_ZIP_MODE_READING)
            {
                start_read();
            }

            int index = mz_zip_reader_locate_file(archive_.get(), name.c_str(), nullptr, 0);

            if(index == -1)
            {
                throw std::runtime_error("not found");
            }

            return getinfo(index);
        }

        inline std::vector<zip_info> infolist()
        {
            if(archive_->m_zip_mode != MZ_ZIP_MODE_READING)
            {
                start_read();
            }

            std::vector<zip_info> info;

            for(std::size_t i = 0; i < mz_zip_reader_get_num_files(archive_.get()); i++)
            {
                info.push_back(getinfo(static_cast<int>(i)));
            }

            return info;
        }

        inline std::vector<std::string> namelist()
        {
            std::vector<std::string> names;

            for(auto &info : infolist())
            {
                names.push_back(info.filename);
            }

            return names;
        }

        inline std::ostream &open(const std::string &name)
        {
            return open(getinfo(name));
        }

        inline std::ostream &open(const zip_info &name)
        {
            auto data = read(name);
            std::string data_string(data.begin(), data.end());
            open_stream_ << data_string;
            return open_stream_;
        }

        inline void extract(const std::string &member, const std::string &path)
        {
            std::fstream stream(detail::join_path({path, member}), std::ios::binary | std::ios::out);
            stream << open(member).rdbuf();
        }

        inline void extract(const zip_info &member, const std::string &path)
        {
            std::fstream stream(detail::join_path({path, member.filename}), std::ios::binary | std::ios::out);
            stream << open(member).rdbuf();
        }

        inline void extractall(const std::string &path)
        {
            extractall(path, infolist());
        }

        inline void extractall(const std::string &path, const std::vector<std::string> &members)
        {
            for(auto &member : members)
            {
                extract(member, path);
            }
        }

        inline void extractall(const std::string &path, const std::vector<zip_info> &members)
        {
            for(auto &member : members)
            {
                extract(member, path);
            }
        }

        inline void printdir()
        {
            printdir(std::cout);
        }

        inline void printdir(std::ostream &stream)
        {
            stream << "  Length " << "  " << "   " << "Date" << "   " << " " << "Time " << "   " << "Name" << std::endl;
            stream << "---------  ---------- -----   ----" << std::endl;

            std::size_t sum_length = 0;
            std::size_t file_count = 0;

            for(auto &member : infolist())
            {
                sum_length += member.file_size;
                file_count++;

                std::string length_string = std::to_string(member.file_size);
                while(length_string.length() < 9)
                {
                    length_string = " " + length_string;
                }
                stream << length_string;

                stream << "  ";
                stream << (member.date_time.month < 10 ? "0" : "") << member.date_time.month;
                stream << "/";
                stream << (member.date_time.day < 10 ? "0" : "") << member.date_time.day;
                stream << "/";
                stream << member.date_time.year;
                stream << " ";
                stream << (member.date_time.hours < 10 ? "0" : "") << member.date_time.hours;
                stream << ":";
                stream << (member.date_time.minutes < 10 ? "0" : "") << member.date_time.minutes;
                stream << "   ";
                stream << member.filename;
                stream << std::endl;
            }

            stream << "---------                     -------" << std::endl;

            std::string length_string = std::to_string(sum_length);
            while(length_string.length() < 9)
            {
                length_string = " " + length_string;
            }
            stream << length_string << "                     " << file_count << " " << (file_count == 1 ? "file" : "files");
            stream << std::endl;
        }

        inline std::string read(const zip_info &info)
        {
            std::size_t size;
            char *data = static_cast<char *>(mz_zip_reader_extract_file_to_heap(archive_.get(), info.filename.c_str(), &size, 0));
            if(data == nullptr)
            {
                throw std::runtime_error("file couldn't be read");
            }
            std::string extracted(data, data + size);
            mz_free(data);
            return extracted;
        }

        inline std::string read(const std::string &name)
        {
            return read(getinfo(name));
        }

        inline std::pair<bool, std::string> testzip()
        {
            if(archive_->m_zip_mode == MZ_ZIP_MODE_INVALID)
            {
                throw std::runtime_error("not open");
            }

            for(auto &file : infolist())
            {
                auto content = read(file);
                auto crc = detail::crc32buf(content.c_str(), content.size());

                if(crc != file.crc)
                {
                    return {false, file.filename};
                }
            }

            return {true, ""};
        }

        inline void write(const std::string &filename)
        {
            auto split = detail::split_path(filename);
            if(split.size() > 1)
            {
                split.erase(split.begin());
            }
            auto arcname = detail::join_path(split);
            write(filename, arcname);
        }

        inline void write(const std::string &filename, const std::string &arcname)
        {
            std::fstream file(filename, std::ios::binary | std::ios::in);
            std::stringstream ss;
            ss << file.rdbuf();
            std::string bytes = ss.str();

            writestr(arcname, bytes);
        }

        inline void writestr(const std::string &arcname, const std::string &bytes)
        {
            if(archive_->m_zip_mode != MZ_ZIP_MODE_WRITING)
            {
                start_write();
            }

            if(!mz_zip_writer_add_mem(archive_.get(), arcname.c_str(), bytes.data(), bytes.size(), MZ_BEST_COMPRESSION))
            {
                throw std::runtime_error("write error");
            }
        }

        inline void writestr(const zip_info &info, const std::string &bytes)
        {
            if(info.filename.empty() || info.date_time.year < 1980)
            {
                throw std::runtime_error("must specify a filename and valid date (year >= 1980");
            }

            if(archive_->m_zip_mode != MZ_ZIP_MODE_WRITING)
            {
                start_write();
            }

            auto crc = detail::crc32buf(bytes.c_str(), bytes.size());

            if(!mz_zip_writer_add_mem_ex(archive_.get(), info.filename.c_str(), bytes.data(), bytes.size(), info.comment.c_str(), static_cast<mz_uint16>(info.comment.size()), MZ_BEST_COMPRESSION, 0, crc))
            {
                throw std::runtime_error("write error");
            }
        }

        std::string get_filename() const { return filename_; }

        std::string comment;

    private:
        inline void start_read()
        {
            if(archive_->m_zip_mode == MZ_ZIP_MODE_READING) return;

            if(archive_->m_zip_mode == MZ_ZIP_MODE_WRITING)
            {
                mz_zip_writer_finalize_archive(archive_.get());
            }

            if(archive_->m_zip_mode == MZ_ZIP_MODE_WRITING_HAS_BEEN_FINALIZED)
            {
                mz_zip_writer_end(archive_.get());
            }

            if(!mz_zip_reader_init_mem(archive_.get(), buffer_.data(), buffer_.size(), 0))
            {
                throw std::runtime_error("bad zip");
            }
        }

        inline void start_write()
        {
            if(archive_->m_zip_mode == MZ_ZIP_MODE_WRITING) return;

            switch(archive_->m_zip_mode)
            {
                case MZ_ZIP_MODE_READING:
                {
                    mz_zip_archive archive_copy;
                    std::memset(&archive_copy, 0, sizeof(mz_zip_archive));
                    std::vector<char> buffer_copy(buffer_.begin(), buffer_.end());

                    if(!mz_zip_reader_init_mem(&archive_copy, buffer_copy.data(), buffer_copy.size(), 0))
                    {
                        throw std::runtime_error("bad zip");
                    }

                    mz_zip_reader_end(archive_.get());

                    archive_->m_pWrite = &detail::write_callback;
                    archive_->m_pIO_opaque = &buffer_;
                    buffer_ = std::vector<char>();

                    if(!mz_zip_writer_init(archive_.get(), 0))
                    {
                        throw std::runtime_error("bad zip");
                    }

                    for(unsigned int i = 0; i < static_cast<unsigned int>(archive_copy.m_total_files); i++)
                    {
                        if(!mz_zip_writer_add_from_zip_reader(archive_.get(), &archive_copy, i))
                        {
                            throw std::runtime_error("fail");
                        }
                    }

                    mz_zip_reader_end(&archive_copy);
                    return;
                }
                case MZ_ZIP_MODE_WRITING_HAS_BEEN_FINALIZED:
                    mz_zip_writer_end(archive_.get());
                    break;
                case MZ_ZIP_MODE_INVALID:
                case MZ_ZIP_MODE_WRITING:
                    break;
            }

            archive_->m_pWrite = &detail::write_callback;
            archive_->m_pIO_opaque = &buffer_;

            if(!mz_zip_writer_init(archive_.get(), 0))
            {
                throw std::runtime_error("bad zip");
            }
        }

        inline void append_comment()
        {
            if(!comment.empty())
            {
                auto comment_length = std::min(static_cast<uint16_t>(comment.length()), std::numeric_limits<uint16_t>::max());
                buffer_[buffer_.size() - 2] = static_cast<char>(comment_length);
                buffer_[buffer_.size() - 1] = static_cast<char>(comment_length >> 8);
                std::copy(comment.begin(), comment.end(), std::back_inserter(buffer_));
            }
        }

        inline void remove_comment()
        {
            if(buffer_.empty()) return;

            std::size_t position = buffer_.size() - 1;

            for(; position >= 3; position--)
            {
                if(buffer_[position - 3] == 'P'
                    && buffer_[position - 2] == 'K'
                    && buffer_[position - 1] == '\x05'
                    && buffer_[position] == '\x06')
                {
                    position = position + 17;
                    break;
                }
            }

            if(position == 3)
            {
                throw std::runtime_error("didn't find end of central directory signature");
            }

            uint16_t length = static_cast<uint16_t>(buffer_[position + 1]);
            length = static_cast<uint16_t>(length << 8) + static_cast<uint16_t>(buffer_[position]);
            position += 2;

            if(length != 0)
            {
                comment = std::string(buffer_.data() + position, buffer_.data() + position + length);
                buffer_.resize(buffer_.size() - length);
                buffer_[buffer_.size() - 1] = 0;
                buffer_[buffer_.size() - 2] = 0;
            }
        }

        inline zip_info getinfo(int index)
        {
            if(archive_->m_zip_mode != MZ_ZIP_MODE_READING)
            {
                start_read();
            }

            mz_zip_archive_file_stat stat;
            mz_zip_reader_file_stat(archive_.get(), static_cast<mz_uint>(index), &stat);

            zip_info result;

            result.filename = std::string(stat.m_filename, stat.m_filename + std::strlen(stat.m_filename));
            result.comment = std::string(stat.m_comment, stat.m_comment + stat.m_comment_size);
            result.compress_size = static_cast<std::size_t>(stat.m_comp_size);
            result.file_size = static_cast<std::size_t>(stat.m_uncomp_size);
            result.header_offset = static_cast<std::size_t>(stat.m_local_header_ofs);
            result.crc = stat.m_crc32;
            auto time = detail::safe_localtime(stat.m_time);
            result.date_time.year = 1900 + time.tm_year;
            result.date_time.month = 1 + time.tm_mon;
            result.date_time.day = time.tm_mday;
            result.date_time.hours = time.tm_hour;
            result.date_time.minutes = time.tm_min;
            result.date_time.seconds = time.tm_sec;
            result.flag_bits = stat.m_bit_flag;
            result.internal_attr = stat.m_internal_attr;
            result.external_attr = stat.m_external_attr;
            result.extract_version = stat.m_version_needed;
            result.create_version = stat.m_version_made_by;
            result.volume = stat.m_file_index;
            result.create_system = stat.m_method;

            return result;
        }

        std::unique_ptr<mz_zip_archive> archive_;
        std::vector<char> buffer_;
        std::stringstream open_stream_;
        std::string filename_;
    };

} // namespace miniz_cpp

#endif //OPENXLSX_ALL_MINIZ_CPP_H
