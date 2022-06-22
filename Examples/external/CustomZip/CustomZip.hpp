//
// Created by Kenneth Balslev on 18/04/2022.
//

#ifndef OPENXLSX_CUSTOMZIP_HPP
#define OPENXLSX_CUSTOMZIP_HPP

#include <cstring>
#include <fstream>
#include <list>
// #include <libzippp.h>
#include <zip.h>
#include "nowide/cstdio.hpp"
#include <memory>
#include <random>
#include <string>

class CustomZip
{
public:
    /**
     * @brief
     */
    CustomZip() : m_zipHandle(nullptr) {};

    /**
     * @brief
     * @param other
     */
    CustomZip(const CustomZip& other) = default;

    /**
     * @brief
     * @param other
     */
    CustomZip(CustomZip&& other) noexcept = default;

    /**
     * @brief
     */
    ~CustomZip() = default;

    /**
     * @brief
     * @param other
     * @return
     */
    CustomZip& operator=(const CustomZip& other) = default;

    /**
     * @brief
     * @param other
     * @return
     */
    CustomZip& operator=(CustomZip&& other) noexcept = default;

    /**
     * @brief
     * @return
     */
    explicit operator bool() const { return isValid(); };

    /**
     *
     * @return
     */
    bool isValid() const { return isOpen(); }

    /**
     * @brief
     * @return
     */
    inline bool isOpen() const { return m_zipHandle != nullptr; }

    /**
     * @brief
     * @param fileName
     */
    void open(const std::string& fileName)
    {
        // ===== If a file is already open, close (discard) it.
        if (isOpen()) close();

        // ===== Generate name for the temporary file.
        const int NAME_LENGTH = 20;
        m_fileName = fileName;
        m_tempName = m_fileName.substr(0, m_fileName.rfind('/') + 1) + GenerateRandomName(NAME_LENGTH);

        // ===== Copy the contents of the file to open, to the temporary file.
        std::ifstream src(m_fileName, std::ios::binary);
        std::ofstream dst(m_tempName, std::ios::binary);
        dst << src.rdbuf();
        src.close();
        dst.close();

        // ===== Open the temporary zip file
        int errorFlag = 0;
        m_zipHandle = zip_open(m_tempName.c_str(), ZIP_CREATE, &errorFlag);
    }

    /**
     * @brief
     */
    void close()
    {
        // ===== If the file is open, discard changes and set the handle to nullptr.
        if (isOpen()) {
            zip_discard(m_zipHandle);
            m_zipHandle = nullptr;
        }

        // ===== Delete the temporary file.
        nowide::remove(m_tempName.c_str());    // NOLINT
    }

    /**
     * @brief
     * @param path
     */
    void save(const std::string& path = "")
    {
        // ===== Discard any changes made since last commit;
        if (isOpen()) {
            zip_discard(m_zipHandle);
            m_zipHandle = nullptr;
        }

        // ===== Delete the original file, and rename the temporary file to that of the original.
        if (!path.empty()) m_fileName = path;
        nowide::remove(m_fileName.c_str());                        // NOLINT
        nowide::rename(m_tempName.c_str(), m_fileName.c_str());    // NOLINT

        // ===== Reopen the file.
        open(m_fileName);
    }

    /**
     * @brief
     * @param name
     * @param data
     */
    void addEntry(const std::string& name, const std::string& data)
    {
        // ===== Create a buffer with the data and immediately commit it to the zip file.
        zip_source_t* source = zip_source_buffer(m_zipHandle, data.data(), data.size(), false);
        zip_file_add(m_zipHandle, name.c_str(), source, ZIP_FL_OVERWRITE | ZIP_FL_ENC_GUESS);

        // ===== Commit the change to the temporary archive (this is safer and - surprisingly - faster than committing everything at once)
        zip_close(m_zipHandle);
        int errorFlag = 0;
        m_zipHandle   = zip_open(m_tempName.c_str(), ZIP_CREATE, &errorFlag);
    }

    /**
     * @brief
     * @param entryName
     */
    void deleteEntry(const std::string& entryName) {
        if (hasEntry(entryName)) {
            zip_stat_t stat;
            zip_stat(m_zipHandle, entryName.c_str(), ZIP_FL_ENC_GUESS, &stat);
            zip_delete(m_zipHandle, stat.index);
        }
    }

    /**
     * @brief
     * @param name
     * @return
     */
    std::string getEntry(const std::string& name) {

        // ===== Create the output string and open the file in the zip archive.
        std::string result;
        std::unique_ptr<zip_file_t, decltype(&zip_fclose)> zipFile(zip_fopen(m_zipHandle, name.c_str(), ZIP_FL_ENC_GUESS), &zip_fclose);

        // ===== Access the file data and copy the data to the output string.
        if (zipFile) {
            zip_stat_t stat;
            zip_stat(m_zipHandle, name.c_str(), ZIP_FL_ENC_GUESS, &stat);

            auto data = std::unique_ptr<char[]>(new char[stat.size + 1]);
            zip_fread(zipFile.get(), data.get(), stat.size);
            data[stat.size] = '\0';
            result = data.get();
        }

        return result;
    }

    /**
     * @brief
     * @param entryName
     * @return
     */
    bool hasEntry(const std::string& entryName) {

        // ===== If the index of the entryName is higher than zero, it means that an entry with that name exists in the archive.
        auto index = zip_name_locate(m_zipHandle, entryName.c_str(), ZIP_FL_ENC_GUESS);
        return index >= 0;
    }

private:
    inline std::string GenerateRandomName(int length)
    {
        std::string letters = "abcdefghijklmnopqrstuvwxyz0123456789"; // NOLINT

        std::random_device                 rand_dev;
        std::mt19937                       generator(rand_dev());
        std::uniform_int_distribution<int> distr(0, letters.size() - 1);

        std::string result;
        for (int i = 0; i < length; ++i) {
            result += letters[distr(generator)];
        }

        return result + ".tmp";
    }

    zip* m_zipHandle = nullptr;
    std::string          m_fileName;
    std::string          m_tempName;
};

#endif    // OPENXLSX_CUSTOMZIP_HPP
