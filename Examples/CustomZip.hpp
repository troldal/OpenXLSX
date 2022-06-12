//
// Created by Kenneth Balslev on 18/04/2022.
//

#ifndef OPENXLSX_CUSTOMZIP_HPP
#define OPENXLSX_CUSTOMZIP_HPP

#include <cstring>
#include <fstream>
// #include <libzippp.h>
#include <memory>
#include <nowide/cstdio.hpp>
#include <random>
#include <string>
#include <zip.h>

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
        if (isOpen()) close();

        const int NAME_LENGTH = 20;
        m_fileName = fileName;
        m_tempName = m_fileName.substr(0, m_fileName.rfind('/') + 1) + GenerateRandomName(NAME_LENGTH);

        std::ifstream src(m_fileName, std::ios::binary);
        std::ofstream dst(m_tempName, std::ios::binary);
        dst << src.rdbuf();
        src.close();
        dst.close();

        int errorFlag = 0;
        m_zipHandle = zip_open(m_tempName.c_str(), ZIP_CREATE, &errorFlag);

        auto count = zip_get_num_entries(m_zipHandle, 0);
        for (int i = 0; i < count; ++i)
            std::cout << zip_get_name(m_zipHandle, i, 0) << std::endl;

        if (errorFlag != ZIP_ER_OK) m_zipHandle = nullptr;
    }

    /**
     * @brief
     */
    void close()
    {
        // ===== zip_discard is called automatically
//        m_zipHandle.reset();

//        auto num = zip_get_num_entries(m_zipHandle, 0);
//        if(m_zipHandle)
//            zip_discard(m_zipHandle);
//        nowide::remove(m_tempName.c_str());    // NOLINT

        if (isOpen()) {
            zip_discard(m_zipHandle);
            m_zipHandle = nullptr;
        }
    }

    /**
     * @brief
     * @param path
     */
    void save(const std::string& path = "")
    {
        auto i = zip_close(m_zipHandle);
        if (i == -1) {
            std::cout << zip_strerror(m_zipHandle) << std::endl;
        }
        m_zipHandle = nullptr;
//        m_zipHandle.reset();

        if (!path.empty()) m_fileName = path;
        nowide::remove(m_fileName.c_str());                        // NOLINT
        nowide::rename(m_tempName.c_str(), m_fileName.c_str());    // NOLINT

        open(m_fileName);
    }

    /**
     * @brief
     * @param name
     * @param data
     */
    void addEntry(const std::string& name, const std::string& data)
    {
//        if (hasEntry(name))
//            deleteEntry(name);
//
//        int sepLocation = name.rfind('/');
//        if (sepLocation != std::string::npos) {
//            auto directory = name.substr(0, sepLocation + 1);
//
//            int nextSlash = directory.find('/');
//            while (nextSlash != std::string::npos) {
//                auto pathToCreate = directory.substr(0, nextSlash + 1);
//                if (!hasEntry(pathToCreate)) {
//                    zip_dir_add(m_zipHandle, pathToCreate.c_str(), ZIP_FL_ENC_GUESS);
//                }
//                nextSlash = directory.find('/', nextSlash + 1);
//            }
//        }
//
        char* sourcedata = new char[data.size()];
        std::strcpy(sourcedata, data.data());
        zip_source* source = zip_source_buffer(m_zipHandle, sourcedata, data.size(), true);
        zip_file_add(m_zipHandle, name.c_str(), source, ZIP_FL_OVERWRITE);

    }

    /**
     * @brief
     * @param entryName
     */
    void deleteEntry(const std::string& entryName) {
        if (hasEntry(entryName)) {
            zip_delete(m_zipHandle, zip_name_locate(m_zipHandle, entryName.c_str(), 0));
        }
    }

    /**
     * @brief
     * @param name
     * @return
     */
    std::string getEntry(const std::string& name) {
        std::string result;

        auto* zipFile = zip_fopen(m_zipHandle, name.c_str(), ZIP_FL_ENC_GUESS);
        if (zipFile) {
            zip_stat_t stat;
            zip_stat(m_zipHandle, name.c_str(), ZIP_FL_ENC_GUESS, &stat);

            char* data = new char[stat.size + 1];
            zip_fread(zipFile, data, stat.size);
            data[stat.size] = '\0';

            result = data;
        }

        return result;
    }

    /**
     * @brief
     * @param entryName
     * @return
     */
    bool hasEntry(const std::string& entryName) {
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
