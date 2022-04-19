//
// Created by Kenneth Balslev on 18/04/2022.
//

#ifndef OPENXLSX_CUSTOMZIP_H
#define OPENXLSX_CUSTOMZIP_H

#include <fstream>
#include <libzippp.h>
#include <random>
#include <string>
#include <nowide/cstdio.hpp>

class CustomZip
{
public:
    /**
     * @brief
     */
    CustomZip() : m_archive(nullptr) {}

    /**
     * @brief
     * @param other
     */
    CustomZip(const CustomZip& other) = default;

    /**
     * @brief
     * @param other
     */
    CustomZip(CustomZip&& other) = default;

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
    CustomZip& operator=(CustomZip&& other) = default;

    /**
     * @brief
     * @return
     */
    explicit operator bool() const { return isValid(); }

    bool isValid() const { return m_archive != nullptr; }

    /**
     * @brief
     * @return
     */
    bool isOpen() const { return m_archive && m_archive->isOpen(); }

    /**
     * @brief
     * @param fileName
     */
    void open(const std::string& fileName)
    {
        m_fileName = fileName;
        m_tempName = m_fileName.substr(0, m_fileName.rfind('/') + 1) + GenerateRandomName(20);

        std::ifstream src(m_fileName, std::ios::binary);
        std::ofstream dst(m_tempName, std::ios::binary);
        dst << src.rdbuf();
        src.close();
        dst.close();

        m_archive = std::make_shared<libzippp::ZipArchive>(m_tempName);
        m_archive->open(libzippp::ZipArchive::Write);
    }

    /**
     * @brief
     */
    void close()
    {
        m_archive->discard();
        m_archive = nullptr;
        nowide::remove(m_tempName.c_str()); // NOLINT
    }

    /**
     * @brief
     * @param path
     */
    void save(const std::string& path = "")
    {
        m_archive->close();
        m_archive = nullptr;

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
    void addEntry(const std::string& name, const std::string& data) {
        char* dst = new char [data.size() + 1];
        std::strcpy(dst, data.c_str());
        if (hasEntry(name))
            m_archive->deleteEntry(name);
        m_archive->addData(name, dst, data.size(), true);
    }

    /**
     * @brief
     * @param entryName
     */
    void deleteEntry(const std::string& entryName) {
        m_archive->deleteEntry(entryName);
    }

    /**
     * @brief
     * @param name
     * @return
     */
    std::string getEntry(const std::string& name) {
        return m_archive->getEntry(name).readAsText();
    }

    /**
     * @brief
     * @param entryName
     * @return
     */
    bool hasEntry(const std::string& entryName) {
        return m_archive->hasEntry(entryName);
    }

private:
    inline std::string GenerateRandomName(int length)
    {
        std::string letters = "abcdefghijklmnopqrstuvwxyz0123456789";

        std::random_device                 rand_dev;
        std::mt19937                       generator(rand_dev());
        std::uniform_int_distribution<int> distr(0, letters.size() - 1);

        std::string result;
        for (int i = 0; i < length; ++i) {
            result += letters[distr(generator)];
        }

        return result + ".tmp";
    }

    std::shared_ptr<libzippp::ZipArchive> m_archive; /**< */
    std::string                           m_fileName;
    std::string                           m_tempName;
};

#endif    // OPENXLSX_CUSTOMZIP_H
