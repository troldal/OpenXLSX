//
// Created by Kenneth Balslev on 19/05/2020.
//

#ifndef OPENXLSX_XLZIPFILEINTERFACE_HPP
#define OPENXLSX_XLZIPFILEINTERFACE_HPP

#include <Zippy.hpp>

namespace OpenXLSX::Impl {

    /**
     * @brief
     */
    class XLZipFileInterface {

    public:

        /**
         * @brief
         */
        XLZipFileInterface() : m_archive(nullptr) {}

        /**
         * @brief
         */
        ~XLZipFileInterface() = default;

        /**
         * @brief
         * @param other
         */
        XLZipFileInterface(const XLZipFileInterface& other) = delete;

        /**
         * @brief
         * @param other
         */
        XLZipFileInterface(XLZipFileInterface&& other) noexcept = default;

        /**
         * @brief
         * @param lhs
         * @return
         */
        XLZipFileInterface& operator=(const XLZipFileInterface& lhs) = delete;

        /**
         * @brief
         * @param lhs
         * @return
         */
        XLZipFileInterface& operator=(XLZipFileInterface&& lhs) noexcept = default;

        /**
         * @brief
         * @param fileName
         */
        void openArchive(const std::string& fileName) {
            if (m_archive && m_archive->IsOpen()) // TODO: Consider to throw if an archive is already open.
                m_archive->Close();

            m_archive = std::make_unique<Zippy::ZipArchive>();
            m_archive->Open(fileName);
        }

        /**
         * @brief
         */
        void closeArchive() {
            if (m_archive)
                m_archive->Close();
        }

        /**
         * @brief
         * @param fileName
         */
        void saveArchive(const std::string& fileName) {
            m_archive->Save(fileName);
        }

        /**
         * @brief
         * @return
         */
        bool isOpen() {
            return (m_archive && m_archive->IsOpen());
        }

        /**
         * @brief
         * @param path
         * @param content
         */
        void addEntry(const std::string& path, const std::string& content) {
            m_archive->AddEntry(path, content);
        }

        /**
         * @brief
         * @param path
         */
        void deleteEntry(const std::string& path) {
            m_archive->DeleteEntry(path);
        }

        /**
         * @brief
         * @param path
         * @return
         */
        bool hasEntry(const std::string& path) {
            return m_archive->HasEntry(path);
        }

        /**
         * @brief
         * @param path
         * @return
         */
        std::string getEntry(const std::string& path) {
            return m_archive->GetEntry(path).GetDataAsString();
        }

        std::vector<std::string> GetPaths() const {
            return m_archive->GetEntryNames(false, true);
        }

    private:

        std::unique_ptr<Zippy::ZipArchive> m_archive; /**<  */
    };

}  // namespace OpenXLSX::Impl


#endif //OPENXLSX_XLZIPFILEINTERFACE_HPP
