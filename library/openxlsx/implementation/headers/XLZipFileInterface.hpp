//
// Created by Kenneth Balslev on 19/05/2020.
//

#ifndef OPENXLSX_XLZIPFILEINTERFACE_HPP
#define OPENXLSX_XLZIPFILEINTERFACE_HPP

#include <Zippy.h>

namespace OpenXLSX::Impl {

    /**
     * @brief
     */
    class XLZipFileInterface {

    public:
        XLZipFileInterface() : m_archive(nullptr) {}
        ~XLZipFileInterface() = default;
        XLZipFileInterface(const XLZipFileInterface& other) = delete;
        XLZipFileInterface(XLZipFileInterface&& other) noexcept = default;
        XLZipFileInterface& operator=(const XLZipFileInterface& lhs) = delete;
        XLZipFileInterface& operator=(XLZipFileInterface&& lhs) noexcept = default;

        /**
         * @brief
         * @param fileName
         */
        void openArchive(const std::string& fileName) {
            if (m_archive && m_archive->IsOpen()) // TODO: Consider to thow if an archive is already open.
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

        void addEntry(const std::string& path, const std::string& content) {
            m_archive->AddEntry(path, content);
        }

        void deleteEntry(const std::string& path) {
            m_archive->DeleteEntry(path);
        }

        bool hasEntry(const std::string& path) {
            return m_archive->HasEntry(path);
        }

        std::string getEntry(const std::string& path) {
            return m_archive->GetEntry(path).GetDataAsString();
        }

    private:

        std::unique_ptr<Zippy::ZipArchive> m_archive; /**<  */
    };

}  // namespace OpenXLSX::Impl


#endif //OPENXLSX_XLZIPFILEINTERFACE_HPP
