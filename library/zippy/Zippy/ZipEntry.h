//
// Created by Troldal on 2019-03-11.
//

#ifndef Zippy_ZIPENTRY_H
#define Zippy_ZIPENTRY_H

// ===== Standard Library Includes
#include <string>
#include <vector>
#include <iostream>

// ===== Other Includes
#include "miniz/miniz.h"

namespace Zippy
{

    class ZipArchive;

    class ZipEntry;

    // ===== Alias Declarations
    using ZipEntryInfo = mz_zip_archive_file_stat;
    using ZipEntryData = std::vector<std::byte>;

    struct ZipEntryMetaData
    {
        explicit ZipEntryMetaData(const ZipEntryInfo& info)
                : Index(info.m_file_index),
                  CompressedSize(info.m_comp_size),
                  UncompressedSize(info.m_uncomp_size),
                  IsDirectory(info.m_is_directory),
                  IsEncrypted(info.m_is_encrypted),
                  IsSupported(info.m_is_supported),
                  Filename(info.m_filename),
                  Comment(info.m_comment),
                  Time(info.m_time) {

        }

        uint32_t Index;
        uint64_t CompressedSize;
        uint64_t UncompressedSize;
        bool IsDirectory;
        bool IsEncrypted;
        bool IsSupported;
        std::string Filename;
        std::string Comment;
        const time_t Time;
    };

    namespace Impl
    {
        /**
         * @brief
         */
        class ZipEntry
        {
            friend class Zippy::ZipArchive;

            friend class Zippy::ZipEntry;

        public:

            // ===== Constructors, Destructor and Operators

            /**
             * @brief Constructor. Creates a new ZipEntry with the given ZipEntryInfo parameter. This is only used for creating
             * a ZipEntry for an entry already present in the ZipArchive.
             * @param info
             */
            explicit ZipEntry(const ZipEntryInfo& info)
                    : m_EntryInfo(info) {

                GetNewIndex(info.m_file_index);
            }

            /**
             * @brief Constructor. Creates a new ZipEntry with the given name and binary data. This should only be used for creating
             * new entries, not already present in the ZipArchive
             * @param name
             * @param data
             */
            ZipEntry(const std::string& name, const ZipEntryData& data) {

                m_EntryInfo = CreateInfo(name);
                m_EntryData = data;
                m_IsModified = true;

            }

            /**
             * @brief Constructor. Creates a new ZipEntry with the given name and string data. This should only be used for creating
             * new entries, not already present in the ZipArchive
             * @param name
             * @param data
             */
            ZipEntry(const std::string& name, const std::string& data) {

                m_EntryInfo = CreateInfo(name);
                m_EntryData.reserve(data.size());

                for (auto& ch : data)
                    m_EntryData.emplace_back(static_cast<std::byte>(ch));

                m_IsModified = true;
            }

            /**
             * @brief
             * @param other
             */
            ZipEntry(const ZipEntry& other) = delete;

            /**
             * @brief
             * @param other
             */
            ZipEntry(ZipEntry&& other) noexcept = default;

            /**
             * @brief
             */
            virtual ~ZipEntry() = default;

            /**
             * @brief
             * @param other
             * @return
             */
            ZipEntry& operator=(const ZipEntry& other) = delete;

            /**
             * @brief
             * @param other
             * @return
             */
            ZipEntry& operator=(ZipEntry&& other) noexcept = default;


            // ===== Data Access and Manipulation

            /**
             * @brief
             * @return
             */
            ZipEntryData GetData() const {

                return m_EntryData;
            }

            /**
             * @brief
             * @return
             */
            std::string GetDataAsString() const {

                std::string result;
                for (auto& ch : m_EntryData)
                    result += static_cast<char>(ch);

                return result;
            }

            /**
             * @brief
             * @param data
             */
            void SetData(const std::string& data) {

                ZipEntryData result;

                for (auto& ch : data)
                    result.push_back(static_cast<std::byte>(ch));

                m_EntryData = result;
                m_IsModified = true;
            }

            /**
             * @brief
             * @param data
             */
            void SetData(const ZipEntryData& data) {

                m_EntryData = data;
                m_IsModified = true;
            }

            /**
             * @brief
             * @return
             * @todo To be implemented.
             */
            std::string GetName() const {

                return std::string();
            }

            /**
             * @brief
             * @param data
             * @todo To be implemented
             */
            void SetName(const std::string& name) {

            }

            // ===== Metadata Access

            /**
             * @brief
             * @return
             */
            uint32_t Index() const {

                return m_EntryInfo.m_file_index;
            }

            /**
             * @brief
             * @return
             */
            uint64_t CompressedSize() const {

                return m_EntryInfo.m_comp_size;
            }

            /**
             * @brief
             * @return
             */
            uint64_t UncompressedSize() const {

                return m_EntryInfo.m_uncomp_size;
            }

            /**
             * @brief
             * @return
             */
            bool IsDirectory() const {

                return m_EntryInfo.m_is_directory;
            }

            /**
             * @brief
             * @return
             */
            bool IsEncrypted() const {

                return m_EntryInfo.m_is_encrypted;
            }

            /**
             * @brief
             * @return
             */
            bool IsSupported() const {

                return m_EntryInfo.m_is_supported;
            }

            /**
             * @brief
             * @return
             */
            std::string Filename() const {

                return m_EntryInfo.m_filename;
            }

            /**
             * @brief
             * @return
             */
            std::string Comment() const {

                return m_EntryInfo.m_comment;
            }

            /**
             * @brief
             * @return
             */
            const time_t& Time() const {

                return m_EntryInfo.m_time;
            }

        private:
            ZipEntryInfo m_EntryInfo = ZipEntryInfo(); /**< */
            ZipEntryData m_EntryData = ZipEntryData(); /**< */

            bool m_IsModified = false; /**< */

            /**
             * @brief
             * @return
             */
            bool IsModified() const {

                return m_IsModified;
            }

            static uint32_t s_LatestIndex; /**< */

            /**
             * @brief
             * @return
             */
            static uint32_t GetNewIndex(uint32_t latestIndex = 0) {

                static uint32_t index{0};

                if (latestIndex > index) {
                    index = latestIndex;
                    return index;
                }

                return ++index;
            }

            /**
             * @brief
             * @param name
             * @return
             */
            static ZipEntryInfo CreateInfo(const std::string& name) {

                ZipEntryInfo info;

                info.m_file_index = GetNewIndex(0);
                info.m_central_dir_ofs = 0;
                info.m_version_made_by = 0;
                info.m_version_needed = 0;
                info.m_bit_flag = 0;
                info.m_method = 0;
                info.m_time = time(nullptr);
                info.m_crc32 = 0;
                info.m_comp_size = 0;
                info.m_uncomp_size = 0;
                info.m_internal_attr = 0;
                info.m_external_attr = 0;
                info.m_local_header_ofs = 0;
                info.m_comment_size = 0;
                info.m_is_directory = (name.back() == '/');
                info.m_is_encrypted = false;
                info.m_is_supported = true;

                strcpy(info.m_filename, name.c_str());
                strcpy(info.m_comment, "");

                return info;
            }

        };
    }  // namespace Impl

    /**
     * @brief
     */
    class ZipEntry
    {
        friend class ZipArchive;

    public:

        // ===== Constructors, Destructor and Operators

        /**
         * @brief Constructor. Creates a new ZipEntry with the given ZipEntryInfo parameter. This is only used for creating
         * a ZipEntry for an entry already present in the ZipArchive.
         * @param entry
         */
        explicit ZipEntry(Impl::ZipEntry* entry)
                : m_ZipEntry(entry) {

        }

        /**
         * @brief
         * @param other
         */
        ZipEntry(const ZipEntry& other) = default;

        /**
         * @brief
         * @param other
         */
        ZipEntry(ZipEntry&& other) noexcept = default;

        /**
         * @brief
         */
        virtual ~ZipEntry() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        ZipEntry& operator=(const ZipEntry& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        ZipEntry& operator=(ZipEntry&& other) noexcept = default;


        // ===== Data Access and Manipulation

        /**
         * @brief
         * @return
         */
        ZipEntryData GetData() const {

            return m_ZipEntry->GetData();
        }

        /**
         * @brief
         * @return
         */
        std::string GetDataAsString() const {

            return m_ZipEntry->GetDataAsString();
        }

        /**
         * @brief
         * @param data
         */
        void SetData(const std::string& data) {

            m_ZipEntry->SetData(data);
        }

        /**
         * @brief
         * @param data
         */
        void SetData(const ZipEntryData& data) {

            m_ZipEntry->SetData(data);
        }


        // ===== Metadata Access

        /**
         * @brief
         * @return
         */
        uint32_t Index() const {

            return m_ZipEntry->Index();
        }

        /**
         * @brief
         * @return
         */
        uint64_t CompressedSize() const {

            return m_ZipEntry->CompressedSize();
        }

        /**
         * @brief
         * @return
         */
        uint64_t UncompressedSize() const {

            return m_ZipEntry->UncompressedSize();
        }

        /**
         * @brief
         * @return
         */
        bool IsDirectory() const {

            return m_ZipEntry->IsDirectory();
        }

        /**
         * @brief
         * @return
         */
        bool IsEncrypted() const {

            return m_ZipEntry->IsEncrypted();
        }

        /**
         * @brief
         * @return
         */
        bool IsSupported() const {

            return m_ZipEntry->IsSupported();
        }

        /**
         * @brief
         * @return
         */
        std::string Filename() const {

            return m_ZipEntry->Filename();
        }

        /**
         * @brief
         * @return
         */
        std::string Comment() const {

            return m_ZipEntry->Comment();
        }

        /**
         * @brief
         * @return
         */
        const time_t& Time() const {

            return m_ZipEntry->Time();
        }

    private:

        /**
         * @brief
         * @return
         */
        bool IsModified() const {

            return m_ZipEntry->IsModified();
        }

        Impl::ZipEntry* m_ZipEntry;  /**< */

    };

}  // namespace Zippy


#endif //Zippy_ZIPENTRY_H
