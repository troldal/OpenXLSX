#ifndef ZIPPYHEADERONLY_ZIPPY_HPP
#define ZIPPYHEADERONLY_ZIPPY_HPP

#pragma warning(push)
#pragma warning(disable : 4334)
#pragma warning(disable : 4996)
#pragma warning(disable : 4127)
#pragma warning(disable : 4267)
#pragma warning(disable : 4100)
#pragma warning(disable : 4245)
#pragma warning(disable : 4267)
#pragma warning(disable : 4242)
#pragma warning(disable : 4244)

#include <algorithm>
#include <cstddef>
#include <cstdio>
#include <fstream>
#include <iostream>
#include <memory>
#include <random>
#include <stdexcept>
#include <string>
#include <vector>
#include <utility>
#include <optional>

// TODO(troldal): Is this required???
#ifdef _WIN32
#    include <direct.h>
#endif

#include <filesystem>
#include <nowide/cstdio.hpp>

#include "miniz.h"

namespace KZip
{

    namespace fs = std::filesystem;

    /**
     * @brief The ZipRuntimeError class is a custom exception class derived from the std::runtime_error class.
     * @details In case of an error in the KZip library, a ZipRuntimeError object will be thrown, with a message
     * describing the details of the error.
     */
    class ZipRuntimeError : public std::runtime_error
    {
    public:
        /**
         * @brief Constructor.
         * @param err A string with a description of the error.
         */
        inline explicit ZipRuntimeError(const std::string& err) : runtime_error(err) {}
    };

    /**
    * @brief The ZipLogicError class is a custom exception class derived from the std::logic_error class.
    * @details In case of an error in the input, a ZipLogicError object will be thrown, with a message
    * describing the details of the error.
     */
    class ZipLogicError : public std::logic_error
    {
    public:
        /**
         * @brief
         * @param err
         */
        inline explicit ZipLogicError(const std::string& err) : logic_error(err) {}
    };

}    // namespace KZip

namespace KZip::Impl
{
    /**
     * @brief Generates a random filename, which is used to generate a temporary archive when modifying and saving
     * archive files.
     * @param length The length of the filename to create.
     * @return Returns the generated filenamen, appended with '.tmp'.
     */
    inline std::string GenerateRandomName(int length)
    {
        std::string letters = "abcdefghijklmnopqrstuvwxyz0123456789";

        std::random_device                 rand_dev;
        std::mt19937                       generator(rand_dev());
        std::uniform_int_distribution<int> distr(0, letters.size() - 1);

        std::string result = ".~$";
        for (int i = 0; i < length; ++i) {
            result += letters[distr(generator)];
        }

        return result + ".tmp";
    }

}    // namespace KZip::Impl

namespace KZip
{
    class ZipArchive;
    class ZipEntry;

    // ===== Alias Declarations

    /**
     * @brief The ZipEntryInfo entity is an alias of the mz_zip_archive_file_stat from the miniz library.
     * @details The ZipEntryInfo/mz_zip_archive_file_stat struct holds various meta data related to a particular
     * entry (or item) in a zip archive, such as: comments, file size, date stamp etc.
     * @note A new ZipEntryInfo should not be created manually.
     */
//    using ZipEntryInfo = mz_zip_archive_file_stat;

    /**
     * @brief The ZipEntryData entity is an alias for a std::vector of std::bytes.
     * @details This is used as a generic container of file data of any kind, both character strings and binary.
     * A vector of char or an array of char can also be used, but a vector of bytes makes it clearer that it can
     * also be used for non-text data.
     */
    using ZipEntryData = std::vector<unsigned char>;

    /**
     * @brief The ZipEntryMetaData is essentially a wrapper around the ZipEntryInfo struct, which is an alias for a
     * miniz struct.
     */
    struct ZipEntryMetaData // NOLINT
    {
        /**
         * @brief Constructor.
         * @param info A reference to a ZipEntryInfo object.
         */
        explicit ZipEntryMetaData(const mz_zip_archive_file_stat& info)
            : Index(info.m_file_index),
              CompressedSize(info.m_comp_size),
              UncompressedSize(info.m_uncomp_size),
              IsDirectory(info.m_is_directory),
              IsEncrypted(info.m_is_encrypted),
              IsSupported(info.m_is_supported),
              Filename(info.m_filename),
              Comment(info.m_comment),
              Time(info.m_time)
        {}

        uint32_t     Index;            /**< */
        uint64_t     CompressedSize;   /**< */
        uint64_t     UncompressedSize; /**< */
        bool         IsDirectory;      /**< */
        bool         IsEncrypted;      /**< */
        bool         IsSupported;      /**< */
        std::string  Filename;         /**< */
        std::string  Comment;          /**< */
        const time_t Time;             /**< */
    };

    namespace Impl
    {
        /**
         * @brief The Impl::ZipEntry class implements the functionality required for manipulating entries in a zip archive.
         * @details This is the implementation class. The ZipEntry class in the KZip namespace implements the public interface.
         * The reason for having a separate implementation class, is that it enables easy copy and move operations of
         * ZipEntry objects, without duplicating the underlying implementation objects.
         */
        class ZipEntry
        {
            friend class KZip::ZipArchive;
            friend class KZip::ZipEntry;

        public:
            // ===== Constructors, Destructor and Operators

            /**
             * @brief Constructor. Creates a new ZipEntry with the given ZipEntryInfo parameter. This is only used for creating
             * a ZipEntry for an entry already present in the ZipArchive, when opening an existing archive.
             * @param info A reference to a ZipEntryInfo object with the entry info.
             */
            explicit ZipEntry(const mz_zip_archive_file_stat& info) : m_entryInfo(info)
            {
                // ===== Call GetNewIndex to update the index counter with the getValue in the ZipEntryInfo object.
                getNewIndex(info.m_file_index);    // The return getValue is deliberately not used.
            }

            /**
             * @brief Constructor. Creates a new ZipEntry with the given name and binary data. This should only be used for creating
             * new entries, not already present in the ZipArchive
             * @param name The name of the new entry to add to the zip archive.
             * @param data A reference to a ZipEntryData object with the file data to add.
             */
            ZipEntry(const std::string& name, const ZipEntryData& data) : m_entryInfo(createInfo(name)),
                                                                          m_entryData(data),
                                                                          m_isModified(true)

            {}

            /**
             * @brief Constructor. Creates a new ZipEntry with the given name and string data. This should only be used for creating
             * new entries, not already present in the ZipArchive
             * @param name The name of the new entry to add to the zip archive.
             * @param data A string with the text data to add to the zip archive.
             */
            ZipEntry(const std::string& name, const std::string& data) : m_entryInfo(createInfo(name)),
                                                                         m_entryData(std::vector<unsigned char> (data.begin(), data.end())),
                                                                         m_isModified(true)
            {}

            /**
             * @brief Copy Constructor (Deleted)
             * @param other The object to copy.
             * @note The copy constructor has been deleted, because it is not obvious what should happen to the
             * underlying file when copying the entry object.
             */
            ZipEntry(const ZipEntry& other) = delete;

            /**
             * @brief Move Constructor.
             * @param other The object to be moved.
             */
            ZipEntry(ZipEntry&& other) noexcept = default;

            /**
             * @brief Destructor.
             */
            virtual ~ZipEntry() = default;

            /**
             * @brief Copy Assignment Operator (deleted)
             * @param other The object to be copied.
             * @return A reference to the copied-to object.
             * @note The copy assignment operator has been deleted, because it is not obvious what should happen to the
             * underlying file when copying the entry object.
             */
            ZipEntry& operator=(const ZipEntry& other) = delete;

            /**
             * @brief Move Assignment Operator.
             * @param other The object to be moved.
             * @return A reference to the moved-to object.
             */
            ZipEntry& operator=(ZipEntry&& other) noexcept = default;

            // ===== Data Access and Manipulation

            /**
             * @brief Get data from entry.
             * @return Returns a ZipEntryData (alias for std::vector<std::byte>) object with the fil data.
             */
            inline ZipEntryData getData() const
            {
                return m_entryData;
            }

            /**
             * @brief Get entry data as a std::string.
             * @return Returns a std::string with the file data.
             */
            inline std::string getDataAsString() const
            {
                return std::string{m_entryData.begin(), m_entryData.end()};
            }

            /**
             * @brief Set the data for the entry.
             * @param data A std::string with the file data.
             */
            inline void setData(const std::string& data)
            {
                setData(ZipEntryData{data.begin(), data.end()});
            }

            /**
             * @brief Set the data for the entry.
             * @param data A ZipEntryData (alias for std::vector<std::byte>) object with the file data.
             */
            inline void setData(const ZipEntryData& data)
            {
                m_entryData  = data;
                m_isModified = true;
            }

            /**
             * @brief Get the name of the entry.
             * @return Returns a std::string with the entry name.
             */
            inline std::string getName() const
            {
                return m_entryInfo.m_filename;
            }

            /**
             * @brief Set the name of the entry.
             * @param name A std::string with the new name of the entry.
             */
            inline void setName(const std::string& name)
            {
#if _MSC_VER    // On MSVC, use the safe version of strcpy
                strcpy_s(m_entryInfo.m_filename, sizeof m_entryInfo.m_filename, name.c_str());
#else    // Otherwise, use the unsafe version as fallback :(
                strncpy(m_entryInfo.m_filename, name.c_str(), sizeof m_entryInfo.m_filename);
#endif
            }

            // ===== Metadata Access

            /**
             * @brief Get the entry's index in the zip archive.
             * @return Returns a uint32_t with the entry index.
             */
            inline uint32_t index() const
            {
                return m_entryInfo.m_file_index;
            }

            /**
             * @brief Get the uncompressed size of the entry (in bytes)
             * @return A uint64_t with the uncompressed size.
             */
            inline uint64_t compressedSize() const
            {
                return m_entryInfo.m_comp_size;
            }

            /**
             * @brief Get the compressed size of the entry (in bytes)
             * @return A uint64_t with the compressed size.
             */
            inline uint64_t uncompressedSize() const
            {
                return m_entryInfo.m_uncomp_size;
            }

            /**
             * @brief Is the zip entry a directory?
             * @return Returns true if the entry is a directory; otherwise false.
             */
            inline bool isDirectory() const
            {
                return m_entryInfo.m_is_directory;
            }

            /**
             * @brief
             * @return Returns true if the entry is encrypted; otherwise false.
             */
            inline bool isEncrypted() const
            {
                return m_entryInfo.m_is_encrypted;
            }

            /**
             * @brief Is the zip entry encryption supported?
             * @return Returns true if the entry encryption (if any) is supported by the library; otherwise false.
             */
            inline bool isSupported() const
            {
                return m_entryInfo.m_is_supported;
            }

            /**
             * @brief Get the zip entry comments.
             * @return Returns a std::string with the comments.
             */
            inline std::string comment() const
            {
                return std::string(m_entryInfo.m_comment, m_entryInfo.m_comment_size); // NOLINT
            }

            /**
             * @brief Get the zip entry time stamp.
             * @return Returns a time_t object with the time stamp.
             */
            inline const time_t& time() const
            {
                return m_entryInfo.m_time;
            }

            /**
             * @brief Has the zip entry been modified?
             * @return Returns true if the entry is has been modified; otherwise false.
             */
            inline bool isModified() const
            {
                return m_isModified;
            }

        private:
            mz_zip_archive_file_stat m_entryInfo = mz_zip_archive_file_stat(); /**< The zip entry metadata. */
            mutable ZipEntryData m_entryData = ZipEntryData(); /**< The zip entry data. */
            mz_zip_archive* m_archive = nullptr;
            bool m_isModified = false; /**< Boolean flag indicating if the file has been modified since opening. */

            /**
             * @brief Generate a new file index.
             * @details The file index in existing zip archives may not be incrementing trivially. When opening existing
             * zip archives, this function is simply used to update the index. When adding new entries to an existing
             * archive, this function is guaranteed to provide a unique index.
             * @return Returns a uint32_t (32 bit unsigned int) with the new index.
             */
            static uint32_t getNewIndex(uint32_t latestIndex = 0)
            {
                // ===== Set up a static index counter (set to zero the first time the function is executed)
                static uint32_t index { 0 };

                // ===== If the input value is larger than the current index getValue, set the index equal to the input.
                if (latestIndex > index) {
                    index = latestIndex;
                    return index;
                }

                // ===== Increment the index and return the getValue.
                return ++index;
            }

            /**
             * @brief Create a new ZipEntryInfo structure.
             * @details This function will create a new ZipEntryInfo structure, based on the input file name. The
             * structure values will mostly be dummy values, except for the file index, the time stamp, the file name
             * and the is_directory flag.
             * @param name The file name for the new ZipEntryInfo.
             * @return The newly created ZipEntryInfo structure is returned
             */
            static mz_zip_archive_file_stat createInfo(const std::string& name)
            {
                mz_zip_archive_file_stat info;

                info.m_file_index       = getNewIndex(0);
                info.m_central_dir_ofs  = 0;
                info.m_version_made_by  = 0;
                info.m_version_needed   = 0;
                info.m_bit_flag         = 0;
                info.m_method           = 0;
                info.m_time             = std::time(nullptr);
                info.m_crc32            = 0;
                info.m_comp_size        = 0;
                info.m_uncomp_size      = 0;
                info.m_internal_attr    = 0;
                info.m_external_attr    = 0;
                info.m_local_header_ofs = 0;
                info.m_comment_size     = 0;
                info.m_is_directory     = (name.back() == '/');
                info.m_is_encrypted     = false;
                info.m_is_supported     = true;

#if _MSC_VER    // On MSVC, use the safe version of strcpy
                strcpy_s(info.m_filename, sizeof info.m_filename, name.c_str());
                strcpy_s(info.m_comment, sizeof info.m_comment, "");
#else    // Otherwise, use the unsafe version as fallback :(
                strncpy(info.m_filename, name.c_str(), sizeof info.m_filename);
                strncpy(info.m_comment, "", sizeof info.m_comment);
#endif

                return info;
            }
        };
    }    // namespace Impl

    /**
     * @brief The ZipEntry class implements the interface required for manipulating entries in a zip archive.
     * @details This is the interface class. The ZipEntry class in the KZip::Impl namespace implements the private implementation.
     * The reason for having a separate implementation class, is that it enables easy copy and move operations of
     * ZipEntry objects, without duplicating the underlying implementation objects.
     */
    class ZipEntry
    {
        friend class ZipArchive;

    public:
        // ===== Constructors, Destructor and Operators

        /**
         * @brief Constructor. Creates a new ZipEntry.
         * @param entry A raw (non-owning) pointer to the existing Impl::ZipEntry object.
         * @todo Can this be made private?
         */
        explicit ZipEntry(Impl::ZipEntry* entry) : m_zipEntry(entry) {}

        /**
         * @brief Copy Constructor.
         * @param other The object to be copied.
         */
        ZipEntry(const ZipEntry& other) = default;

        /**
         * @brief Move Constructor.
         * @param other The object to be moved.
         */
        ZipEntry(ZipEntry&& other) noexcept = default;

        /**
         * @brief Destructor
         */
        virtual ~ZipEntry() {
            // TODO: What happens if two identical objects, and one is destroyed before the other?
            if (!m_zipEntry->isModified()) m_zipEntry->m_entryData.clear();
        };

        /**
         * @brief Copy Assignment Operator.
         * @param other The object to be copied.
         * @return A reference to the copied-to object.
         */
        ZipEntry& operator=(const ZipEntry& other) = default;

        /**
         * @brief Move Assignment Operator.
         * @param other The object to be moved.
         * @return A reference to the moved-to object.
         */
        ZipEntry& operator=(ZipEntry&& other) noexcept = default;

        // ===== Data Access and Manipulation

        /**
         * @brief Get data from entry.
         * @return Returns a ZipEntryData (alias for std::vector<std::byte>) object with the file data.
         */
        inline ZipEntryData getData() const
        {
            return m_zipEntry->getData();
        }

        /**
         * @brief Get entry data as a std::string.
         * @return Returns a std::string with the file data.
         */
        inline std::string getDataAsString() const
        {
            return m_zipEntry->getDataAsString();
        }

        /**
         * @brief Set the data for the entry.
         * @param data A std::string with the file data.
         */
        inline void setData(const std::string& data)
        {
            m_zipEntry->setData(data);
        }

        /**
         * @brief Set the data for the entry.
         * @param data A ZipEntryData (alias for std::vector<std::byte>) object with the file data.
         */
        inline void setData(const ZipEntryData& data)
        {
            m_zipEntry->setData(data);
        }

        // ===== Metadata Access

        /**
         * @brief Get the entry's index in the zip archive.
         * @return Returns a uint32_t with the entry index.
         */
        inline uint32_t index() const
        {
            return m_zipEntry->index();
        }

        /**
         * @brief Get the uncompressed size of the entry (in bytes)
         * @return A uint64_t with the uncompressed size.
         */
        inline uint64_t compressedSize() const
        {
            return m_zipEntry->compressedSize();
        }

        /**
         * @brief Get the compressed size of the entry (in bytes)
         * @return A uint64_t with the compressed size.
         */
        inline uint64_t uncompressedSize() const
        {
            return m_zipEntry->uncompressedSize();
        }

        /**
         * @brief Is the entry a directory?
         * @return Returns true if the entry is a directory; otherwise false.
         */
        inline bool isDirectory() const
        {
            return m_zipEntry->isDirectory();
        }

        /**
         * @brief Is the zip entry supported?
         * @return Returns true if the entry is encrypted; otherwise false.
         */
        inline bool isEncrypted() const
        {
            return m_zipEntry->isEncrypted();
        }

        /**
         * @brief Is the zip entry encryption supported?
         * @return Returns true if the entry encryption (if any) is supported; otherwise false.
         */
        inline bool isSupported() const
        {
            return m_zipEntry->isSupported();
        }

        /**
         * @brief Get the entry filename.
         * @return Returns a std::string with the filename.
         * @todo Consider renaming to GetName and add a SetName member function.
         */
        inline std::string filename() const
        {
            return m_zipEntry->getName();
        }

        /**
         * @brief Get the zip entry comments.
         * @return Returns a std::string with the comments.
         */
        inline std::string comment() const
        {
            return m_zipEntry->comment();
        }

        /**
         * @brief Get the entry time stamp.
         * @return Returns a time_t object with the time stamp.
         */
        inline const time_t& time() const
        {
            return m_zipEntry->time();
        }

    private:
        /**
         * @brief Has the file been modified?
         * @return Returns true if the entry has been modified; otherwise false.
         */
        inline bool isModified() const
        {
            return m_zipEntry->isModified();
        }

        Impl::ZipEntry* m_zipEntry; /**< A raw (non-owning) pointer to the implementation object. */
    };

    /**
     *
     */
    class ZipEntry2
    {

        friend class ZipArchive;

    public:
        ~ZipEntry2() = default;



        template<
            typename T,
            typename std::enable_if<std::is_convertible_v<typename T::value_type, unsigned char> >::type* = nullptr>
        ZipEntry2& operator=(T data) {
            // TODO(troldal): To be implemented
        }

        template<
            typename T,
            typename std::enable_if<std::is_convertible_v<typename T::value_type, unsigned char> >::type* = nullptr>
        void setData(T data) {
            // TODO(troldal): To be implemented
        }

        /**
         * @brief Function for extracting the zip data to any compatible container.
         * @details This templated getter allows extraction of the zip data to any container
         * that holds a type convertible from unsigned char, and that can be constructed from begin and end
         * iterators from a source.
         * @note While any iterator-based container with an unsigned char value type can be used,
         * the function is optimized for std::string and std::vector<unsigned char>
         * @tparam T The type to convert to (e.g. std::string or std::vector<unsigned char>
         * @return An object of type T, holding the zip data.
         */
        template<
            typename T,
            typename std::enable_if<std::is_convertible_v<typename T::value_type, unsigned char> >::type* = nullptr>
        T getData() const {

            // ===== Optimization for std::string
            if constexpr (std::is_same_v<T, std::string>) {
                // ===== Create a temporary vector of unsinged char, to hold the zip data
                std::string data;
                data.resize(m_info.m_uncomp_size);
                mz_zip_reader_extract_to_mem(m_archive, m_info.m_file_index, data.data(), m_info.m_uncomp_size, 0);

                // ===== Check that the operation was successful
                if (!m_info.m_is_directory && data.size() != m_info.m_uncomp_size) {
                    throw ZipRuntimeError(mz_zip_get_error_string(m_archive->m_last_error));
                }

                return data;
            }

            // ===== Optimization for std::vector<unsigned char>
            else if constexpr (std::is_same_v<T, std::vector<unsigned char> >) {
                // ===== Create a temporary vector of unsinged char, to hold the zip data
                std::vector<unsigned char> data;
                data.resize(m_info.m_uncomp_size);
                mz_zip_reader_extract_to_mem(m_archive, m_info.m_file_index, data.data(), m_info.m_uncomp_size, 0);

                // ===== Check that the operation was successful
                if (!m_info.m_is_directory && data.size() != m_info.m_uncomp_size) {
                    throw ZipRuntimeError(mz_zip_get_error_string(m_archive->m_last_error));
                }

                return data;
            }

            // ===== General case
            else {
                // ===== Create a temporary vector of unsinged char, to hold the zip data
                std::vector<unsigned char> data;
                data.resize(m_info.m_uncomp_size);
                mz_zip_reader_extract_to_mem(m_archive, m_info.m_file_index, data.data(), m_info.m_uncomp_size, 0);

                // ===== Check that the operation was successful
                if (!m_info.m_is_directory && data.size() != m_info.m_uncomp_size) {
                    throw ZipRuntimeError(mz_zip_get_error_string(m_archive->m_last_error));
                }

                return {data.begin(), data.end()};
            }
        }

        /**
         * @brief Implicit type conversion operator.
         * @details This templated type conversion operator allows extraction of the zip data to any container
         * that holds a type convertible from unsigned char, and that can be constructed from begin and end
         * iterators from a source. This function simply calls the getData<T>() function.
         * @tparam T The type to convert to (e.g. std::string or std::vector<unsigned char>
         * @return An object of type T, holding the zip data.
         */
        template<
            typename T,
            typename std::enable_if<std::is_convertible_v<unsigned char, typename T::value_type> >::type* = nullptr>
        operator T() const { // NOLINT
            return getData<T>();
        }

        operator ZipEntry() { // NOLINT
            // TODO(troldal): To be implemented
        }

        void erase() {
            // TODO(troldal): To be implemented
        }

        ZipEntryMetaData metadata() const {
            // TODO(troldal): To be implemented
        }

    private:
        //---------- Private Member Functions ---------- //

        ZipEntry2(mz_zip_archive* archive, mz_zip_archive_file_stat info)
            : m_archive(archive),
              m_info(info)
        {}

        ZipEntry2(const ZipEntry2& other) = default;

        ZipEntry2(ZipEntry2&& other) noexcept = default;

        ZipEntry2& operator=(const ZipEntry2& other) = default;

        ZipEntry2& operator=(ZipEntry2&& other) noexcept = default;


        //---------- Private Member Variables ---------- //

        mz_zip_archive* m_archive = nullptr; /**< Non-owning pointer to the mz_zip_archive struct. */
        mz_zip_archive_file_stat m_info = mz_zip_archive_file_stat(); /**< File stats for the entry. */
    };

    enum class ZipFlags : uint8_t
    {
        IncludeFiles   = 1,
        IncludeDirectories     = 2
    };

    inline ZipFlags operator|(ZipFlags first, ZipFlags second)
    {
        return static_cast<ZipFlags>(static_cast<uint8_t>(first) | static_cast<uint8_t>(second));
    }

    inline ZipFlags operator&(ZipFlags first, ZipFlags second)
    {
        return static_cast<ZipFlags>(static_cast<uint8_t>(first) & static_cast<uint8_t>(second));
    }

    namespace Impl {
        class ZipArchive {
            using Data = std::pair<mz_zip_archive_file_stat, std::optional<std::vector<unsigned char> > >;

        public:

            ZipArchive() = default;

            explicit ZipArchive(const fs::path& fileName) : m_archivePath(fileName) {
                //TODO (troldal): To be implemented;
            }

            ZipArchive(const ZipArchive& other) = delete;

            ZipArchive(ZipArchive&& other) noexcept = default;

            ~ZipArchive() {
                //TODO (troldal): To be implemented;
            }

            ZipArchive& operator=(const ZipArchive& other) = delete;

            ZipArchive& operator=(ZipArchive&& other) = default;

            inline void create(const fs::path& fileName) {
                //TODO (troldal): To be implemented;
            }

            inline void open(const fs::path& fileName) {

                if (isOpen()) close();

                // ===== Open the zip archive file. If unsuccessful, throw exception.
                m_archivePath = fileName;
                if (!mz_zip_reader_init_file(&m_archive, m_archivePath.c_str(), 0)) {
                    throw ZipRuntimeError(mz_zip_get_error_string(m_archive.m_last_error));
                }
                m_isOpen = true;

                // ===== Iterate through the archive and add the entries to the internal data structure
                mz_zip_archive_file_stat info;
                for (unsigned int i = 0; i < mz_zip_reader_get_num_files(&m_archive); i++) {
                    if (!mz_zip_reader_file_stat(&m_archive, i, &info)) {
                        throw ZipRuntimeError(mz_zip_get_error_string(m_archive.m_last_error));
                    }

                    m_zipEntryData.emplace_back(std::make_pair(info, std::optional<std::vector<unsigned char> >()));
                }

                // ===== Remove entries with identical names. The newest entries will be retained.
                // TODO (troldal): Consider using rbegin;rend instead of reversing.
                auto isEqual = [](const Data& a, const Data& b) { return strcmp(a.first.m_filename, b.first.m_filename) == 0; };
                std::reverse(m_zipEntryData.begin(), m_zipEntryData.end());
                m_zipEntryData.erase(std::unique(m_zipEntryData.begin(), m_zipEntryData.end(), isEqual), m_zipEntryData.end());
                std::reverse(m_zipEntryData.begin(), m_zipEntryData.end());

                // ===== Add folder entries if they don't exist
                for (auto& entry : entryNames(ZipFlags::IncludeDirectories)) {
                    if (entry.find('/') != std::string::npos) {
                        addEntry(entry.substr(0, entry.rfind('/') + 1));
                    }
                }
            }

            inline ZipEntry addEntry(const std::string& name, const ZipEntryData& data = {})
            {
                if (!isOpen()) throw ZipLogicError("Cannot call AddEntry on an empty ZipArchive object!");

                // ===== Ensure that all folders and subfolders have an entry in the archive
                auto folders = entryNames(ZipFlags::IncludeDirectories);
                auto position = uint64_t {0};
                while (name.find('/', position) != std::string::npos) {
                    position    = name.find('/', position) + 1;
                    auto folder = name.substr(0, position);

                    // ===== If folder isn't registered in the archive, add it.
                    if (std::find(folders.begin(), folders.end(), folder) == folders.end()) {
                        m_zipEntries.emplace_back(Impl::ZipEntry(folder, ""));
                        folders.emplace_back(folder);
                    }
                }

                // ===== Check if an entry with the given name already exists in the archive.
                auto result = std::find_if(m_zipEntries.begin(), m_zipEntries.end(), [&](const Impl::ZipEntry& entry) {
                    return name == entry.getName();
                });

                // ===== If the entry exists, replace the existing data with the new data, and return the ZipEntry object.
                if (result != m_zipEntries.end()) {
                    result->setData(data);
                    return ZipEntry(&*result);
                }

                // ===== Finally, add a new entry with the given name and data, and return the object.
                return ZipEntry(&m_zipEntries.emplace_back(Impl::ZipEntry(name, data)));
            }

            /**
             * @brief Get a list of the entries in the archive. Depending on the input parameters, the list will include
             * directories, files or both.
             * @param includeDirs If true, the list will include directories; otherwise not. Default is true
             * @param includeFiles If true, the list will include files; otherwise not. Default is true
             * @return A std::vector of std::strings with the entry names.
             */
            inline std::vector<std::string> entryNames(ZipFlags flags = (ZipFlags::IncludeFiles | ZipFlags::IncludeDirectories)) const
            {
                if (!isOpen()) throw ZipLogicError("Cannot call GetEntryNames on empty ZipArchive object!");

                std::vector<std::string> result;

                // ===== Iterate through all the entries in the archive
                for (const auto& item : m_zipEntryData) {
                    // ===== If directories should be included and the current entry is a directory, add it to the result.
                    if (static_cast<bool>(flags & ZipFlags::IncludeDirectories) && item.first.m_is_directory) {
                        result.emplace_back(item.first.m_filename);
                        continue;
                    }

                    // ===== If files should be included and the current entry is a file, add it to the result.
                    if (static_cast<bool>(flags & ZipFlags::IncludeFiles) && !item.first.m_is_directory) {
                        result.emplace_back(item.first.m_filename);
                        continue;
                    }
                }

                return result;
            }

            inline void close() {
                //TODO (troldal): To be implemented;
            }

            inline bool isOpen() const {
                //TODO (troldal): To be implemented;
            }

            inline void save(fs::path filename = "") {
                //TODO (troldal): To be implemented;
            }

        private:

            /**
             * @brief Generate a new file index.
             * @details The file index in existing zip archives may not be incrementing trivially. When opening existing
             * zip archives, this function is simply used to update the index. When adding new entries to an existing
             * archive, this function is guaranteed to provide a unique index.
             * @return Returns a uint32_t (32 bit unsigned int) with the new index.
             */
            static uint32_t getNewIndex(uint32_t latestIndex = 0)
            {
                // ===== Set up a static index counter (set to zero the first time the function is executed)
                static uint32_t index { 0 };

                // ===== If the input value is larger than the current index getValue, set the index equal to the input.
                if (latestIndex > index) {
                    index = latestIndex;
                    return index;
                }

                // ===== Increment the index and return the getValue.
                return ++index;
            }

            /**
             * @brief Create a new ZipEntryInfo structure.
             * @details This function will create a new ZipEntryInfo structure, based on the input file name. The
             * structure values will mostly be dummy values, except for the file index, the time stamp, the file name
             * and the is_directory flag.
             * @param name The file name for the new ZipEntryInfo.
             * @return The newly created ZipEntryInfo structure is returned
             */
            static mz_zip_archive_file_stat createInfo(const std::string& name)
            {
                mz_zip_archive_file_stat info;

                info.m_file_index       = getNewIndex(0);
                info.m_central_dir_ofs  = 0;
                info.m_version_made_by  = 0;
                info.m_version_needed   = 0;
                info.m_bit_flag         = 0;
                info.m_method           = 0;
                info.m_time             = std::time(nullptr);
                info.m_crc32            = 0;
                info.m_comp_size        = 0;
                info.m_uncomp_size      = 0;
                info.m_internal_attr    = 0;
                info.m_external_attr    = 0;
                info.m_local_header_ofs = 0;
                info.m_comment_size     = 0;
                info.m_is_directory     = (name.back() == '/');
                info.m_is_encrypted     = false;
                info.m_is_supported     = true;

#if _MSC_VER    // On MSVC, use the safe version of strcpy
                strcpy_s(info.m_filename, sizeof info.m_filename, name.c_str());
                strcpy_s(info.m_comment, sizeof info.m_comment, "");
#else    // Otherwise, use the unsafe version as fallback :(
                strncpy(info.m_filename, name.c_str(), sizeof info.m_filename);
                strncpy(info.m_comment, "", sizeof info.m_comment);
#endif

                return info;
            }

            mz_zip_archive    m_archive      = mz_zip_archive(); /**< The struct used by miniz, to handle archive files. */
            std::vector<Data> m_zipEntryData = std::vector<Data>();
            fs::path          m_archivePath;    /**< The path of the archive file. */
            bool              m_isOpen = false; /**< A flag indicating if the file is currently open for reading and writing. */
        };
    } // namespace Impl


    /**
     * @brief The ZipArchive class represents the zip archive file as a whole. It consists of the individual zip entries, which
     * can be both files and folders. It is the main access point into a .zip archive on disk and can be
     * used to create new archives and to open and modify existing archives.
     * @details
     * #### Implementation and usage details
     * Using the ZipArchive class, it is possible to create new .zip archive files, as well as open and modify existing ones.
     *
     * A ZipArchive object holds a mz_zip_archive object (a miniz struct representing a .zip archive) as well as a std::vector
     * with ZipEntry objects representing each file (entry) in the archive. The actual entry data is lazy-instantiated, so that
     * the data is only loaded from the .zip archive when it is actually needed.
     *
     * The following example shows how a new .zip archive can be created and new entries added.
     * ```cpp
     * int main() {
     *
     *       // ===== Creating empty archive
     *       KZip::ZipArchive arch;
     *       arch.Create("./TestArchive.zip");
     *
     *       // ===== Adding 10 entries to the archive
     *       for (int i = 0; i <= 9; ++i)
     *           arch.AddEntry(to_string(i) + ".txt", "this is test " + to_string(i));
     *
     *       // ===== Delete the first entry
     *       arch.DeleteEntry("0.txt");
     *
     *       // ===== Save and close the archive
     *       arch.Save();
     *       arch.Close();
     *
     *       // ===== Reopen and check contents
     *       arch.Open("./TestArchive.zip");
     *       cout << "Number of entries in archive: " << arch.GetNumEntries() << endl;
     *       cout << "Content of \"9.txt\": " << endl << arch.GetEntry("9.txt").GetDataAsString();
     *
     *       return 0;
     *   }
     * ```
     *
     * For further information, please refer to the full API documentation below.
     *
     * Note that the actual files in the .zip archive can be retrieved via the ZipEntry interface, not the ZipArchive interface.
     */
    class ZipArchive
    {
    public:
        /**
         * @brief Constructor. Constructs a null-archive, which can be used for creating a new .zip file
         * or opening an existing one.
         * @warning Before using an archive object that has been default constructed, a call to either Open() or Create() must be
         * performed. Otherwise, the object will be in a null-state and calls to member functions will be undefined.
         */
        ZipArchive() = default;

        /**
         * @brief Constructor. Constructs an archive object, using the fileName input parameter. If the file already exists,
         * it will be opened. Otherwise, a new object will be created.
         * @details
         * #### Implementation details
         * The constructors tries to open an std::ifstream. If it is valid, it means that a file already exists
         * and will be opened. Otherwise, the file does not exist and will be created.
         * @param fileName The name of the file to open or create.
         */
        explicit ZipArchive(const fs::path& fileName) : m_archivePath(fileName)
        {
            // ===== If successful, continue to open the file.
            if (fs::exists(fileName))
                open(fileName);

            // ===== If unsuccessful, create the archive file and continue.
            else
                create(fileName);
        }

        /**
         * @brief Copy Constructor (deleted).
         * @param other The object to copy.
         * @note The copy constructor has been deleted, because it is not obvious what should happen to the underlying .zip file
         * when copying the archive object. Instead, if sharing of the resource is required, a std::shared_ptr can be used.
         */
        ZipArchive(const ZipArchive& other) = delete;

        /**
         * @brief Move Constructor.
         * @param other The object to be moved.
         */
        ZipArchive(ZipArchive&& other) = default;

        /**
         * @brief Destructor.
         * @note The destructor will call the Close() member function. If the archive has been modified but not saved,
         * all changes will be discarded.
         */
        virtual ~ZipArchive()
        {
            close();
        }

        /**
         * @brief Copy Assignment Operator (deleted)
         * @param other The object to be copied.
         * @return A reference to the copied-to object.
         * @note The copy assignment operator has been deleted, because it is not obvious what should happen to the underlying
         * .zip file when copying the archive object. Instead, if sharing of the resource is required,
         * a std::shared_ptr can be used.
         */
        ZipArchive& operator=(const ZipArchive& other) = delete;

        /**
         * @brief Move Assignment Operator.
         * @param other The object to be moved.
         * @return A reference to the moved-to object.
         */
        ZipArchive& operator=(ZipArchive&& other) = default;

        /**
         * @brief Creates a new (empty) archive file with the given filename.
         * @details
         * #### Implementation details
         * A new archive file is created and initialized and thereafter closed, creating an empty archive file.
         * After ensuring that the file is valid (i.e. not somehow corrupted), it is opened using the Open() member function.
         * @param fileName The filename for the new archive.
         */
        inline void create(const fs::path& fileName)
        {
            using namespace KZip;

            // ===== Prepare an archive file;
            mz_zip_archive archive = mz_zip_archive();
            mz_zip_writer_init_file(&archive, fileName.c_str(), 0);

            // ===== Finalize and close the temporary archive
            mz_zip_writer_finalize_archive(&archive);
            mz_zip_writer_end(&archive);

            // ===== Validate the temporary file
            mz_zip_error errordata{mz_zip_error()};
            if (!mz_zip_validate_file_archive(fileName.c_str(), 0, &errordata)) {
                throw ZipRuntimeError(mz_zip_get_error_string(errordata));
            }

            // ===== If everything is OK, open the newly created archive.
            open(fileName);
        }

        /**
         * @brief Open an existing archive file with the given filename.
         * @details
         * ##### Implementation details
         * The archive file is opened and meta data for all the entries in the archive is loaded into memory.
         * @param fileName The filename of the archive to open.
         * @note If more than one entry with the same name exists in the archive, only the newest one will be loaded.
         * When saving the archive, only the loaded entries will be kept; other entries with the same name will be deleted.
         */
        inline void open(const fs::path& fileName)
        {
            // ===== Open the archive file for reading.
            if (isOpen()) close();

            m_archivePath = fileName;
            if (!mz_zip_reader_init_file(&m_archive, m_archivePath.c_str(), 0)) {
                throw ZipRuntimeError(mz_zip_get_error_string(m_archive.m_last_error));
            }
            m_isOpen = true;

            // ===== Iterate through the archive and add the entries to the internal data structure
            mz_zip_archive_file_stat info;
            for (unsigned int i = 0; i < mz_zip_reader_get_num_files(&m_archive); i++) {
                if (!mz_zip_reader_file_stat(&m_archive, i, &info)) {
                    throw ZipRuntimeError(mz_zip_get_error_string(m_archive.m_last_error));
                }

                m_zipEntries.emplace_back(Impl::ZipEntry(info));
                m_zipStats.emplace_back(info);
            }

            // ===== Remove entries with identical names. The newest entries will be retained.
            auto isEqual = [](const Impl::ZipEntry& a, const Impl::ZipEntry& b) { return a.getName() == b.getName(); };
            std::reverse(m_zipEntries.begin(), m_zipEntries.end());
            m_zipEntries.erase(std::unique(m_zipEntries.begin(), m_zipEntries.end(), isEqual), m_zipEntries.end());
            std::reverse(m_zipEntries.begin(), m_zipEntries.end());

            // ===== Add folder entries if they don't exist
            for (auto& entry : getEntryNames(false, true)) {
                if (entry.find('/') != std::string::npos) {
                    addEntry(entry.substr(0, entry.rfind('/') + 1), "");
                }
            }
        }

        /**
         * @brief Close the archive for reading and writing.
         * @note If the archive has been modified but not saved, all changes will be discarded.
         */
        inline void close()
        {
            if (isOpen()) {
                mz_zip_reader_end(&m_archive);
            }
            m_zipEntries.clear();
            m_archivePath.clear();
        }

        /**
         * @brief Checks if the archive file is open for reading and writing.
         * @return true if it is open; otherwise false;
         */
        inline bool isOpen() const
        {
            return m_isOpen;
        }

        /**
         * @brief Get a list of the entries in the archive. Depending on the input parameters, the list will include
         * directories, files or both.
         * @param includeDirs If true, the list will include directories; otherwise not. Default is true
         * @param includeFiles If true, the list will include files; otherwise not. Default is true
         * @return A std::vector of std::strings with the entry names.
         */
        inline std::vector<std::string> getEntryNames(bool includeDirs = true, bool includeFiles = true) const
        {
            if (!isOpen()) throw ZipLogicError("Cannot call GetEntryNames on empty ZipArchive object!");

            std::vector<std::string> result;

            // ===== Iterate through all the entries in the archive
            for (const auto& item : m_zipEntries) {
                // ===== If directories should be included and the current entry is a directory, add it to the result.
                if (includeDirs && item.isDirectory()) {
                    result.emplace_back(item.getName());
                    continue;
                }

                // ===== If files should be included and the current entry is a file, add it to the result.
                if (includeFiles && !item.isDirectory()) {
                    result.emplace_back(item.getName());
                    continue;
                }
            }

            return result;
        }

        /**
         * @brief Get a list of the entries in a specific directory of the archive. Depending on the input parameters,
         * the list will include directories, files or both.
         * @details This function behaves slightly different than the GetEntryNames() function. All entries in the
         * requested folder will be returned as absolute paths; however, only one level of subfolders will be returned.
         * @param dir The directory with the entries to get.
         * @param includeDirs If true, the list will include directories; otherwise not. Default is true
         * @param includeFiles If true, the list will include files; otherwise not. Default is true
         * @return A std::vector of std::strings with the entry names. The directory itself is not included.
         */
        inline std::vector<std::string> getEntryNamesInDir(const std::string& dir, bool includeDirs = true, bool includeFiles = true) const
        {
            if (!isOpen()) throw ZipLogicError("Cannot call GetEntryNamesInDir on empty ZipArchive object!");

            // ===== Get the full list of entries
            auto result = getEntryNames(includeDirs, includeFiles);

            // ===== Remove all entries not in the directory in question, as well as the root directory itself.
            if (!dir.empty()) {
                auto theDir = dir;
                if (theDir.back() != '/') {
                    theDir += '/';
                }

                result.erase(std::remove_if(result.begin(),
                                            result.end(),
                                            [&](const std::string& filename) {
                                                return filename.substr(0, theDir.size()) != theDir || filename == theDir;
                                            }),
                             result.end());
            }

            // ===== Ensure that only one level of subdirectories are included.
            result.erase(std::remove_if(result.begin(),
                                        result.end(),
                                        [&](const std::string& filename) {
                                            auto subFolderDepth = std::count(filename.begin(), filename.end(), '/');
                                            auto rootDepth =
                                                (dir.empty() ? 1
                                                             : (dir.back() == '/' ? std::count(dir.begin(), dir.end(), '/') + 1
                                                                                  : std::count(dir.begin(), dir.end(), '/') + 2));
                                            return (subFolderDepth == rootDepth && filename.back() != '/') || (subFolderDepth > rootDepth);
                                        }),
                         result.end());

            return result;
        }

        /**
         * @brief Get a list of the metadata for entries in the archive. Depending on the input parameters, the list will include
         * directories, files or both.
         * @param includeDirs If true, the list will include directories; otherwise not. Default is true
         * @param includeFiles If true, the list will include files; otherwise not. Default is true
         * @return A std::vector of ZipEntryMetaData structs with the entry metadata.
         */
        inline std::vector<ZipEntryMetaData> getMetaData(bool includeDirs = true, bool includeFiles = true)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call GetMetaData on empty ZipArchive object!");

            std::vector<ZipEntryMetaData> result;
            for (auto& item : m_zipEntries) {
                if (includeDirs && item.isDirectory()) {
                    result.emplace_back(ZipEntryMetaData(item.m_entryInfo));
                    continue;
                }

                if (includeFiles && !item.isDirectory()) {
                    result.emplace_back(ZipEntryMetaData(item.m_entryInfo));
                    continue;
                }
            }

            return result;
        }

        /**
         * @brief Get a list of the metadata for entries in a specific directory of the archive. Depending on the input
         * parameters, the list will include directories, files or both.
         * @param dir The directory with the entries to get.
         * @param includeDirs If true, the list will include directories; otherwise not. Default is true
         * @param includeFiles If true, the list will include files; otherwise not. Default is true
         * @return A std::vector of ZipEntryMetaData structs with the entry metadata. The directory itself is not included.
         */
        std::vector<ZipEntryMetaData> getMetaDataInDir(const std::string& dir, bool includeDirs = true, bool includeFiles = true)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call GetMetaDataInDir on empty ZipArchive object!");

            std::vector<ZipEntryMetaData> result;
            for (auto& item : m_zipEntries) {
                if (item.getName().substr(0, dir.size()) != dir) {
                    continue;
                }

                if (includeDirs && item.isDirectory()) {
                    result.emplace_back(ZipEntryMetaData(item.m_entryInfo));
                    continue;
                }

                if (includeFiles && !item.isDirectory()) {
                    result.emplace_back(ZipEntryMetaData(item.m_entryInfo));
                    continue;
                }
            }

            return result;
        }

        /**
         * @brief Get the number of entries in the archive. Depending on the input parameters, the number will include
         * directories, files or both.
         * @param includeDirs If true, the number will include directories; otherwise not. Default is true
         * @param includeFiles If true, the number will include files; otherwise not. Default is true
         * @return An int with the number of entries.
         */
        inline int getNumEntries(bool includeDirs = true, bool includeFiles = true) const
        {
            if (!isOpen()) throw ZipLogicError("Cannot call GetNumEntries on empty ZipArchive object!");

            return getEntryNames(includeDirs, includeFiles).size();
        }

        /**
         * @brief Get the number of entries in a specific directory of the archive. Depending on the input parameters,
         * the list will include directories, files or both.
         * @param dir The directory with the entries to get.
         * @param includeDirs If true, the number will include directories; otherwise not. Default is true.
         * @param includeFiles If true, the number will include files; otherwise not. Default is true.
         * @return An int with the number of entries. The directory itself is not included.
         */
        inline int getNumEntriesInDir(const std::string& dir, bool includeDirs = true, bool includeFiles = true) const
        {
            if (!isOpen()) throw ZipLogicError("Cannot call GetNumEntriesInDir on empty ZipArchive object!");

            return getEntryNamesInDir(dir, includeDirs, includeFiles).size();
        }

        /**
         * @brief Check if an entry with a given name exists in the archive.
         * @param entryName The name of the entry to check for.
         * @return true if it exists; otherwise false;
         */
        inline bool hasEntry(const std::string& entryName) const
        {
            if (!isOpen()) throw ZipLogicError("Cannot call HasEntry on empty ZipArchive object!");

            auto result = getEntryNames(true, true);
            return std::find(result.begin(), result.end(), entryName) != result.end();
        }

        /**
         * @brief Save the archive with a new name. The original archive will remain unchanged.
         * @param filename The new filename.
         * @note If no filename is provided, the file will be saved with the existing name, overwriting any existing data.
         * @throws ZipException A ZipException object is thrown if calls to miniz function fails.
         */
        inline void save(fs::path filename = "")
        {

            if (!isOpen()) throw ZipLogicError("Cannot call Save on empty ZipArchive object!");

            if (filename.empty()) {
                filename = m_archivePath;
            }

            // ===== Generate a random file name with the same path as the current file
            fs::path tempPath = filename.parent_path() /= Impl::GenerateRandomName(12);

            // ===== Prepare an temporary archive file with the random filename;
            mz_zip_archive tempArchive = mz_zip_archive();
            mz_zip_writer_init_file(&tempArchive, tempPath.c_str(), 0);

            // ===== Iterate through the ZipEntries and add entries to the temporary file
            for (auto& file : m_zipEntries) {
                if (file.isDirectory()) continue;
                if (!file.isModified()) {
                    if (!mz_zip_writer_add_from_zip_reader(&tempArchive, &m_archive, file.index())) {
                        throw ZipRuntimeError(mz_zip_get_error_string(m_archive.m_last_error));
                    }
                }

                else {
                    if (!mz_zip_writer_add_mem(&tempArchive,
                                               file.getName().c_str(),
                                               file.m_entryData.data(),
                                               file.m_entryData.size(),
                                               MZ_DEFAULT_COMPRESSION)) {
                        throw ZipRuntimeError(mz_zip_get_error_string(m_archive.m_last_error));
                    }
                }
            }

            // ===== Finalize and close the temporary archive
            mz_zip_writer_finalize_archive(&tempArchive);
            mz_zip_writer_end(&tempArchive);

            // ===== Validate the temporary file
            mz_zip_error errordata;
            if (!mz_zip_validate_file_archive(tempPath.c_str(), 0, &errordata)) {
                throw ZipRuntimeError(mz_zip_get_error_string(errordata));
            }

            // ===== Close the current archive, delete the file with input filename (if it exists), rename the temporary and call Open.
            close();
            nowide::remove(filename.c_str());
            nowide::rename(tempPath.c_str(), filename.c_str());
            open(filename);
        }

        /**
         * @brief
         * @param stream
         */
        inline void save(std::ostream& stream)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call Save on empty ZipArchive object!");

            // TODO: To be implemented
        }

        /**
         * @brief Deletes an entry from the archive.
         * @param name The name of the entry to delete.
         */
        inline void deleteEntry(const std::string& name)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call DeleteEntry on empty ZipArchive object!");

            m_zipEntries.erase(std::remove_if(m_zipEntries.begin(),
                                              m_zipEntries.end(),
                                              [&](const Impl::ZipEntry& entry) { return name == entry.getName(); }),
                               m_zipEntries.end());
        }

        /**
         *
         * @param path
         * @return
         */
        inline ZipEntry2 entry(const std::string& path) {

            if (!isOpen()) throw ZipLogicError("Cannot get entry from empty ZipArchive object!");

            // ===== Look up file_stat object.
            auto stats = *std::find_if(m_zipStats.begin(), m_zipStats.end(), [&](const mz_zip_archive_file_stat& info) {
                return strcmp(path.c_str(), info.m_filename) == 0;
            });

            return {&m_archive, stats};
        }

        /**
         *
         * @param index
         * @return
         */
        inline ZipEntry2 entry(int index) {
            if (!isOpen()) throw ZipLogicError("Cannot get entry from empty ZipArchive object!");

            // ===== Look up file_stat object.
            auto stats = *std::find_if(m_zipStats.begin(), m_zipStats.end(), [&](const mz_zip_archive_file_stat& info) {
                return info.m_file_index == index;
            });

            return {&m_archive, stats};
        }

        /**
         * @brief Get the entry with the specified name.
         * @param name The name of the entry in the archive.
         * @return A ZipEntry object with the requested entry.
         */
        inline ZipEntry getEntry(const std::string& name)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call GetEntry on empty ZipArchive object!");

            // ===== Look up ZipEntry object.
            auto result = std::find_if(m_zipEntries.begin(), m_zipEntries.end(), [&](const Impl::ZipEntry& entry) {
                return name == entry.getName();
            });

            // ===== If data has not been extracted from the archive (i.e., m_EntryData is empty),
            // ===== extract the data from the archive to the ZipEntry object.
            if (result->m_entryData.empty()) {
                result->m_entryData.resize(result->uncompressedSize());
                mz_zip_reader_extract_file_to_mem(&m_archive, name.c_str(), result->m_entryData.data(), result->m_entryData.size(), 0);
            }

            // ===== Check that the operation was successful
            if (!result->isDirectory() && result->m_entryData.data() == nullptr) {
                throw ZipRuntimeError(mz_zip_get_error_string(m_archive.m_last_error));
            }

            // ===== Return ZipEntry object with the file data.
            return ZipEntry(&*result);
        }

        /**
         * @brief Extract the entry with the provided name to the destination path.
         * @param name The name of the entry to extract.
         * @param dest The path to extract the entry to.
         */
        inline void extractEntry(const std::string& name, const std::string& dest)
        {
//            using namespace KZip::internal;

            if (!isOpen()) throw ZipLogicError("Cannot call ExtractEntry on empty ZipArchive object!");

            auto entry = getEntry(name);

            // ===== If the entry is a directory, create the directory as a subdirectory to dest
            if (entry.isDirectory())
                fs::create_directory(dest + entry.filename());

            // ===== If the entry is a file, stream the entry data to a file.
            else {
                std::ofstream output(dest + "/" + entry.filename(), std::ios::binary);
                auto data = entry.getData();
                output << std::string(data.begin(), data.begin());
                output.close();
            }
        }

        /**
         * @brief Extract all entries in the given directory to the destination path.
         * @param dir The name of the directory to extract.
         * @param dest The path to extract the entry to.
         * @todo To be implemented
         */
        inline void extractDir(const std::string& dir, const std::string& dest)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call ExtractDir on empty ZipArchive object!");
        }

        /**
         * @brief Extract all archive contents to the destination path.
         * @param dest The path to extract the entry to.
         * @todo To be implemented
         */
        inline void extractAll(const std::string& dest)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call ExtractAll on empty ZipArchive object!");
        }

        /**
         * @brief Add a new entry to the archive.
         * @param name The name of the entry to add.
         * @param data The ZipEntryData to add. This is an alias of a std::vector<std::byte>, holding the binary data of the entry.
         * @return The ZipEntry object that has been added to the archive.
         * @note If an entry with given name already exists, it will be overwritten.
         */
        inline ZipEntry addEntry(const std::string& name, const ZipEntryData& data)
        {
            return addEntryImpl(name, data);
        }

        /**
         * @brief Add a new entry to the archive.
         * @param name The name of the entry to add.
         * @param data The std::string data to add.
         * @return The ZipEntry object that has been added to the archive.
         * @note If an entry with given name already exists, it will be overwritten.
         */
        inline ZipEntry addEntry(const std::string& name, const std::string& data)
        {
            return addEntryImpl(name, ZipEntryData{data.begin(), data.end()});
        }

        /**
         * @brief Add a new entry to the archive, using another entry as input. This is mostly used for cloning existing entries
         * in the same archive, or for copying entries from one archive to another.
         * @param name The name of the entry to add.
         * @param entry The ZipEntry to add.
         * @return The ZipEntry object that has been added to the archive.
         * @note If an entry with given name already exists, it will be overwritten.
         */
        inline ZipEntry addEntry(const std::string& name, const ZipEntry& entry)
        {
            return addEntryImpl(name, entry.getData());
        }

    private:
        /**
         * @brief Add a new entry to the archive.
         * @param name The name of the entry to add.
         * @param data The ZipEntryData to add. This is an alias of a std::vector<std::byte>, holding the binary data of the entry.
         * @return The ZipEntry object that has been added to the archive.
         * @note If an entry with given name already exists, it will be overwritten.
         */
        inline ZipEntry addEntryImpl(const std::string& name, const ZipEntryData& data)
        {
            if (!isOpen()) throw ZipLogicError("Cannot call AddEntry on an empty ZipArchive object!");

            // ===== Ensure that all folders and subfolders have an entry in the archive
            auto folders = getEntryNames(true, false);
            auto position = uint64_t {0};
            while (name.find('/', position) != std::string::npos) {
                position    = name.find('/', position) + 1;
                auto folder = name.substr(0, position);

                // ===== If folder isn't registered in the archive, add it.
                if (std::find(folders.begin(), folders.end(), folder) == folders.end()) {
                    m_zipEntries.emplace_back(Impl::ZipEntry(folder, ""));
                    folders.emplace_back(folder);
                }
            }

            // ===== Check if an entry with the given name already exists in the archive.
            auto result = std::find_if(m_zipEntries.begin(), m_zipEntries.end(), [&](const Impl::ZipEntry& entry) {
                return name == entry.getName();
            });

            // ===== If the entry exists, replace the existing data with the new data, and return the ZipEntry object.
            if (result != m_zipEntries.end()) {
                result->setData(data);
                return ZipEntry(&*result);
            }

            // ===== Finally, add a new entry with the given name and data, and return the object.
            return ZipEntry(&m_zipEntries.emplace_back(Impl::ZipEntry(name, data)));
        }

    private: // NOLINT
        mz_zip_archive m_archive = mz_zip_archive(); /**< The struct used by miniz, to handle archive files. */
        fs::path                        m_archivePath;        /**< The path of the archive file. */
        bool                            m_isOpen     = false; /**< A flag indicating if the file is currently open for reading and writing. */
        std::vector<Impl::ZipEntry>     m_zipEntries = std::vector<Impl::ZipEntry>(); /**< Data structure for all entries in the archive. */
        std::vector<mz_zip_archive_file_stat> m_zipStats = std::vector<mz_zip_archive_file_stat>();
    };
}    // namespace KZip

#pragma warning(pop)
#endif    // ZIPPYHEADERONLY_ZIPPY_HPP
