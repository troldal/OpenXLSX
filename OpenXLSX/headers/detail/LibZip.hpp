#ifndef LIBZIP_WRAPPER_H
#define LIBZIP_WRAPPER_H

#include <cstdio>       // fprintf
#include <filesystem>   // std::filesystem::remove, std::filesystem::rename
#include <random>       // std::random_device, std::mt19937, std::uniform_int_distribution
#include <string>       // std::string
#include <string.h>     // strerror
#include <sys/stat.h>   // stat, struct stat

#include <zip.h>        // libzip

namespace LibZip {
    /**
     * @brief The LibZipException class is a custom exception class derived from the std::runtime_error class.
     * @details In case of an error in the LibZip library, an LibZipException object will be thrown, with a message
     * describing the details of the error.
     */
    class LibZipException : public std::runtime_error
    {
    public:
        explicit LibZipException(const std::string& err) : runtime_error(err) {};
    };

    /**
     * @brief
     */
    class LibZipInputError : public LibZipException
    {
    public:
        explicit LibZipInputError(const std::string& err) : LibZipException(err) {};
    };

    /**
     * @brief
     */
    class LibZipInternalError : public LibZipException
    {
    public:
        explicit LibZipInternalError(const std::string& err) : LibZipException(err) {};
    };


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

        std::string result;
        for (int i = 0; i < length; ++i) {
            result += letters[distr(generator)];
        }

        return result + ".tmp";
    }

    class ZipArchive {
    private:
        void           *m_zipData;    // the raw data of the unmodified source archive, (re-)set on ZipArchive::Open
        size_t          m_zipSize;    // the size in bytes of the unmodified source archive stored at m_zipData
        zip_error_t     m_zipError;   // TBD how useful: save zip_error_t states across different methods
        zip_source_t   *m_zipSrc;     // the zip source, as opposed to the archive - TBD what the logic of this is
        zip_t          *m_za;         // the zip archive - changed must be commited before they can be read via GetEntryDataAsString
        std::string     m_name;       // name of the archive file for which ZipArchive::Open was called

        /**
         * @brief close archive, commit changes and - if the underlying reference count reaches zero - free resources (m_zipData, m_zipSrc)
         * @return N/A
         * @throw LibZipInternalError
         */
        void closeZip() {
            // 
            if (zip_close(m_za) < 0) {
                using namespace std::literals::string_literals;
                throw LibZipInternalError("ZipArchive::closeZip: can't close zip archive '"s + m_name + "': "s + zip_strerror(m_za));
            }
            m_za = nullptr;
        }

        /**
         * @brief (re)open zip archive from source, allowing further read & write access
         * @return N/A
         * @throw LibZipInternalError
         */
        void reopenFromZipSource() {
            if ((m_za = zip_open_from_source(m_zipSrc, 0, &m_zipError)) == nullptr) {
                zip_source_free(m_zipSrc);
                m_zipSrc = nullptr; // prevent double-free
                zip_error_fini(&m_zipError);
                using namespace std::literals::string_literals;
                throw LibZipInternalError("ZipArchive::reopenFromZipSource: can't open zip from source: "s + zip_error_strerror(&m_zipError));
            }
        }

        /**
         * @brief load data from filename into this->m_zipData
         * @param filename load a file from here
         * @return 0 on success, -1 on failure
         */
        int loadArchiveData(std::string filename)
        {
            m_zipSize = 0; // init in case of failure

            struct stat st;
            if (stat(filename.c_str(), &st) < 0) {
                if (errno != ENOENT) {
                    fprintf(stderr, "ZipArchive::loadArchiveData: can't obtain info about %s: %s\n", filename.c_str(), strerror(errno));
                    return -1;
                }

                m_zipData = nullptr;
                return -1;
            }

            if ((m_zipData = malloc(static_cast<size_t>(st.st_size))) == nullptr) {
                fprintf(stderr, "ZipArchive::loadArchiveData: can't allocate buffer\n");
                return -1;
            }

            FILE *fp = fopen(filename.c_str(), "r");
            if (fp == nullptr) {
                free(m_zipData);
                m_zipData = nullptr;
                fprintf(stderr, "ZipArchive::loadArchiveData: can't open %s: %s\n", filename.c_str(), strerror(errno));
                return -1;
            }

            if (fread(m_zipData, 1, static_cast<size_t>(st.st_size), fp) < static_cast<size_t>(st.st_size)) {
                free(m_zipData);
                m_zipData = nullptr;
                fprintf(stderr, "ZipArchive::loadArchiveData: can't read %s: %s\n", filename.c_str(), strerror(errno));
                fclose(fp);
                return -1;
            }

            fclose(fp);

            m_zipSize = static_cast<size_t>(st.st_size);
            return 0;
        }

        /**
         * @brief Save size bytes from data to a temporary file, then move that to filename
         * @param data save data from here
         * @param size amount of bytes to save
         * @param filename (final) save destination
         * @return 0 on success, -1 on failure
         */
        int saveArchiveFile(void *data, size_t size, std::string filename)
        {
        #ifdef _WIN32
            std::replace( filename.begin(), filename.end(), '\\', '/' ); // pull request #210, alternate fix: fopen etc work fine with forward slashes
        #endif

            // ===== Determine path of the current file
            size_t pathPos = filename.rfind('/');

            // pull request #191, support AmigaOS style paths
        #ifdef __amigaos__
            constexpr const char * localFolder = "";    // local folder on AmigaOS can not be explicitly expressed in a path
            if (pathPos == std::string::npos) pathPos = filename.rfind(':'); // if no '/' found, attempt to find amiga drive root path
        #else
            constexpr const char * localFolder = "./"; // local folder on _WIN32 && __linux__ is .
        #endif
            std::string tempPath{};
            if (pathPos != std::string::npos) tempPath = filename.substr(0, pathPos + 1);
            else tempPath = localFolder; // prepend explicit identification of local folder in case path did not contain a folder

            // ===== Generate a random file name with the same path as the current file
            tempPath = tempPath + GenerateRandomName(20);

            FILE *fp = fopen(tempPath.c_str(), "wb");
            if (fp  == nullptr) {
                fprintf(stderr, "ZipArchive::saveArchiveFile: can't open %s: %s\n", tempPath.c_str(), strerror(errno));
                return -1;
            }
            if (fwrite(data, 1, size, fp) < size) {
                fprintf(stderr, "ZipArchive::saveArchiveFile: can't write %s: %s\n", tempPath.c_str(), strerror(errno));
                fclose(fp);
                return -1;
            }
            if (fclose(fp) < 0) {
                fprintf(stderr, "ZipArchive::saveArchiveFile: can't close %s: %s\n", tempPath.c_str(), strerror(errno));
                return -1;
            }
            std::filesystem::remove(filename.c_str());
            std::filesystem::rename(tempPath.c_str(), filename.c_str());

            return 0;
        }


    public:
        /**
         * @brief construct an empty ZipArchive
         */
        ZipArchive()
        : m_zipData(nullptr), m_zipSize(0), m_zipSrc(nullptr), m_za(nullptr), m_name("") {}

        /**
         * @brief destructor: close zip archive if open
         * @note data will not be saved automatically on destruction - user must call ZipArchive::Save
         */
        ~ZipArchive() {
            if (IsOpen()) Close();
        }

        /**
         * @brief test whether zip archive is open / ready for reads & writes
         * @return true if a zip file has been opened, false if archive is not open
         */
        bool IsOpen() const { return m_za != nullptr; }

        /**
         * @brief close an open archive and free resources
         * @return N/A
         * @throw LibZipInputError if archive is not open
         */
        void Close() {
            if (!IsOpen())
                throw LibZipInputError("ZipArchive::Close: archive is not open!");

// struct zip_file_attributes {
//     zip_uint64_t valid;                     /* which fields have valid values */
//     zip_uint8_t version;                    /* version of this struct, currently 1 */
//     zip_uint8_t host_system;                /* host system on which file was created */
//     zip_uint8_t ascii;                      /* flag whether file is ASCII text */
//     zip_uint8_t version_needed;             /* minimum version needed to extract file */
//     zip_uint32_t external_file_attributes;  /* external file attributes (host-system specific) */
//     zip_uint16_t general_purpose_bit_flags; /* general purpose big flags, only some bits are honored */
//     zip_uint16_t general_purpose_bit_mask;  /* which bits in general_purpose_bit_flags are valid */
// };
// zip_file_attributes_t zipAttrs;
// zipAttrs.general_purpose_bit_flags = 0;
// zip_source_get_file_attributes(m_zipSrc, &zipAttrs);
// printf( "zip_file_attributes::general_purpose_bit_flags is 0x%08x\n", static_cast< uint16_t >( zipAttrs.general_purpose_bit_flags ) );

            closeZip();    // invalidates m_zipData and m_zipSrc because reference count for the underlying zip source should reach zero
            m_zipData = nullptr;
            m_zipSize = 0;
            m_zipSrc = nullptr;
            m_name = "";
        }

        /**
         * @brief Open an existing archive file
         * @param archiveName the file to load from
         * @return N/A
         * @throw LibZipInputError if archive is already open
         * @throw LibZipInternalError if archiveName can't be opened / loaded
         */
        void Open(std::string archiveName) {
            using namespace std::literals::string_literals;
            if (IsOpen())
                throw LibZipInputError("ZipArchive::Open: archive is already open for source file "s + m_name); // user must close the archive before re-opening

            /* get buffer with zip archive inside */
            if (loadArchiveData(archiveName) < 0)
                throw LibZipInternalError("ZipArchive::Open: failed to load archive data from file "s + archiveName);

            zip_error_init(&m_zipError);
            /* create source from buffer */
            if ((m_zipSrc = zip_source_buffer_create(m_zipData, m_zipSize, 1, &m_zipError)) == nullptr) {
                free(m_zipData);
                m_zipData = nullptr;   // prevent double-free
                zip_error_fini(&m_zipError);
                throw LibZipInternalError("ZipArchive::Open: can't create source: "s + zip_error_strerror(&m_zipError));
            }

            reopenFromZipSource();

            zip_error_fini(&m_zipError);
            m_name = archiveName;
        }

        /**
         * @brief make modified archive data available for reading via GetEntryDataAsString
         * @return N/A
         * @throw LibZipInternalError thrown from closeZip or reopenFromZipSource upon failure
         */
        void CommitChanges() { // make changes readable
            zip_source_keep(m_zipSrc); // ensure that m_zipSrc survives zip_close
            closeZip();                // close zip and thereby commit changes
            reopenFromZipSource();     // ensure that archive is valid again for further reading & modifications
        }

        /**
         * @brief Save the (modified) archive to savePath
         * @param savePath save here, or save to original filename if empty string is provided
         * @return N/A
         * @throw LibZipInternalError upon any failure
         */
        void Save(std::string savePath)
        {
            if (savePath == "") // saving in original file location!
                savePath = m_name;

            zip_source_keep(m_zipSrc); // ensure that m_zipSrc survives zip_close
            closeZip();                // close the zip archive to commit changes

            if (zip_source_is_deleted(m_zipSrc)) {
                /* new archive is empty, thus no data */
                throw LibZipInternalError("ZipArchive::Save: zip_source_is_deleted returned true, this should not happen");
            }
            else {
                using namespace std::literals::string_literals;

                void *saveData = nullptr;
                zip_stat_t zst;
                if (zip_source_stat(m_zipSrc, &zst) < 0)
                    throw LibZipInternalError("ZipArchive::Save: can't stat source: "s + zip_error_strerror(zip_source_error(m_zipSrc)));

                size_t saveSize = zst.size;
                if (zip_source_open(m_zipSrc) < 0)                    // prepare m_zipSrc for reading archive data from it
                    throw LibZipInternalError("ZipArchive::Save: can't open source: "s + zip_error_strerror(zip_source_error(m_zipSrc)));

                if ((saveData = malloc(saveSize)) == nullptr) {       // allocate memory to store the modified archive for saving to a file
                    zip_source_close(m_zipSrc);
                    throw LibZipInternalError("ZipArchive::Save: malloc failed: "s + strerror(errno));
                }
                // ==== Read the archive raw data from m_zipSrc
                if (static_cast<size_t>(zip_source_read(m_zipSrc, saveData, saveSize)) < saveSize) {
                    zip_source_close(m_zipSrc);
                    free(saveData);
                    throw LibZipInternalError("ZipArchive::Save: can't read saveData from source: "s + zip_error_strerror(zip_source_error(m_zipSrc)));
                }
                zip_source_close(m_zipSrc);                           // close the zip source
                if (saveArchiveFile(saveData, saveSize, savePath.c_str()) < 0) // and save the archive
                    throw LibZipInternalError("ZipArchive::Save: failed to save archive data to "s + savePath);

                free(saveData);         // free temporary buffer used for saving the archive
            }

            reopenFromZipSource();  // ensure that archive is valid again for further modifications
        }

        /**
         * @brief Add a file to the archive and write data to it - overwrite existing files
         * @param entryName file path in the archive - if existing, destination will be replaced
         * @param data contents for a text file to be placed in the archive for entryName
         * @return N/A
         * @throw LibZipInputError when called on a non-valid archive
         * @throw LibZipInternalError upon any other failure
         */
        void AddEntry(const std::string& entryName, const std::string& data)
        {
            if (!IsOpen())
                throw LibZipInputError("ZipArchive::AddEntry: archive is not open!");

            void *buf = malloc(data.size());       // allocate memory for a persistent copy of data.data()
            memcpy(buf, data.data(), data.size()); // copy data.data() to the allocated memory

            // create the zip source for a single entry from data
            struct zip_source *zipSrc = zip_source_buffer(m_za, buf, data.size(), 1);  // create zip source from buffer, flagging that buf shall be freed by libzip

            if (!zipSrc)
                throw LibZipInternalError("ZipArchive::AddEntry: failed to create a zip source from entry data");

            if (-1 == zip_file_add(m_za, entryName.c_str(), zipSrc, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8)) {
                using namespace std::literals::string_literals;
                throw LibZipInternalError("ZipArchive::AddEntry: failed to add file to archive: "s +  zip_error_strerror(zip_get_error(m_za)));
            }
        }

        /**
         * @brief Delete a file from the archive
         * @param entryName archive file to delete
         * @return N/A
         * @throw LibZipInputError when called on a non-valid archive or if entryName does not exist in the archive
         * @throw LibZipInternalError upon any other failure
         */
        void DeleteEntry(const std::string& entryName)
        {
            using namespace std::literals::string_literals;

            if (!IsOpen())
                throw LibZipInputError("ZipArchive::DeleteEntry: archive is not open!");

            int index = zip_name_locate(m_za, entryName.c_str(), 0);
            if (index == -1)
                throw LibZipInputError("ZipArchive::DeleteEntry: archive does not contain file: "s + entryName);

            if (0 != zip_delete(m_za, index))
                throw LibZipInternalError("ZipArchive::DeleteEntry: failed to delete file "s + entryName + " from archive: "s + zip_error_strerror(zip_get_error(m_za)));
        }

        /**
         * @brief Return the contents of an archive file as a string
         * @note function is intended for text files only
         * @param entryName archive file to fetch
         * @return std::string with the contents of the archive file
         * @throw LibZipInputError when called on a non-valid archive or if entryName does not exist in the archive
         * @throw LibZipInternalError upon any other failure
         */
        std::string GetEntryDataAsString(const std::string& entryName) const
        {
            using namespace std::literals::string_literals;

            if (!IsOpen())
                throw LibZipInputError("ZipArchive::GetEntryDataAsString: archive is not open!");

            int index = zip_name_locate(m_za, entryName.c_str(), 0);   // ensure that entryName exists in the archive
            if (index == -1)
                throw LibZipInputError("ZipArchive::GetEntryDataAsString: archive does not contain file: "s + entryName);

            // determine how many bytes entry data encompasses
            struct zip_stat entryInfo;
            if (0 != zip_stat_index(m_za, index, 0, &entryInfo))
                throw LibZipInternalError("ZipArchive::GetEntryDataAsString: failed to obtain archive entry info for "s + entryName + ": "s + zip_error_strerror(zip_get_error(m_za)));

            zip_file_t *fd = zip_fopen_index(m_za, index, 0);                            // get a file descriptor for desired entry
            if (fd == nullptr)
                throw LibZipInternalError("ZipArchive::GetEntryDataAsString: an error occurred trying to open archive file "s + entryName + ": "s + zip_error_strerror(zip_get_error(m_za)));

            char *entryDataChr = reinterpret_cast<char *>(malloc(entryInfo.size + 1)); // temporary buffer to construct return string from
            zip_fread(fd, entryDataChr, entryInfo.size);                               // read entry data into temporary buffer
            entryDataChr[entryInfo.size] = 0;                                          // ensure zero-termination of buffer data
            std::string entryDataAsString(entryDataChr);                               // construct string return value from buffer data
            free(entryDataChr);                                                        // free temporary buffer
            zip_fclose(fd);                                                            // close file descriptor

            return entryDataAsString;
        }

        /**
         * @brief test whether a file exists in the archive
         * @param entryName archive file to locate
         * @return true if archive has file, false otherwise
         */
        bool HasEntry(const std::string& entryName) const
        {
            return (-1 != zip_name_locate(m_za, entryName.c_str(), 0)); // zip_name_locate also returns -1 if m_za is nullptr, thus no extra check for IsOpen()
        }

        /**
         * @brief get valid entry index range [0;EntryCount()-1]
         * @return amount of entries in the zip archive
         * @return -1 on error (m_za is not open / nullptr)
         */
        ssize_t EntryCount() const { return zip_get_num_files(m_za); }

        /**
         * @brief get entry name by index
         * @param index which entry name to get, must be within [0;EntryCount()-1]
         * @return std::string with the name of the entry at index
         * @return empty string if entry does not have a name / index is invalid
         * @throw LibZipInputError if archive is closed (nullptr)
         */
        std::string EntryName( int index ) const                       // get name of the archive entry at index
        {
            if (!IsOpen())
                throw LibZipInputError("ZipArchive::EntryName: archive is not open");
            const char *name = zip_get_name(m_za, index, 0);
            return std::string(name == nullptr ? "" : name);
        }
    };
} // namespace LibZip

#endif // LIBZIP_WRAPPER_H
