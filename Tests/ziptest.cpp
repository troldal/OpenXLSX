/*
 * Program:  ziptest
 * Compile instructions: g++ ziptest.cpp -lzip -o ziptest
 * Part of: OpenXLSX
 *
 * OpenXLSX Copyright (c) 2018, Kenneth Troldal Balslev
 * ziptest Copyright 2025 Lars Uffmann <coding23@uffmann.name>
 *
 * --- NOTE: The below license is not applicable to namespace OpenXLSX! ---
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/*!
 * @program ziptest
 * @depends libzip
 * @version 2025-04-26 21:30 CET
 * @brief   provide a wrapper API for libzip to interface with OpenXLSX XLZipArchive
 * @disclaimer This module is in a draft stage, to be reviewed / tested
 */
#include <cstdio>       // fprintf
#include <iostream>     // std::cout / std::cerr
#include <memory>       // std::shared_ptr
#include <string>       // std::string
#include <string.h>     // stderror
#include <sys/stat.h>   // stat, struct stat

#include <zip.h>        // libzip API


#ifdef ZIPPY2_LIBRARY_H
	#define ZIPPY_IN_USE
	#define XLZipImplementation Zippy::ZipArchive
#else
	#undef ZIPPY_IN_USE
	#define XLZipImplementation LibZip::ZipArchive
#endif

namespace LibZip {
	static int loadArchiveData(void **datap, size_t *sizep, const char *archive)
	{
		/* example implementation that reads data from file */
		struct stat st;
		FILE *fp;

		if (stat(archive, &st) < 0) {
			if (errno != ENOENT) {
				fprintf(stderr, "can't stat %s: %s\n", archive, strerror(errno));
				return -1;
			}

			*datap = nullptr;
			*sizep = 0;
		
			return 0;
		}

		if ((*datap = malloc(static_cast<size_t>(st.st_size))) == nullptr) {
			fprintf(stderr, "can't allocate buffer\n");
			return -1;
		}

		if ((fp=fopen(archive, "r")) == nullptr) {
			free(*datap);
			fprintf(stderr, "can't open %s: %s\n", archive, strerror(errno));
			return -1;
		}

		if (fread(*datap, 1, (size_t)st.st_size, fp) < (size_t)st.st_size) {
			free(*datap);
			fprintf(stderr, "can't read %s: %s\n", archive, strerror(errno));
			fclose(fp);
			return -1;
		}

		fclose(fp);

		*sizep = static_cast<size_t>(st.st_size);
		return 0;
	}

	static int saveArchiveFile(void *data, size_t size, const char *archive)
	{
		/* example implementation that writes data to file */
		FILE *fp;

		if (data == nullptr) {
			if (remove(archive) < 0 && errno != ENOENT) {
				fprintf(stderr, "can't remove %s: %s\n", archive, strerror(errno));
				return -1;
			}
			return 0;
		}

		if ((fp = fopen(archive, "wb")) == nullptr) {
			fprintf(stderr, "can't open %s: %s\n", archive, strerror(errno));
			return -1;
		}
		if (fwrite(data, 1, size, fp) < size) {
			fprintf(stderr, "can't write %s: %s\n", archive, strerror(errno));
			fclose(fp);
			return -1;
		}
		if (fclose(fp) < 0) {
			fprintf(stderr, "can't write %s: %s\n", archive, strerror(errno));
			return -1;
		}

		return 0;
	}


	class ZipArchive {
	private:
		void           *m_zipData;    // the raw data of the unmodified source archive, (re-)set on ZipArchive::Open
		zip_error_t     m_zipError;   // TBD how useful: save zip_error_t states across different methods
		zip_source_t   *m_zipSrc;     // the zip source, as opposed to the archive - TBD what the logic of this is
		zip_t          *m_za;         // the zip archive - changed must be commited before they can be read via GetEntryDataAsString
		std::string     m_name;       // name of the archive file for which ZipArchive::Open was called

		void closeZip() {
			// close archive and commit changes
			if (zip_close(m_za) < 0) {
				std::cerr << "can't close zip archive '" << m_name << "': " << zip_strerror(m_za) << std::endl;
				throw 1;
			}
			m_za = nullptr;
		}
		void reopenFromZipSource() {
			/* (re)open zip archive from source */
			if ((m_za = zip_open_from_source(m_zipSrc, 0, &m_zipError)) == nullptr) {
				std::cerr << "can't open zip from source: " << zip_error_strerror(&m_zipError) << std::endl;
				zip_source_free(m_zipSrc);
				m_zipSrc = nullptr; // prevent double-free
				zip_error_fini(&m_zipError);
				throw 1;
			}
		}

	public:
		ZipArchive()
		: m_zipData(nullptr), m_zipSrc(nullptr), m_za(nullptr), m_name("") {}
		~ZipArchive() {
			if (IsOpen()) Close();
		}
		bool IsOpen() const { return m_za != nullptr; }
		void Close() {
			if (!IsOpen())
				throw 1;

			closeZip();    // invalidates m_zipSrc and m_zipData

			m_zipSrc = nullptr;
			m_zipData = nullptr;
		}

		void Open(std::string archiveName) {
			if (IsOpen()) throw 1; // user must close the archive before re-opening

			/* get buffer with zip archive inside */
			size_t size;
			if (loadArchiveData(&m_zipData, &size, archiveName.c_str()) < 0)
				throw 1;

			zip_error_init(&m_zipError);
			/* create source from buffer */
			if ((m_zipSrc = zip_source_buffer_create(m_zipData, size, 1, &m_zipError)) == nullptr) {
				std::cerr << "can't create source: " << zip_error_strerror(&m_zipError) << std::endl;
				free(m_zipData);
				m_zipData = nullptr;   // prevent double-free
				zip_error_fini(&m_zipError);
				throw 1;
			}

			reopenFromZipSource();

			zip_error_fini(&m_zipError);
			m_name = archiveName;
		}

		void CommitChanges() { // make changes readable
			zip_source_keep(m_zipSrc); // ensure that m_zipSrc survives zip_close
			closeZip();                // close zip and thereby commit changes
			reopenFromZipSource();     // ensure that archive is valid again for further reading & modifications
		}

		void Save(std::string savePath) {
			void *saveData = nullptr;
			if( savePath == "") {
				// saving in original file location!
				savePath = m_name;
			}

			zip_source_keep(m_zipSrc); // ensure that m_zipSrc survives zip_close
			closeZip();                // close the zip archive to commit changes

			if (zip_source_is_deleted(m_zipSrc)) {
				/* new archive is empty, thus no data */
				throw 1; // should not happen
			}
			else {
				zip_stat_t zst;
				if (zip_source_stat(m_zipSrc, &zst) < 0) {
					std::cerr << "can't stat source: " << zip_error_strerror(zip_source_error(m_zipSrc)) << std::endl;
					throw 1;
				}

				size_t size = zst.size;
				if (zip_source_open(m_zipSrc) < 0) {               // prepare m_zipSrc for reading archive data from it
					std::cerr << "can't open source: " << zip_error_strerror(zip_source_error(m_zipSrc)) << std::endl;
					throw 1;
				}
				if ((saveData = malloc(size)) == nullptr) {        // allocate memory to store the modified archive for saving to a file
					std::cerr << "malloc failed: " << strerror(errno) << std::endl;
					zip_source_close(m_zipSrc);
					throw 1;
				}
				if ((zip_uint64_t)zip_source_read(m_zipSrc, saveData, size) < size) {   // read the archive raw data from m_zipSrc
					std::cerr << "can't read saveData from source: " << zip_error_strerror(zip_source_error(m_zipSrc)) << std::endl;
					zip_source_close(m_zipSrc);
					free(saveData);
					throw 1;
				}
				zip_source_close(m_zipSrc);                                    // close the zip source
				if( saveArchiveFile(saveData, size, savePath.c_str()) < 0 )    // and save the archive
					throw 1;

				free(saveData);         // free temporary buffer used for saving the archive
				reopenFromZipSource();  // ensure that archive is valid again for further modifications
			}
		}
		
		void AddEntry(const std::string& entryName, const std::string& data)
		{
			// create the zip source for a single entry from data
			struct zip_source *zipSrc = zip_source_buffer(m_za, data.c_str(), data.length(), 0);
			if (!zipSrc) throw 1;

			if( -1 == zip_file_add(m_za, entryName.c_str(), zipSrc, ZIP_FL_OVERWRITE | ZIP_FL_ENC_UTF_8) ) throw 1;
			return;
		}

		void DeleteEntry(const std::string& entryName)
		{
			int index = zip_name_locate(m_za, entryName.c_str(), 0);
			if (index == -1) throw 1;
			if( 0 != zip_delete(m_za, index) ) throw 1;
		}

		std::string GetEntryDataAsString(const std::string& entryName) const
		{
			int index = zip_name_locate(m_za, entryName.c_str(), 0);   // ensure that entryName exists in the archive
			if (index == -1) throw 1;

			// determine how many bytes entry data encompasses
			struct zip_stat entryInfo;
			if (0 != zip_stat_index(m_za, index, 0, &entryInfo))
				throw 1;

			// std::cout << "stats for file " << entryName << std::endl;
			// std::cout << "  name:  " << entryInfo.name << std::endl;
			// std::cout << "  index: " << entryInfo.index << std::endl;
			// std::cout << "  crc:   " << entryInfo.crc << std::endl;
			// std::cout << "  size:  " << entryInfo.size << std::endl;
			// std::cout << "  mtime: " << entryInfo.mtime << std::endl;
			// std::cout << "  comp_size:         " << entryInfo.comp_size << std::endl;
			// std::cout << "  comp_method:       " << entryInfo.comp_method << std::endl;
			// std::cout << "  encryption_method: " << entryInfo.encryption_method << std::endl;

			zip_file_t *fd = zip_fopen_index(m_za, index, 0);                            // get a file descriptor for desired entry
			if (fd == nullptr) {
				std::cerr << "an error occurred trying to open archive file " << entryName << ": " << zip_error_strerror(zip_get_error(m_za)) << std::endl;
				throw 1;
			}

			char *entryDataChr = reinterpret_cast<char *>(malloc(entryInfo.size + 1)); // temporary buffer to construct return string from
			zip_fread(fd, entryDataChr, entryInfo.size);                               // read entry data into temporary buffer
			entryDataChr[entryInfo.size] = 0;                                          // ensure zero-termination of buffer data
			std::string entryDataAsString(entryDataChr);                               // construct string return value from buffer data
			free(entryDataChr);                                                        // free temporary buffer
			zip_fclose(fd);                                                            // close file descriptor

			return entryDataAsString;
		}

		bool HasEntry(const std::string& entryName) const
		{
			return (-1 != zip_name_locate(m_za, entryName.c_str(), 0));
			
		}

		ssize_t EntryCount() const { return zip_get_num_files(m_za); }   // get valid entry index range [0;EntryCount()-1]

		std::string EntryName( int index ) const                       // get name of the archive entry at index
		{
			const char *name = zip_get_name(m_za, index, 0);
			return std::string(name == nullptr ? "" : name);
		}
	};

} // namespace LibZip

namespace OpenXLSX
{
#define OPENXLSX_EXPORT
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLZipArchive
    {
    public:
        /**
         * @brief
         */
        XLZipArchive() : m_archive(nullptr) {}

        /**
         * @brief
         * @param other
         */
        XLZipArchive(const XLZipArchive& other) = default;

        /**
         * @brief
         * @param other
         */
        XLZipArchive(XLZipArchive&& other) = default;

        /**
         * @brief
         */
        ~XLZipArchive() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLZipArchive& operator=(const XLZipArchive& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLZipArchive& operator=(XLZipArchive&& other) = default;

        /**
         * @brief check if zippy is in use (or ziplib)
         * @param
         * @return true if zippy is used for the zip implementation, otherwise false
         */
        bool isZippy() {
        #ifdef ZIPPY_IN_USE
            return true;
        #else
            return false;
        #endif
        }

        /**
         * @brief
         * @return
         */
        explicit operator bool() const {
            return isValid();
        }

        bool isValid() const { return m_archive != nullptr; }

        /**
         * @brief
         * @return
         */
        bool isOpen() const {
            return m_archive && m_archive->IsOpen();
        }

        /**
         * @brief
         * @param fileName
         */
        void open(const std::string& fileName)
        {
            if( isOpen() ) throw 1; // prevent double-open
            m_archive = std::make_shared<XLZipImplementation>();
            try {
                m_archive->Open(fileName);
            }
            catch( ... ) {    // catch all exceptions
                m_archive.reset();    // make m_archive invalid again
                throw;                // re-throw
            }
        }

        /**
         * @brief
         */
        void close()
        {
            if (!m_archive) throw 1; // prevent SEGFAULT
            m_archive->Close();
            m_archive.reset();
        }

        /**
         * @brief make archive updates (from addEntry) available to calls via getEntry
         */
        void commitChanges()
        {
            m_archive->CommitChanges();
        }

        /**
         * @brief
         * @param path
         */
        void save(const std::string& path = "")
        {
            m_archive->Save(path);
        }

        /**
         * @brief
         * @param name
         * @param data
         */
        void addEntry(const std::string& name, const std::string& data)
        {
            m_archive->AddEntry(name, data);
        }

        /**
         * @brief
         * @param entryName
         */
        void deleteEntry(const std::string& entryName)
        {
            m_archive->DeleteEntry(entryName);
        }

        /**
         * @brief
         * @param name
         * @return
         */
        std::string getEntry(const std::string& name) const
        {
        #ifdef ZIPPY_IN_USE
            return m_archive->GetEntry(name).GetDataAsString();
        #else
            return m_archive->GetEntryDataAsString(name);
        #endif
        }

        /**
         * @brief
         * @param entryName
         * @return
         */
        bool hasEntry(const std::string& entryName) const
        {
            return m_archive->HasEntry(entryName);
        }

        ssize_t entryCount() const { return m_archive->EntryCount(); }
        std::string entryName( int index ) const { return m_archive->EntryName(index); }

    private:
        std::shared_ptr<XLZipImplementation> m_archive; /**< */
    };
}    // namespace OpenXLSX

int main(int argc, char *argv[])
{
	std::string archiveName = "Demo01";
	std::string archiveExt = ".xlsx";

	OpenXLSX::XLZipArchive myArc{};

	std::cout << "myArc.isZippy() is " << ( myArc.isZippy() ? "true" : "false" ) << std::endl;
	std::cout << "myArc.isOpen() is " << ( myArc.isOpen() ? "true" : "false" ) << std::endl;
	myArc.open(archiveName + archiveExt);
	std::cout << "myArc.isOpen() is " << ( myArc.isOpen() ? "true" : "false" ) << std::endl;
	std::string entryName = "docProps/core.xml";
	std::cout << "myArc.hasEntry(\"" << entryName << "\") is " << ( myArc.hasEntry(entryName) ? "true" : "false" ) << std::endl;

	int i = 0;
	ssize_t entryCount = myArc.entryCount();
	std::cout << "myArc.entryCount() is " << entryCount << std::endl;
	while (i < entryCount) {
		std::cout << "myArc.entryName(" << i << ") is " << myArc.entryName(i) << std::endl;
		++i;
	}

	std::string testEntry = "xl/worksheets/sheet2.xml";
	std::cout << "adding a file to myArc: " << testEntry << std::endl;
	myArc.addEntry(testEntry, "Hello World!");

	myArc.commitChanges();

	// testEntry = "xl/worksheets/sheet1.xml";
	std::string test = myArc.getEntry(testEntry);
	std::cout << "contents of " << testEntry << ":" << std::endl;
	std::cout << "------------------------------------------------" << std::endl;
	std::cout << test << std::endl;
	std::cout << "------------------------------------------------" << std::endl;


	myArc.save(archiveName + "_1" + archiveExt); // saves and re-opens from source
	myArc.close();
	std::cout << "myArc.isOpen() is " << ( myArc.isOpen() ? "true" : "false" ) << std::endl;

	return 0;
}
