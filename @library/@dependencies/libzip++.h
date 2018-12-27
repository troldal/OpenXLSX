
#ifndef LIBZIPPP_H
#define	LIBZIPPP_H

/*
  libzippp.h -- exported declarations.
  Copyright (C) 2013 Cédric Tabin

  This file is part of libzippp, a library that wraps libzip for manipulating easily
  ZIP files in C++.
  The author can be contacted on http://www.astorm.ch/blog/index.php?contact

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions
  are met:
  1. Redistributions of source code must retain the above copyright
     notice, this list of conditions and the following disclaimer.
  2. Redistributions in binary form must reproduce the above copyright
     notice, this list of conditions and the following disclaimer in
     the documentation and/or other materials provided with the
     distribution.
  3. The names of the authors may not be used to endorse or promote
     products derived from this software without specific prior
     written permission.
 
  THIS SOFTWARE IS PROVIDED BY THE AUTHORS ``AS IS'' AND ANY EXPRESS
  OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
  ARE DISCLAIMED.  IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
  DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE
  GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER
  IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
  OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
  IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

#include <cstdio>
#include <string>
#include <vector>

//defined in libzip
struct zip;

#define DIRECTORY_SEPARATOR '/'
#define IS_DIRECTORY(str) ((str).length()>0 && (str)[(str).length()-1]==DIRECTORY_SEPARATOR)
#define DEFAULT_CHUNK_SIZE 524288

//libzip documentation
//- http://www.nih.at/libzip/libzip.html
//- http://slash.developpez.com/tutoriels/c/utilisation-libzip/

//standard unsigned int
typedef unsigned int uint;

#ifdef WIN32
        typedef long long libzippp_int64;
        typedef unsigned long long libzippp_uint64;
        typedef unsigned short libzippp_uint16;
#else
        //standard ISO c++ does not support long long
        typedef long int libzippp_int64;
        typedef unsigned long int libzippp_uint64;
        typedef unsigned short libzippp_uint16;
#endif


// special return code for libzippp
#define LIBZIPPP_OK 0
#define LIBZIPPP_ERROR_NOT_OPEN -1
#define LIBZIPPP_ERROR_NOT_ALLOWED -2
#define LIBZIPPP_ERROR_INVALID_ENTRY -3
#define LIBZIPPP_ERROR_INVALID_PARAMETER -4
#define LIBZIPPP_ERROR_MEMORY_ALLOCATION -16
#define LIBZIPPP_ERROR_FOPEN_FAILURE -25
#define LIBZIPPP_ERROR_FREAD_FAILURE -26
#define LIBZIPPP_ERROR_OWRITE_FAILURE -35
#define LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE -36
#define LIBZIPPP_ERROR_UNKNOWN -99

namespace libzippp {
    class ZipEntry;
    
    /**
     * Represents a ZIP archive. This class provides useful methods to handle an archive
     * content. It is simply a wrapper around libzip.
     */
    class ZipArchive {
    public:
        
        /**
         * Defines how the zip file must be open.
         * NOT_OPEN is a special mode where the file is not open.
         * READ_ONLY is the basic mode to only read the archive.
         * WRITE will append to an existing archive or create a new one if it does not exist.
         * NEW will create a new archive or erase all the data if a previous one exists.
         */
        enum OpenMode {
            NOT_OPEN,
            READ_ONLY,
            WRITE,
            NEW
        };
        
        /**
         * Defines how the reading of the data should be made in the archive.
         * ORIGINAL will read the data of the original archive file, without any change.
         * CURRENT will read the current content of the archive.
         */
        enum State {
            ORIGINAL,
            CURRENT
        };
        
        /**
         * Creates a new ZipArchive with the given path. If the password is defined, it
         * will be used to read encrypted archive. It won't affect the files added
         * to the archive.
         * 
         * http://nih.at/listarchive/libzip-discuss/msg00219.html
         */
        explicit ZipArchive(const std::string& zipPath, const std::string& password="");
        virtual ~ZipArchive(void); //commit all the changes if open
        
        /**
         * Return the path of the ZipArchive.
         */
        std::string getPath(void) const { return path; }
        
        /**
         * Open the ZipArchive with the given mode. This method will return true if the operation
         * is successful, false otherwise. If the OpenMode is NOT_OPEN an invalid_argument
         * will be thrown. If the archive is already open, this method returns true only if the
         * mode is the same.
         */
        bool open(OpenMode mode=READ_ONLY, bool checkConsistency=false);
        
        /**
         * Closes the ZipArchive and releases all the resources held by it. If the ZipArchive was
         * not open previously, this method does nothing. If the archive was open in modification
         * and some were done, they will be committed.
         * This method returns LIBZIPPP_OK if the archive was successfully closed, otherwise it 
         * returns the value returned by the zip_close() function.
         */
        int close(void);
        
        /**
         * Closes the ZipArchive and releases all the resources held by it. If the ZipArchive was
         * not open previously, this method does nothing. If the archive was open in modification
         * and some were done, they will be rollbacked.
         */
        void discard(void);
        
        /**
         * Deletes the file denoted by the path. If the ZipArchive is open, all the changes will
         * be discarded and the file removed.
         */
        bool unlink(void);
        //bool delete(void) { return unlink(); } //delete is a reserved keyword
        
        /**
         * Returns true if the ZipArchive is currently open.
         */
        inline bool isOpen(void) const { return zipHandle!=NULL; }
        
        /**
         * Returns true if the ZipArchive is open and mutable.
         */
        inline bool isMutable(void) const { return isOpen() && mode!=NOT_OPEN && mode!=READ_ONLY; }
        
        /**
         * Returns true if the ZipArchive is encrypted. This method returns true only if
         * a password has been set in the constructor.
         */
        inline bool isEncrypted(void) const { return !password.empty(); }
        
        /**
         * Defines the comment of the archive. In order to set the comment, the archive
         * must have been open in WRITE or NEW mode. If the archive is not open, the getComment
         * method will return an empty string.
         */
        std::string getComment(State state=CURRENT) const;
        bool setComment(const std::string& comment) const;
        
        /**
         * Removes the comment of the archive, if any. The archive must have been open
         * in WRITE or NEW mode.
         */
        inline bool removeComment(void) const { return setComment(std::string()); }
        
        /**
         * Returns the number of entries in this zip file (folders are included).
         * The zip file must be open otherwise LIBZIPPP_ERROR_NOT_OPEN will be returned. 
         * If the state is ORIGINAL, then the number entries of the original archive are returned.
         * Any change will not be considered.
         * Note also that the deleted entries does not affect the result of this method
         * with the CURRENT state. For instance, if there are 3 entries and you delete one,
         * this method will still return 3. However, if you add one entry, it will return
         * 4 with the state CURRENT and 3 with the state ORIGINAL.
         * If you wanna know the "real" entries effectively in the archive, you might use
         * the getEntries method.
         */
        libzippp_int64 getNbEntries(State state=CURRENT) const;
        inline libzippp_int64 getEntriesCount(State state=CURRENT) const { return getNbEntries(state); }
        //libzippp_int64 size(State state=CURRENT) const { return getNbEntries(state); } //not clear enough => could be the size of the file instead...

        /**
         * Returns all the entries of the ZipArchive. If the state is ORIGINAL, then
         * returns the entries in the original archive, any change will not be considered.
         * The zip file must be open otherwise an empty vector will be returned.
         */
        std::vector<ZipEntry> getEntries(State state=CURRENT) const;
        
        /**
         * Return true if an entry with the specified name exists. If no such entry exists,
         * then false will be returned. If a directory is searched, the name must end with a '/' !
         * The zip file must be open otherwise false will be returned.
         */
        bool hasEntry(const std::string& name, bool excludeDirectories=false, bool caseSensitive=true, State state=CURRENT) const;
        
        /**
         * Return the ZipEntry for the specified entry name. If no such entry exists,
         * then a null-ZiPEntry will be returned. If a directory is searched, the name
         * must end with a '/' !
         * The zip file must be open otherwise a null-ZipEntry will be returned.
         */
        ZipEntry getEntry(const std::string& name, bool excludeDirectories=false, bool caseSensitive=true, State state=CURRENT) const;
        
        /**
         * Return the ZipEntry for the specified index. If the index is out of range,
         * then a null-ZipEntry will be returned.
         * The zip file must be open otherwise a null-ZipEntry will be returned.
         */
        ZipEntry getEntry(libzippp_int64 index, State state=CURRENT) const;
        
        /**
         * Defines the comment of the entry. If the ZipArchive is not open or the
         * entry is not linked to this archive, then an empty string or false will 
         * be returned.
         */
        std::string getEntryComment(const ZipEntry& entry, State state=CURRENT) const;
        bool setEntryComment(const ZipEntry& entry, const std::string& comment) const;
        
        /**
         * Defines the compression method of an entry. If the ZipArchive is not open
         * or the entry is not linked to this archive, false will be returned.
         **/
        bool isEntryCompressionEnabled(const ZipEntry& entry) const;
        bool setEntryCompressionEnabled(const ZipEntry& entry, bool value) const;
        
        /**
         * Read the specified ZipEntry of the ZipArchive and returns its content within
         * a char array. If there is an error while reading the entry, then null will be returned.
         * The data must be deleted by the developer once not used anymore. If the asText
         * is set to true, then the returned void* will be ended by a \0 (hence the size of
         * the returned array will be zipEntry.getSize()+1 or size+1 if the latter is specified).
         * The zip file must be open otherwise null will be returned. If the ZipEntry was not
         * created by this ZipArchive, null will be returned.
         */
        void* readEntry(const ZipEntry& zipEntry, bool asText=false, State state=CURRENT, libzippp_uint64 size=0) const;
        
        /**
         * Read the specified ZipEntry of the ZipArchive and returns its content within
         * a char array. If there is an error while reading the entry, then null will be returned.
         * The data must be deleted by the developer once not used anymore. If the asText
         * is set to true, then the returned void* will be ended by a \0 (hence the size of
         * the returned array will be zipEntry.getSize()+1 or size+1 if the latter is specified).
         * The zip file must be open otherwise null will be returned. If the ZipEntry was not
         * created by this ZipArchive, null will be returned. If the zipEntry does not exist,
         * this method returns NULL:
         */
        void* readEntry(const std::string& zipEntry, bool asText=false, State state=CURRENT, libzippp_uint64 size=0) const;
        
        /**
         * Read the specified ZipEntry of the ZipArchive and inserts its content in the provided reference to an already
         * opened std::ofstream, gradually, with chunks of size "chunksize" to reduce memory usage when dealing with big files.
         * The method returns LIBZIPPP_OK if the extraction has succeeded with no problems, LIBZIPPP_ERROR_INVALID_PARAMETER if the 
         * ofstream is not opened, LIBZIPPP_ERROR_NOT_OPEN if the archive is not opened, LIBZIPPP_ERROR_INVALID_ENTRY if the zipEntry 
         * doesn't belong to the archive, LIBZIPPP_ERROR_FOPEN_FAILURE if zip_fopen_index() has failed, LIBZIPPP_ERROR_MEMORY_ALLOCATION if 
         * a memory allocation has failed, LIBZIPPP_ERROR_FREAD_FAILURE if zip_fread() didn't succeed to read data, 
         * LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE if the last ofstream operation has failed, LIBZIPPP_ERROR_OWRITE_FAILURE if fread() didn't 
         * return the exact amount of requested bytes and -9 if the amount of extracted bytes didn't match the size of the file (unknown error).
         * If the provided chunk size is zero, it will be defaulted to DEFAULT_CHUNK_SIZE (512KB).
         * The method doesn't close the ofstream after the extraction.
         */
        int readEntry(const ZipEntry& zipEntry, std::ofstream& ofOutput, State state=CURRENT, libzippp_uint64 chunksize=DEFAULT_CHUNK_SIZE) const;

        /**
         * Deletes the specified entry from the zip file. If the entry is a folder, all its
         * subentries will be removed. This method returns the number of entries removed.
         * If the open mode does not allow a deletion, this method will return LIBZIPPP_ERROR_NOT_ALLOWED. 
         * If the ZipArchive is not open, LIBZIPPP_ERROR_NOT_OPEN will be returned. If the entry is not handled 
         * by this ZipArchive or is a null-ZipEntry, then LIBZIPPP_ERROR_INVALID_ENTRY will be returned.
         * If an error occurs during deletion, this method will return LIBZIPPP_ERROR_UNKNOWN.
         * Note that this method does not affect the result returned by getNbEntries !
         */
        int deleteEntry(const ZipEntry& entry) const;
        
        /**
         * Deletes the specified entry from the zip file. If the entry is a folder, all its
         * subentries will be removed. This method returns the number of entries removed.
         * If the open mode does not allow a deletion, this method will return LIBZIPPP_ERROR_NOT_ALLOWED. 
         * If the ZipArchive is not open, LIBZIPPP_ERROR_NOT_OPEN will be returned. If the entry is not handled 
         * by this ZipArchive or is a null-ZipEntry, then LIBZIPPP_ERROR_INVALID_ENTRY will be returned.
         * If an error occurs during deletion, this method will return LIBZIPPP_ERROR_UNKNOWN.
         * If the entry does not exist, this method returns LIBZIPPP_ERROR_INVALID_PARAMETER.
         * Note that this method does not affect the result returned by getNbEntries !
         */
        int deleteEntry(const std::string& entry) const;
        
        /**
         * Renames the entry with the specified newName. The method returns the number of entries
         * that have been renamed, LIBZIPPP_ERROR_INVALID_PARAMETER if the new name is invalid, 
         * LIBZIPPP_ERROR_NOT_ALLOWED if the mode doesn't allow modification or LIBZIPPP_ERROR_UNKNOWN if an error 
         * occurred.  If the entry is a directory, a '/' will automatically be appended at the end of newName if the 
         * latter hasn't it already. All the files in the folder will be moved.
         * If the ZipArchive is not open or the entry was not edited by this ZipArchive or is a null-ZipEntry,
         * then LIBZIPPP_ERROR_INVALID_ENTRY will be returned.
         */
        int renameEntry(const ZipEntry& entry, const std::string& newName) const;
        
        /**
         * RRenames the entry with the specified newName. The method returns the number of entries
         * that have been renamed, LIBZIPPP_ERROR_INVALID_PARAMETER if the new name is invalid, 
         * LIBZIPPP_ERROR_NOT_ALLOWED if the mode doesn't allow modification or LIBZIPPP_ERROR_UNKNOWN if an error 
         * occurred.  If the entry is a directory, a '/' will automatically be appended at the end of newName if the 
         * latter hasn't it already. All the files in the folder will be moved.
         * If the ZipArchive is not open or the entry was not edited by this ZipArchive or is a null-ZipEntry,
         * then LIBZIPPP_ERROR_INVALID_ENTRY will be returned. If the entry does not exist, this method returns LIBZIPPP_ERROR_INVALID_PARAMETER.
         */
        int renameEntry(const std::string& entry, const std::string& newName) const;
        
        /**
         * Add the specified file in the archive with the given entry. If the entry already exists,
         * it will be replaced. This method returns true if the file has been added successfully. 
         * If the entryName specifies folders that doesn't exist in the archive, they will be automatically created.
         * If the entryName denotes a directory, this method returns false.
         * The zip file must be open otherwise false will be returned.
         */
        bool addFile(const std::string& entryName, const std::string& file) const;
        
        /**
         * Add the given data to the specified entry name in the archive. If the entry already exists,
         * its content will be erased. 
         * If the entryName specifies folders that doesn't exist in the archive, they will be automatically created.
         * If the entryName denotes a directory, this method returns false.
         * If the zip file is not open, this method returns false.
         */
        bool addData(const std::string& entryName, const void* data, libzippp_uint64 length, bool freeData=false) const;
        
        /**
         * Add the specified entry to the ZipArchive. All the needed hierarchy will be created.
         * The entryName must be a directory (end with '/').
         * If the ZipArchive is not open or the entryName is not a directory, this method will
         * returns false. If the entry already exists, this method returns true.
         * This method will only add the specified entry. The 'real' directory may exist or not.
         * If the directory exists, the files in it won't be added to the archive.
         */
        bool addEntry(const std::string& entryName) const;
        
        /**
         * Returns the mode in which the file has been open.
         * If the archive is not open, then NOT_OPEN will be returned.
         */
        inline OpenMode getMode(void) const { return mode; }

    private:
        std::string path;
        zip* zipHandle;
        OpenMode mode;
        std::string password;
        
        //generic method to create ZipEntry
        ZipEntry createEntry(struct zip_stat* stat) const;
        
        //prevent copy across functions
        ZipArchive(const ZipArchive& zf);
        ZipArchive& operator=(const ZipArchive&);
    };
    
    /**
     * Represents an entry in a zip file.
     * This class is meant to be used by the ZipArchive class.
     */
    class ZipEntry {
    friend class ZipArchive;
    public:
        /**
         * Creates a new null-ZipEntry. Only a ZipArchive will create a valid ZipEntry
         * usable to read and modify an archive.
         */
        explicit ZipEntry(void) : zipFile(NULL), index(0), time(0), compressionMethod(-1), encryptionMethod(-1), size(0), sizeComp(0), crc(0)  {}
        virtual ~ZipEntry(void) {}
        
        /**
         * Returns the name of the entry.
         */
        inline std::string getName(void) const { return name; }
        
        /**
         * Returns the index of the file in the zip.
         */
        inline libzippp_uint64 getIndex(void) const { return index; }
        
        /**
         * Returns the timestamp of the entry.
         */
        inline time_t getDate(void) const { return time; }
        
        /**
         * Returns the compression method.
         */
        inline libzippp_uint16 getCompressionMethod(void) const { return compressionMethod; }
        
        /**
         * Returns the encryption method.
         */
        inline libzippp_uint16 getEncryptionMethod(void) const { return encryptionMethod; }
        
        /**
         * Returns the size of the file (not deflated).
         */
        inline libzippp_uint64 getSize(void) const { return size; }
        
        /**
         * Returns the size of the inflated file.
         */
        inline libzippp_uint64 getInflatedSize(void) const { return sizeComp; }
        
        /**
         * Returns the CRC of the file.
         */
        inline int getCRC(void) const { return crc; }
        
        /**
         * Returns true if the entry is a directory.
         */
        inline bool isDirectory(void) const { return IS_DIRECTORY(name); }
        
        /**
         * Returns true if the entry is a file.
         */
        inline bool isFile(void) const { return !isDirectory(); }
        
        /**
         * Returns true if this entry is null (means no more entry is available).
         */
        inline bool isNull(void) const { return zipFile==NULL; }
        
        /**
         * Defines if the compression is enabled for this entry.
         * Those methods are wrappers arount ZipArchive::isEntryCompressionEnabled and
         * ZipArchive::setEntryCompressionEnabled.
         */
        bool isCompressionEnabled(void) const;
        bool setCompressionEnabled(bool value) const;
        
        /**
         * Defines the comment of the entry. In order to call either one of those
         * methods, the corresponding ZipArchive must be open otherwise an empty string
         * or false will be returned. Those methods are wrappers around ZipArchive::getEntryComment
         * and ZipArchive::setEntryComment.
         */
        std::string getComment(void) const;
        bool setComment(const std::string& str) const;
        
        /**
         * Read the content of this ZipEntry as text. 
         * The returned string will be of size getSize() if the latter is not specified or too big. 
         * If the ZipArchive is not open, this method returns an
         * empty string. This method is a wrapper around ZipArchive::readEntry(...).
         */
        std::string readAsText(ZipArchive::State state=ZipArchive::CURRENT, libzippp_uint64 size=0) const;
        
        /**
         * Read the content of this ZipEntry as binary. 
         * The returned void* will be of size getSize() if the latter is not specified or too big.
         * If the ZipArchive is not open, this method returns NULL. 
         * The data must be deleted by the developer once not used anymore.
         * This method is a wrapper around ZipArchive::readEntry(...).
         */
        void* readAsBinary(ZipArchive::State state=ZipArchive::CURRENT, libzippp_uint64 size=0) const;
        
        /**
         * Read the specified ZipEntry of the ZipArchive and inserts its content in the provided reference to an already
         * opened std::ofstream, gradually, with chunks of size "chunksize" to reduce memory usage when dealing with big files.
         * The method returns LIBZIPPP_OK if the extraction has succeeded with no problems, LIBZIPPP_ERROR_INVALID_PARAMETER if the 
         * ofstream is not opened, LIBZIPPP_ERROR_NOT_OPEN if the archive is not opened, LIBZIPPP_ERROR_INVALID_ENTRY if the zipEntry 
         * doesn't belong to the archive, LIBZIPPP_ERROR_FOPEN_FAILURE if zip_fopen_index() has failed, LIBZIPPP_ERROR_MEMORY_ALLOCATION if 
         * a memory allocation has failed, LIBZIPPP_ERROR_FREAD_FAILURE if zip_fread() didn't succeed to read data, 
         * LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE if the last ofstream operation has failed, LIBZIPPP_ERROR_OWRITE_FAILURE if fread() didn't 
         * return the exact amount of requested bytes and -9 if the amount of extracted bytes didn't match the size of the file (unknown error).
         * If the provided chunk size is zero, it will be defaulted to DEFAULT_CHUNK_SIZE (512KB).
         * The method doesn't close the ofstream after the extraction.
         */
        int readContent(std::ofstream& ofOutput, ZipArchive::State state=ZipArchive::CURRENT, libzippp_uint64 chunksize=DEFAULT_CHUNK_SIZE) const;
        
    private:
        const ZipArchive* zipFile;
        std::string name;
        libzippp_uint64 index;
        time_t time;
        libzippp_uint16 compressionMethod;
        libzippp_uint16 encryptionMethod;
        libzippp_uint64 size;
        libzippp_uint64 sizeComp;
        int crc;
        
        ZipEntry(const ZipArchive* zipFile, const std::string& name, libzippp_uint64 index, time_t time, libzippp_uint16 compMethod, libzippp_uint16 encMethod, libzippp_uint64 size, libzippp_uint64 sizeComp, int crc) : 
                zipFile(zipFile), name(name), index(index), time(time), compressionMethod(compMethod), encryptionMethod(encMethod), size(size), sizeComp(sizeComp), crc(crc) {}
    };
}

#endif

