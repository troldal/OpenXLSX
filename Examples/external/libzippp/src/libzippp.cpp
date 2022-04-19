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

#include <zip.h>
#include <errno.h>
#include <fstream>
#include <memory>

#include "libzippp.h"

using namespace libzippp;
using namespace std;

#define NEW_CHAR_ARRAY(nb) new (std::nothrow) char[(nb)];

string ZipEntry::getComment(void) const {
    return zipFile->getEntryComment(*this);
}

bool ZipEntry::setComment(const string& str) const {
    return zipFile->setEntryComment(*this, str);
}

bool ZipEntry::isCompressionEnabled(void) const {
    return zipFile->isEntryCompressionEnabled(*this);
}

bool ZipEntry::setCompressionEnabled(bool value) const {
    return zipFile->setEntryCompressionEnabled(*this, value);
}

string ZipEntry::readAsText(ZipArchive::State state, libzippp_uint64 size) const {
    char* content = (char*)zipFile->readEntry(*this, true, state, size);
    if (content==nullptr) { return string(); } //happen if the ZipArchive has been closed
    
    libzippp_uint64 maxSize = getSize();
    string str(content, (size==0 || size>maxSize ? maxSize : size));
    delete[] content;
    return str;
}

void* ZipEntry::readAsBinary(ZipArchive::State state, libzippp_uint64 size) const {
    return zipFile->readEntry(*this, false, state, size); 
}

int ZipEntry::readContent(std::ostream& ofOutput, ZipArchive::State state, libzippp_uint64 chunksize) const {
   return zipFile->readEntry(*this, ofOutput, state, chunksize);
}

ZipArchive::ZipArchive(const string& zipPath, const string& password, Encryption encryptionMethod) : path(zipPath), zipHandle(nullptr), zipSource(nullptr), mode(NotOpen), password(password), progressPrecision(LIBZIPPP_DEFAULT_PROGRESSION_PRECISION), bufferData(nullptr), bufferLength(-1) {
    switch(encryptionMethod) {
#ifdef LIBZIPPP_WITH_ENCRYPTION
        case Encryption::Aes128:
            this->encryptionMethod = ZIP_EM_AES_128;
            break;
        case Encryption::Aes192:
            this->encryptionMethod = ZIP_EM_AES_192;
            break;
        case Encryption::Aes256:
            this->encryptionMethod = ZIP_EM_AES_256;
            break;
        case Encryption::TradPkware:
            this->encryptionMethod = ZIP_EM_TRAD_PKWARE;
            break;
#endif
        case Encryption::None:
        default:
            this->encryptionMethod = ZIP_EM_NONE;
            break;
    }
}

ZipArchive::~ZipArchive(void) { 
    close(); /* discard ??? */ 
    
    // ensures all the values are clared
    zipHandle = nullptr;
    zipSource = nullptr;
    bufferData = nullptr;
    listeners.clear();
}

ZipArchive* ZipArchive::fromBuffer(void* data, libzippp_uint32 size, OpenMode om, bool checkConsistency) {
    ZipArchive* za = new ZipArchive("");
    bool o = za->openBuffer(data, size, om, checkConsistency);
    if(!o) {
        delete za;
        za = nullptr;
    }
    return za;
}

ZipArchive* ZipArchive::fromSource(zip_source* source, OpenMode om, bool checkConsistency) {
    ZipArchive* za = new ZipArchive("");
    bool o = za->openSource(source, om, checkConsistency);
    if(!o) {
        delete za;
        za = nullptr;
    }
    return za;
}

bool ZipArchive::openBuffer(void* data, libzippp_uint32 size, OpenMode om, bool checkConsistency) {
    zip_error_t error;
    zip_error_init(&error);

    /* create source from buffer */
    zip_source* zipSource = zip_source_buffer_create(data, size, 0, &error);
    if (zipSource == nullptr) {
        fprintf(stderr, "can't create zip source: %s\n", zip_error_strerror(&error));
        zip_error_fini(&error);
        return false;
    }
    
    bool open = openSource(zipSource, om, checkConsistency);
    if(open && (om==Write || om==New)) {
        bufferData = data;
        bufferLength = size;
        
        //prevents libzip to delete the source
        zip_source_keep(zipSource);
    }
    return open;
}

bool ZipArchive::openSource(zip_source* source, OpenMode om, bool checkConsistency) {
    int zipFlag = 0;
    if (om == ReadOnly) { zipFlag = 0; }
    else if (om == Write) { zipFlag = ZIP_CREATE; }
    else if (om == New) { zipFlag = ZIP_CREATE | ZIP_TRUNCATE; }
    else { return false; }
    if (checkConsistency) {
        zipFlag = zipFlag | ZIP_CHECKCONS;
    }
    
    zip_error_t error;
    zip_error_init(&error);
    
    /* open zip archive from source */
    zipHandle = zip_open_from_source(source, zipFlag, &error);
    if (zipHandle == nullptr) {
        fprintf(stderr, "can't open zip from source: %s\n", zip_error_strerror(&error));
        zip_source_free(zipSource);
        zipSource = nullptr;
        zip_error_fini(&error);
        return false;
    }
    zip_error_fini(&error);
    
    zipSource = source;

#ifdef LIBZIPPP_WITH_ENCRYPTION
    if (isEncrypted()) {
        int result = zip_set_default_password(zipHandle, password.c_str());
        if (result != 0) {
            close();
            return false;
        }
    }
#endif

    mode = om;
    return true;
}

bool ZipArchive::open(OpenMode om, bool checkConsistency) {
    if (isOpen()) { return om==mode; }
  
    int zipFlag = 0;
    if (om==ReadOnly) { zipFlag = 0; }
    else if (om==Write) { zipFlag = ZIP_CREATE; }
    else if (om==New) { zipFlag = ZIP_CREATE | ZIP_TRUNCATE; }
    else { return false; }
    
    if (checkConsistency) {
        zipFlag = zipFlag | ZIP_CHECKCONS;
    }
    
    int errorFlag = 0;
    zipHandle = zip_open(path.c_str(), zipFlag, &errorFlag);
    
    //error during opening of the file
    if (errorFlag!=ZIP_ER_OK) {
        /*char* errorStr = new char[256];
        zip_error_to_str(errorStr, 255, errorFlag, errno);
        errorStr[255] = '\0';
        cout << "Error: " << errorStr << endl;*/
        
        zipHandle = nullptr;
        return false;
    }
    
    if (zipHandle!=nullptr) {
#ifdef LIBZIPPP_WITH_ENCRYPTION
        if (isEncrypted()) {
            int result = zip_set_default_password(zipHandle, password.c_str());
            if (result!=0) { 
                close();
                return false;
            }
        }
#endif
        
        mode = om;
        return true;
    }
    
    return false;
}

void progress_callback(zip* archive, double progression, void* ud) {
    ZipArchive* za = static_cast<ZipArchive*>(ud);
    vector<ZipProgressListener*> listeners = za->getProgressListeners();
    for(vector<ZipProgressListener*>::const_iterator it=listeners.begin() ; it!=listeners.end() ; ++it) {
        ZipProgressListener* listener = *it;
        listener->progression(progression);
    }
}

int ZipArchive::close(void) {
    if (isOpen()) {

        //do not handle zipSource at all because it will be deleted by libzip
        //directly when not necessary anymore

        if(!listeners.empty()) {
            zip_register_progress_callback_with_state(zipHandle, progressPrecision, progress_callback, nullptr, this);
        }

        progress_callback(zipHandle, 0, this); //enforce the first progression call to be zero
        
        int result = zip_close(zipHandle);
        zipHandle = nullptr;
        progress_callback(zipHandle, 1, this); //enforce the last progression call to be one
        
        if (result!=0) { return result; }

        //push back the changes in the buffer
        if(bufferData!=nullptr && (mode==New || mode==Write)) {
            int srcOpen = zip_source_open(zipSource);
            if(srcOpen==0) {            
                int newLength = zip_source_read(zipSource, bufferData, bufferLength);
                zip_source_close(zipSource);
                zip_source_free(zipSource);
            
                bufferLength = newLength;
            } else {
                fprintf(stderr, "can't read back from source: %d\n", srcOpen);
                return srcOpen;
            }
        }
        
        mode = NotOpen;
    }
    
    return LIBZIPPP_OK;
}

void ZipArchive::discard(void) {
    if (isOpen()) {
        zip_discard(zipHandle);
        zipHandle = nullptr;
        mode = NotOpen;
    }
}

bool ZipArchive::unlink(void) {
    if (isOpen()) { discard(); }
    int result = remove(path.c_str());
    return result==0;
}

string ZipArchive::getComment(State state) const {
    if (!isOpen()) { return string(); }
    
    int flag = 0;
    if (state==Original) { flag = flag | LIBZIPPP_ORIGINAL_STATE_FLAGS; }
    else { flag = flag | ZIP_FL_ENC_GUESS; }
    
    int length = 0;
    const char* comment = zip_get_archive_comment(zipHandle, &length, flag);
    if (comment==nullptr) { return string(); }
    return string(comment, length);
}

bool ZipArchive::setComment(const string& comment) const {
    if (!isOpen()) { return false; }
    if (mode==ReadOnly) { return false; }
    
    int size = comment.size();
    const char* data = comment.c_str();
    int result = zip_set_archive_comment(zipHandle, data, size);
    return result==0;
}

bool ZipArchive::isEntryCompressionEnabled(const ZipEntry& entry) const {
    return entry.compressionMethod==ZIP_CM_DEFLATE;
}

bool ZipArchive::setEntryCompressionEnabled(const ZipEntry& entry, bool value) const {
    if (!isOpen()) { return false; }
    if (entry.zipFile!=this) { return false; }
    if (mode==ReadOnly) { return false; }
    
    libzippp_uint16 compMode = value ? ZIP_CM_DEFLATE : ZIP_CM_STORE;
    return zip_set_file_compression(zipHandle, entry.index, compMode, 0)==0;
}

libzippp_int64 ZipArchive::getNbEntries(State state) const {
    if (!isOpen()) { return LIBZIPPP_ERROR_NOT_OPEN; }
    
    int flag = state==Original ? LIBZIPPP_ORIGINAL_STATE_FLAGS : 0;
    return zip_get_num_entries(zipHandle, flag);
}

ZipEntry ZipArchive::createEntry(struct zip_stat* stat) const {
    string name(stat->name);
    libzippp_uint64 index = stat->index;
    libzippp_uint64 size = stat->size;
    libzippp_uint16 compMethod = stat->comp_method;
    libzippp_uint16 encMethod = stat->encryption_method;
    libzippp_uint64 sizeComp = stat->comp_size;
    int crc = stat->crc;
    time_t time = stat->mtime;

    return ZipEntry(this, name, index, time, compMethod, encMethod, size, sizeComp, crc);
}

vector<ZipEntry> ZipArchive::getEntries(State state) const {
    if (!isOpen()) { return vector<ZipEntry>(); }
    
    struct zip_stat stat;
    zip_stat_init(&stat);

    vector<ZipEntry> entries;
    int flag = state==Original ? LIBZIPPP_ORIGINAL_STATE_FLAGS : ZIP_FL_ENC_GUESS;
    libzippp_int64 nbEntries = getNbEntries(state);
    for(libzippp_int64 i=0 ; i<nbEntries ; ++i) {
        int result = zip_stat_index(zipHandle, i, flag, &stat);
        if (result==0) {
            ZipEntry entry = createEntry(&stat);
            entries.push_back(entry);
        } else {
            //TODO handle read error => crash ?
        }
    }
    return entries;
}

bool ZipArchive::hasEntry(const string& name, bool excludeDirectories, bool caseSensitive, State state) const {
    if (!isOpen()) { return false; }
    
    int flags = 0;
    if (excludeDirectories) { flags = flags | ZIP_FL_NODIR; }
    if (!caseSensitive) { flags = flags | ZIP_FL_NOCASE; }
    if (state==Original) { flags = flags | LIBZIPPP_ORIGINAL_STATE_FLAGS; }
    else { flags = flags | ZIP_FL_ENC_GUESS; }
    
    libzippp_int64 index = zip_name_locate(zipHandle, name.c_str(), flags);
    return index>=0;
}

ZipEntry ZipArchive::getEntry(const string& name, bool excludeDirectories, bool caseSensitive, State state) const {
    if (isOpen()) {
        int flags = 0;
        if (excludeDirectories) { flags = flags | ZIP_FL_NODIR; }
        if (!caseSensitive) { flags = flags | ZIP_FL_NOCASE; }
        if (state==Original) { flags = flags | LIBZIPPP_ORIGINAL_STATE_FLAGS; }
        else { flags = flags | ZIP_FL_ENC_GUESS; }

        libzippp_int64 index = zip_name_locate(zipHandle, name.c_str(), flags);
        if (index>=0) {
            return getEntry(index);
        } else {
            //name not found
        }
    }
    return ZipEntry();
}
        
ZipEntry ZipArchive::getEntry(libzippp_int64 index, State state) const {
    if (isOpen()) {
        struct zip_stat stat;
        zip_stat_init(&stat);
        int flag = state==Original ? LIBZIPPP_ORIGINAL_STATE_FLAGS : ZIP_FL_ENC_GUESS;
        int result = zip_stat_index(zipHandle, index, flag, &stat);
        if (result==0) {
            return createEntry(&stat);
        } else {
            //index not found / invalid index
        }
    }
    return ZipEntry();
}

string ZipArchive::getEntryComment(const ZipEntry& entry, State state) const {
    if (!isOpen()) { return string(); }
    if (entry.zipFile!=this) { return string(); }
    
    int flag = 0;
    if (state==Original) { flag = flag | LIBZIPPP_ORIGINAL_STATE_FLAGS; }
    else { flag = ZIP_FL_ENC_GUESS; }
    
    unsigned int clen;
    const char* com = zip_file_get_comment(zipHandle, entry.getIndex(), &clen, flag);
    string comment = com==nullptr ? string() : string(com, clen);
    return comment;
}

bool ZipArchive::setEntryComment(const ZipEntry& entry, const string& comment) const {
    if (!isOpen()) { return false; }
    if (entry.zipFile!=this) { return false; }
    
    bool result = zip_file_set_comment(zipHandle, entry.getIndex(), comment.c_str(), comment.size(), ZIP_FL_ENC_GUESS);
    return result==0;
}

void* ZipArchive::readEntry(const ZipEntry& zipEntry, bool asText, State state, libzippp_uint64 size) const {
    if (!isOpen()) { return nullptr; }
    if (zipEntry.zipFile!=this) { return nullptr; }
    
    int flag = state==Original ? LIBZIPPP_ORIGINAL_STATE_FLAGS : ZIP_FL_ENC_GUESS;
    struct zip_file* zipFile = zip_fopen_index(zipHandle, zipEntry.getIndex(), flag);
    if (zipFile) {
        libzippp_uint64 maxSize = zipEntry.getSize();
        libzippp_uint64 uisize = size==0 || size>maxSize ? maxSize : size;
        
        char* data = NEW_CHAR_ARRAY(uisize+(asText ? 1 : 0))
        if(!data) { //allocation error
            zip_fclose(zipFile);
            return nullptr; 
        }

        libzippp_int64 result = zip_fread(zipFile, data, uisize);
        zip_fclose(zipFile);
        
        //avoid buffer copy
        if (asText) { data[uisize] = '\0'; }
        
        libzippp_int64 isize = (libzippp_int64)uisize;
        if (result==isize) {
            return data;
        } else { //unexpected number of bytes read => crash ?
            delete[] data;
        }
    } else {
        //unable to read the entry => crash ?
    }     
    
    return nullptr;
}

void* ZipArchive::readEntry(const string& zipEntry, bool asText, State state, libzippp_uint64 size) const {
    ZipEntry entry = getEntry(zipEntry);
    if (entry.isNull()) { return nullptr; }
    return readEntry(entry, asText, state, size);
}

int ZipArchive::deleteEntry(const ZipEntry& entry) const {
    if (!isOpen()) { return LIBZIPPP_ERROR_NOT_OPEN; }
    if (entry.zipFile!=this) { return LIBZIPPP_ERROR_INVALID_ENTRY; }
    if (mode==ReadOnly) { return LIBZIPPP_ERROR_NOT_ALLOWED; } //deletion not allowed
    
    if (entry.isFile()) {
        int result = zip_delete(zipHandle, entry.getIndex());
        if (result==0) { return 1; }
        return LIBZIPPP_ERROR_UNKNOWN; //unable to delete the entry
    } else {
        int counter = 0;
        vector<ZipEntry> allEntries = getEntries();
        vector<ZipEntry>::const_iterator eit;
        for(eit=allEntries.begin() ; eit!=allEntries.end() ; ++eit) {
            ZipEntry ze = *eit;
            int startPosition = ze.getName().find(entry.getName());
            if (startPosition==0) {
                int result = zip_delete(zipHandle, ze.getIndex());
                if (result==0) { ++counter; }
                else { return LIBZIPPP_ERROR_UNKNOWN; } //unable to remove the current entry
            }
        }
        return counter;
    }
}

int ZipArchive::deleteEntry(const string& e) const {
    ZipEntry entry = getEntry(e);
    if (entry.isNull()) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    return deleteEntry(entry);
}

int ZipArchive::renameEntry(const ZipEntry& entry, const string& newName) const {
    if (!isOpen()) { return LIBZIPPP_ERROR_NOT_OPEN; }
    if (entry.zipFile!=this) { return LIBZIPPP_ERROR_INVALID_ENTRY; }
    if (mode==ReadOnly) { return LIBZIPPP_ERROR_NOT_ALLOWED; } //renaming not allowed
    if (newName.length()==0) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    if (newName==entry.getName()) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    
    if (entry.isFile()) {
        if (LIBZIPPP_ENTRY_IS_DIRECTORY(newName)) { return LIBZIPPP_ERROR_INVALID_PARAMETER; } //invalid new name
        
        int lastSlash = newName.rfind(LIBZIPPP_ENTRY_PATH_SEPARATOR);
        if (lastSlash!=1) { 
            bool dadded = addEntry(newName.substr(0, lastSlash+1)); 
            if (!dadded) { return LIBZIPPP_ERROR_UNKNOWN; } //the hierarchy hasn't been created
        }
        
        int result = zip_file_rename(zipHandle, entry.getIndex(), newName.c_str(), ZIP_FL_ENC_GUESS);
        if (result==0) { return 1; }
        return LIBZIPPP_ERROR_UNKNOWN; //renaming was not possible (entry already exists ?)
    } else {
        if (!LIBZIPPP_ENTRY_IS_DIRECTORY(newName)) { return LIBZIPPP_ERROR_INVALID_PARAMETER; } //invalid new name
        
        int parentSlash = newName.rfind(LIBZIPPP_ENTRY_PATH_SEPARATOR, newName.length()-2);
        if (parentSlash!=-1) { //updates the dir hierarchy
            string parent = newName.substr(0, parentSlash+1);
            bool dadded = addEntry(parent);
            if (!dadded) { return LIBZIPPP_ERROR_UNKNOWN; }
        }
        
        int counter = 0;
        string originalName = entry.getName();
        vector<ZipEntry> allEntries = getEntries();
        vector<ZipEntry>::const_iterator eit;
        for(eit=allEntries.begin() ; eit!=allEntries.end() ; ++eit) {
            ZipEntry ze = *eit;
            string currentName = ze.getName();
            
            int startPosition = currentName.find(originalName);
            if (startPosition==0) {
                if (currentName == originalName) {
                    int result = zip_file_rename(zipHandle, entry.getIndex(), newName.c_str(), ZIP_FL_ENC_GUESS);
                    if (result==0) { ++counter; }
                    else { return LIBZIPPP_ERROR_UNKNOWN;  } //unable to rename the folder
                } else  {
                    string targetName = currentName.replace(0, originalName.length(), newName);
                    int result = zip_file_rename(zipHandle, ze.getIndex(), targetName.c_str(), ZIP_FL_ENC_GUESS);
                    if (result==0) { ++counter; }
                    else { return LIBZIPPP_ERROR_UNKNOWN; } //unable to rename a sub-entry
                }
            } else {
                //file not affected by the renaming
            }
        }
        
        /*
         * Special case for moving a directory a/x to a/x/y to avoid to lose
         * the a/x path in the archive.
         */
        bool newNameIsInsideCurrent = (newName.find(entry.getName())==0);
        if (newNameIsInsideCurrent) {
            bool dadded = addEntry(newName);
            if (!dadded) { return LIBZIPPP_ERROR_UNKNOWN; }
        }
        
        return counter;
    }
}

int ZipArchive::renameEntry(const string& e, const string& newName) const {
    ZipEntry entry = getEntry(e);
    if (entry.isNull()) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    return renameEntry(entry, newName);
}

bool ZipArchive::addFile(const string& entryName, const string& file) const {
    if (!isOpen()) { return false; }
    if (mode==ReadOnly) { return false; } //adding not allowed
    if (LIBZIPPP_ENTRY_IS_DIRECTORY(entryName)) { return false; }
    
    int lastSlash = entryName.rfind(LIBZIPPP_ENTRY_PATH_SEPARATOR);
    if (lastSlash!=-1) { //creates the needed parent directories
        string dirEntry = entryName.substr(0, lastSlash+1);
        bool dadded = addEntry(dirEntry);
        if (!dadded) { return false; }
    }
    
    const char* filepath = file.c_str();
    zip_source* source = zip_source_file(zipHandle, filepath, 0, -1);
    if (source!=nullptr) {
        libzippp_int64 result = zip_file_add(zipHandle, entryName.c_str(), source, ZIP_FL_OVERWRITE);
        if (result>=0) {
#ifdef LIBZIPPP_WITH_ENCRYPTION
            if(isEncrypted()) {
                if(zip_file_set_encryption(zipHandle,result,encryptionMethod,nullptr)!=0) { //unable to encrypt
                    zip_source_free(source);
                } else {
                    return true;
                }
            } else {
                return true;
            }
#else
            return true;
#endif
        } else {
            //unable to add the file
            zip_source_free(source);
        }
    } else {
        //unable to create the zip_source
    }
    return false;
}

bool ZipArchive::addData(const string& entryName, const void* data, libzippp_uint64 length, bool freeData) const {
    if (!isOpen()) { return false; }
    if (mode==ReadOnly) { return false; } //adding not allowed
    if (LIBZIPPP_ENTRY_IS_DIRECTORY(entryName)) { return false; }
    
    int lastSlash = entryName.rfind(LIBZIPPP_ENTRY_PATH_SEPARATOR);
    if (lastSlash!=-1) { //creates the needed parent directories
        string dirEntry = entryName.substr(0, lastSlash+1);
        bool dadded = addEntry(dirEntry);
        if (!dadded) { return false; }
    }
    
    zip_source* source = zip_source_buffer(zipHandle, data, length, freeData);
    if (source!=nullptr) {
        libzippp_int64 result = zip_file_add(zipHandle, entryName.c_str(), source, ZIP_FL_OVERWRITE);
        if (result>=0) {
#ifdef LIBZIPPP_WITH_ENCRYPTION
            if(isEncrypted()) {
                if(zip_file_set_encryption(zipHandle,result,encryptionMethod,nullptr)!=0) { //unable to encrypt
                    zip_source_free(source);
                } else {
                    return true;
                }
            } else {
                return true;
            }
#else
            return true;
#endif
        } else {
            //unable to add the file
            zip_source_free(source);
        }
    } else {
        //unable to create the zip_source
    }
    return false;
}

bool ZipArchive::addEntry(const string& entryName) const {
    if (!isOpen()) { return false; }
    if (mode==ReadOnly) { return false; } //adding not allowed
    if (!LIBZIPPP_ENTRY_IS_DIRECTORY(entryName)) { return false; }
    
    int nextSlash = entryName.find(LIBZIPPP_ENTRY_PATH_SEPARATOR);
    while (nextSlash!=-1) {
        string pathToCreate = entryName.substr(0, nextSlash+1);
        if (!hasEntry(pathToCreate)) {
            libzippp_int64 result = zip_dir_add(zipHandle, pathToCreate.c_str(), ZIP_FL_ENC_GUESS);
            if (result==-1) { return false; }
        }
        nextSlash = entryName.find(LIBZIPPP_ENTRY_PATH_SEPARATOR, nextSlash+1);
    }
    
    return true;
}

void ZipArchive::removeProgressListener(ZipProgressListener* listener) {
    for(vector<ZipProgressListener*>::const_iterator it=listeners.begin() ; it!=listeners.end() ; ++it) {
        ZipProgressListener* l = *it;
        if(l==listener) {
            listeners.erase(it);
            break;
        }
    }
}

int ZipArchive::readEntry(const ZipEntry& zipEntry, std::ostream& ofOutput, State state, libzippp_uint64 chunksize) const {
    if (!ofOutput) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    std::function<bool(const void*,libzippp_uint64)> writeFunc = [&ofOutput](const void* data,libzippp_uint64 size){ ofOutput.write((char*)data, size); return bool(ofOutput); };
    return readEntry(zipEntry, writeFunc, state, chunksize);
}

int ZipArchive::readEntry(const ZipEntry& zipEntry, std::function<bool(const void*,libzippp_uint64)> writeFunc, State state, libzippp_uint64 chunksize) const {
    if (!isOpen()) { return LIBZIPPP_ERROR_NOT_OPEN; }
    if (zipEntry.zipFile!=this) { return LIBZIPPP_ERROR_INVALID_ENTRY; }
    
    int iRes = LIBZIPPP_OK;
    int flag = state==Original ? LIBZIPPP_ORIGINAL_STATE_FLAGS : ZIP_FL_ENC_GUESS;
    struct zip_file* zipFile = zip_fopen_index(zipHandle, zipEntry.getIndex(), flag);
    if (zipFile) {
        libzippp_uint64 maxSize = zipEntry.getSize();
        if (!chunksize) { chunksize = LIBZIPPP_DEFAULT_CHUNK_SIZE; } // use the default chunk size (512K) if not specified by the user
    
        if (maxSize<chunksize) {
            char* data = NEW_CHAR_ARRAY(maxSize)
            if (data!=nullptr) {
                libzippp_int64 result = zip_fread(zipFile, data, maxSize);
                if (result>0) {
                    if (result != static_cast<libzippp_int64>(maxSize)) {
                        iRes = LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE;
                    } else if (!writeFunc(data, maxSize)) {
                        iRes = LIBZIPPP_ERROR_OWRITE_FAILURE;
                    }
                } else {
                    iRes = LIBZIPPP_ERROR_FREAD_FAILURE;
                }
                delete[] data;
            } else {
                iRes = LIBZIPPP_ERROR_MEMORY_ALLOCATION;
            }
        } else {
            libzippp_uint64 uWrittenBytes = 0;
            libzippp_int64 result = 0;
            char* data = NEW_CHAR_ARRAY(chunksize)
            if (data!=nullptr) {
                int nbChunks = maxSize/chunksize;
                for (int uiChunk=0 ; uiChunk<nbChunks ; ++uiChunk) {
                    result = zip_fread(zipFile, data, chunksize);
                    if (result>0) {
                        if (result!=static_cast<libzippp_int64>(chunksize)) {
                            iRes = LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE;
                            break;
                        } else {
                            if (!writeFunc(data, chunksize)) {
                                iRes = LIBZIPPP_ERROR_OWRITE_FAILURE;
                                break;
                            }
                            uWrittenBytes += result;
                        }
                    } else {
                        iRes = LIBZIPPP_ERROR_FREAD_FAILURE;
                        break;
                    }
                }
                delete[] data;
            } else {
                iRes = LIBZIPPP_ERROR_MEMORY_ALLOCATION;
            }
            
            int leftOver = maxSize%chunksize;
            if (iRes==0 && leftOver>0) {
                char* data = NEW_CHAR_ARRAY(leftOver);
                if (data!=nullptr) {
                    result = zip_fread(zipFile, data, leftOver);
                    if (result>0) {
                        if (result!=static_cast<libzippp_int64>(leftOver)) {
                            iRes = LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE;
                        } else {
                            if (!writeFunc(data, leftOver)) {
                                iRes = LIBZIPPP_ERROR_OWRITE_FAILURE;
                            } else {
                                uWrittenBytes += result;
                                if (uWrittenBytes!=maxSize) {
                                    iRes = LIBZIPPP_ERROR_UNKNOWN; // shouldn't occur but let's be careful
                                }
                            }
                        }
                    } else {
                        iRes = LIBZIPPP_ERROR_FREAD_FAILURE;
                    }
                } else {
                    iRes = LIBZIPPP_ERROR_MEMORY_ALLOCATION;
                }
                delete[] data;
            }
        }
        zip_fclose(zipFile);
    } else {
       iRes = LIBZIPPP_ERROR_FOPEN_FAILURE;
    }
    return iRes;
}
