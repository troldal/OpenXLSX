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

#include "@libzip/lib/zip.h"
#include <fstream>

#include "libzip++.h"

using namespace libzippp;
using namespace std;

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
    if (content==NULL) { return string(); } //happen if the ZipArchive has been closed
    
    libzippp_uint64 maxSize = getSize();
    string str(content, (size==0 || size>maxSize ? maxSize : size));
    delete[] content;
    return str;
}

void* ZipEntry::readAsBinary(ZipArchive::State state, libzippp_uint64 size) const {
    return zipFile->readEntry(*this, false, state, size); 
}

int ZipEntry::readContent(std::ofstream& ofOutput, ZipArchive::State state, libzippp_uint64 chunksize) const {
   return zipFile->readEntry(*this, ofOutput, state, chunksize);
}

ZipArchive::ZipArchive(const string& zipPath, const string& password) : path(zipPath), zipHandle(NULL), mode(NOT_OPEN), password(password) {
}

ZipArchive::~ZipArchive(void) { 
    close(); /* discard ??? */ 
}

bool ZipArchive::open(OpenMode om, bool checkConsistency) {
    if (isOpen()) { return om==mode; }
  
    int zipFlag = 0;
    if (om==READ_ONLY) { zipFlag = 0; }
    else if (om==WRITE) { zipFlag = ZIP_CREATE; }
    else if (om==NEW) { zipFlag = ZIP_CREATE | ZIP_TRUNCATE; }
    else { return false; }
    
    if (checkConsistency) {
        zipFlag = zipFlag | ZIP_CHECKCONS;
    }
    
    int errorFlag = 0;
    zipHandle = zip_open(path.c_str(), zipFlag, &errorFlag);
    
    //error during opening of the file
    if(errorFlag!=ZIP_ER_OK) {
        /*char* errorStr = new char[256];
        zip_error_to_str(errorStr, 255, errorFlag, errno);
        errorStr[255] = '\0';
        cout << "Error: " << errorStr << endl;*/
        
        zipHandle = NULL;
        return false;
    }
    
    if (zipHandle!=NULL) {
        if (isEncrypted()) {
            int result = zip_set_default_password(zipHandle, password.c_str());
            if (result!=0) { 
                close();
                return false;
            }
        }
        
        mode = om;
        return true;
    }
    
    return false;
}

int ZipArchive::close(void) {
    if (isOpen()) {
        int result = zip_close(zipHandle);
        zipHandle = NULL;
        mode = NOT_OPEN;
        
        if(result!=0) { return result; }
    }
    
    return LIBZIPPP_OK;
}

void ZipArchive::discard(void) {
    if (isOpen()) {
        zip_discard(zipHandle);
        zipHandle = NULL;
        mode = NOT_OPEN;
    }
}

bool ZipArchive::unlink(void) {
    if (isOpen()) { discard(); }
    int result = remove(path.c_str());
    return result==0;
}

string ZipArchive::getComment(State state) const {
    if (!isOpen()) { return string(); }
    
    int flag = ZIP_FL_ENC_GUESS;
    if (state==ORIGINAL) { flag = flag | ZIP_FL_UNCHANGED; }
    
    int length = 0;
    const char* comment = zip_get_archive_comment(zipHandle, &length, flag);
    if (comment==NULL) { return string(); }
    return string(comment, length);
}

bool ZipArchive::setComment(const string& comment) const {
    if (!isOpen()) { return false; }
    if (mode==READ_ONLY) { return false; }
    
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
    if (mode==READ_ONLY) { return false; }
    
    libzippp_uint16 compMode = value ? ZIP_CM_DEFLATE : ZIP_CM_STORE;
    return zip_set_file_compression(zipHandle, entry.index, compMode, 0);
}

libzippp_int64 ZipArchive::getNbEntries(State state) const {
    if (!isOpen()) { return LIBZIPPP_ERROR_NOT_OPEN; }
    
    int flag = state==ORIGINAL ? ZIP_FL_UNCHANGED : 0;
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
    int flag = state==ORIGINAL ? ZIP_FL_UNCHANGED : 0;
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
    
    int flags = ZIP_FL_ENC_GUESS;
    if (excludeDirectories) { flags = flags | ZIP_FL_NODIR; }
    if (!caseSensitive) { flags = flags | ZIP_FL_NOCASE; }
    if (state==ORIGINAL) { flags = flags | ZIP_FL_UNCHANGED; }
    
    libzippp_int64 index = zip_name_locate(zipHandle, name.c_str(), flags);
    return index>=0;
}

ZipEntry ZipArchive::getEntry(const string& name, bool excludeDirectories, bool caseSensitive, State state) const {
    if (isOpen()) {
        int flags = ZIP_FL_ENC_GUESS;
        if (excludeDirectories) { flags = flags | ZIP_FL_NODIR; }
        if (!caseSensitive) { flags = flags | ZIP_FL_NOCASE; }
        if (state==ORIGINAL) { flags = flags | ZIP_FL_UNCHANGED; }

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
        int flag = state==ORIGINAL ? ZIP_FL_UNCHANGED : 0;
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
    
    int flag = ZIP_FL_ENC_GUESS;
    if (state==ORIGINAL) { flag = flag | ZIP_FL_UNCHANGED; }
    
    uint clen;
    const char* com = zip_file_get_comment(zipHandle, entry.getIndex(), &clen, flag);
    string comment = com==NULL ? string() : string(com, clen);
    return comment;
}

bool ZipArchive::setEntryComment(const ZipEntry& entry, const string& comment) const {
    if (!isOpen()) { return false; }
    if (entry.zipFile!=this) { return false; }
    
    bool result = zip_file_set_comment(zipHandle, entry.getIndex(), comment.c_str(), comment.size(), ZIP_FL_ENC_GUESS);
    return result==0;
}

void* ZipArchive::readEntry(const ZipEntry& zipEntry, bool asText, State state, libzippp_uint64 size) const {
    if (!isOpen()) { return NULL; }
    if (zipEntry.zipFile!=this) { return NULL; }
    
    int flag = state==ORIGINAL ? ZIP_FL_UNCHANGED : 0;
    struct zip_file* zipFile = zip_fopen_index(zipHandle, zipEntry.getIndex(), flag);
    if (zipFile) {
        libzippp_uint64 maxSize = zipEntry.getSize();
        libzippp_uint64 uisize = size==0 || size>maxSize ? maxSize : size;
        
        char* data = new char[uisize+(asText ? 1 : 0)];
        if(!data) { //allocation error
            zip_fclose(zipFile);
            return NULL; 
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
    
    return NULL;
}

void* ZipArchive::readEntry(const string& zipEntry, bool asText, State state, libzippp_uint64 size) const {
    ZipEntry entry = getEntry(zipEntry);
    if (entry.isNull()) { return NULL; }
    return readEntry(entry, asText, state, size);
}

int ZipArchive::deleteEntry(const ZipEntry& entry) const {
    if (!isOpen()) { return LIBZIPPP_ERROR_NOT_OPEN; }
    if (entry.zipFile!=this) { return LIBZIPPP_ERROR_INVALID_ENTRY; }
    if (mode==READ_ONLY) { return LIBZIPPP_ERROR_NOT_ALLOWED; } //deletion not allowed
    
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
    if (mode==READ_ONLY) { return LIBZIPPP_ERROR_NOT_ALLOWED; } //renaming not allowed
    if (newName.length()==0) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    if (newName==entry.getName()) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    
    if (entry.isFile()) {
        if (IS_DIRECTORY(newName)) { return LIBZIPPP_ERROR_INVALID_PARAMETER; } //invalid new name
        
        int lastSlash = newName.rfind(DIRECTORY_SEPARATOR);
        if (lastSlash!=1) { 
            bool dadded = addEntry(newName.substr(0, lastSlash+1)); 
            if (!dadded) { return LIBZIPPP_ERROR_UNKNOWN; } //the hierarchy hasn't been created
        }
        
        int result = zip_file_rename(zipHandle, entry.getIndex(), newName.c_str(), ZIP_FL_ENC_GUESS);
        if (result==0) { return 1; }
        return LIBZIPPP_ERROR_UNKNOWN; //renaming was not possible (entry already exists ?)
    } else {
        if (!IS_DIRECTORY(newName)) { return LIBZIPPP_ERROR_INVALID_PARAMETER; } //invalid new name
        
        int parentSlash = newName.rfind(DIRECTORY_SEPARATOR, newName.length()-2);
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
    if (mode==READ_ONLY) { return false; } //adding not allowed
    if (IS_DIRECTORY(entryName)) { return false; }
    
    int lastSlash = entryName.rfind(DIRECTORY_SEPARATOR);
    if (lastSlash!=-1) { //creates the needed parent directories
        string dirEntry = entryName.substr(0, lastSlash+1);
        bool dadded = addEntry(dirEntry);
        if (!dadded) { return false; }
    }
    
    //retrieves the length of the file
    //http://stackoverflow.com/questions/5840148/how-can-i-get-a-files-size-in-c
    const char* filepath = file.c_str();
    ifstream in(filepath, ifstream::in | ifstream::binary);
    in.seekg(0, ifstream::end);
    streampos end = in.tellg();
    
    zip_source* source = zip_source_file(zipHandle, filepath, 0, end);
    if (source!=NULL) {
        libzippp_int64 result = zip_file_add(zipHandle, entryName.c_str(), source, ZIP_FL_OVERWRITE);
        if (result>=0) { return true; } 
        else { zip_source_free(source); } //unable to add the file
    } else {
        //unable to create the zip_source
    }
    return false;
}

bool ZipArchive::addData(const string& entryName, const void* data, libzippp_uint64 length, bool freeData) const {
    if (!isOpen()) { return false; }
    if (mode==READ_ONLY) { return false; } //adding not allowed
    if (IS_DIRECTORY(entryName)) { return false; }
    
    int lastSlash = entryName.rfind(DIRECTORY_SEPARATOR);
    if (lastSlash!=-1) { //creates the needed parent directories
        string dirEntry = entryName.substr(0, lastSlash+1);
        bool dadded = addEntry(dirEntry);
        if (!dadded) { return false; }
    }
    
    zip_source* source = zip_source_buffer(zipHandle, data, length, freeData);
    if (source!=NULL) {
        libzippp_int64 result = zip_file_add(zipHandle, entryName.c_str(), source, ZIP_FL_OVERWRITE);
        if (result>=0) { return true; } 
        else { zip_source_free(source); } //unable to add the file
    } else {
        //unable to create the zip_source
    }
    return false;
}

bool ZipArchive::addEntry(const string& entryName) const {
    if (!isOpen()) { return false; }
    if (mode==READ_ONLY) { return false; } //adding not allowed
    if (!IS_DIRECTORY(entryName)) { return false; }
    
    int nextSlash = entryName.find(DIRECTORY_SEPARATOR);
    while (nextSlash!=-1) {
        string pathToCreate = entryName.substr(0, nextSlash+1);
        if (!hasEntry(pathToCreate)) {
            libzippp_int64 result = zip_dir_add(zipHandle, pathToCreate.c_str(), ZIP_FL_ENC_GUESS);
            if (result==-1) { return false; }
        }
        nextSlash = entryName.find(DIRECTORY_SEPARATOR, nextSlash+1);
    }
    
    return true;
}

int ZipArchive::readEntry(const ZipEntry& zipEntry, std::ofstream& ofOutput, State state, libzippp_uint64 chunksize) const {
    if (!ofOutput.is_open()) { return LIBZIPPP_ERROR_INVALID_PARAMETER; }
    if (!isOpen()) { return LIBZIPPP_ERROR_NOT_OPEN; }
    if (zipEntry.zipFile!=this) { return LIBZIPPP_ERROR_INVALID_ENTRY; }
    
    int iRes = LIBZIPPP_OK;
    int flag = state==ORIGINAL ? ZIP_FL_UNCHANGED : 0;
    struct zip_file* zipFile = zip_fopen_index(zipHandle, zipEntry.getIndex(), flag);
    if (zipFile) {
        libzippp_uint64 maxSize = zipEntry.getSize();
        if (!chunksize) { chunksize = DEFAULT_CHUNK_SIZE; } // use the default chunk size (512K) if not specified by the user
    
        if (maxSize<chunksize) {
            char* data = new char[maxSize];
            if (data!=NULL) {
                libzippp_int64 result = zip_fread(zipFile, data, maxSize);
                if (result>0) {
                    if (result != static_cast<libzippp_int64>(maxSize)) {
                        iRes = LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE;
                    } else {
                        ofOutput.write(data, maxSize);
                        if (!ofOutput) { iRes = LIBZIPPP_ERROR_OWRITE_FAILURE; }
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
            char* data = new char[chunksize];
            if (data!=NULL) {
                int nbChunks = maxSize/chunksize;
                for (int uiChunk=0 ; uiChunk<nbChunks ; ++uiChunk) {
                    result = zip_fread(zipFile, data, chunksize);
                    if (result>0) {
                        if (result!=static_cast<libzippp_int64>(chunksize)) {
                            iRes = LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE;
                            break;
                        } else {
                            ofOutput.write(data, chunksize);
                            if (!ofOutput) {
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
                char* data = new char[leftOver];
                if (data!=NULL) {
                    result = zip_fread(zipFile, data, leftOver);
                    if (result>0) {
                        if (result!=static_cast<libzippp_int64>(leftOver)) {
                            iRes = LIBZIPPP_ERROR_OWRITE_INDEX_FAILURE;
                        } else {
                            ofOutput.write(data, leftOver);
                            if (!ofOutput) {
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
