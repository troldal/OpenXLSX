//
// Created by Troldal on 19/09/16.
//

#include "XLArchive.h"

#include <fstream>

using namespace std;
using namespace libzippp;
using namespace OpenXLSX;

/**
 * @details The constructor opens the zip file with given path, using the libzippp:ZipArchive class.
 * The zip file is opened in WRITE mode.
 */
XLArchive::XLArchive(const std::string &zipFileName)
    : m_archive(make_unique<ZipArchive>(zipFileName)),
      m_fileName(zipFileName)
{
    m_archive->open(ZipArchive::WRITE);
}

/**
 * @details The destructor closes the underlying ZipArchive object.
 */
XLArchive::~XLArchive()
{
    m_archive->close();
}

/**
 * @details This method checks if the file with the given filename exists in the arvhive, by calling the 'hasEntry'
 * method in the ZipArchive object.
 */
bool XLArchive::FileExists(const std::string &fileName) const
{
    return m_archive->hasEntry(fileName);
}

/**
 * @details This method retrieves the file with the given filename from the archive, as a std::string. This is done by
 * means of the 'getEntry' method in the ZipArchive object.
 */
std::string XLArchive::GetFile(const std::string &fileName) const
{
    return m_archive->getEntry(fileName).readAsText();
}

/**
 * @details This method adds a file with the given filename to the archive, as a text file. This is done by means
 * of the 'addData' method in the ZipArchive object.
 */
void XLArchive::AddFile(const std::string &fileName,
                        const std::string &text)
{
    m_archive->addData(fileName, text.c_str(), text.size());
}

/**
 * @details This method deletes the file with the given filename from the archive, via the 'deleteEntry' method
 * in the ZipArchive object.
 */
void XLArchive::DeleteFile(const std::string &fileName)
{
    m_archive->deleteEntry(fileName);
}

/**
 * @details This method copies the current archive to a new file and opens the new archive.
 * This is done in the following manner:
 * - Close the archive.
 * - Open binary streams for source and destination files.
 * - Copy the bits from the source to the destination.
 * - Open the copy of the archive.
 */
void XLArchive::CopyArchive(const std::string &copyName)
{
    m_archive->close();
    ifstream src(m_fileName, ios::binary);
    ofstream dst(copyName, ios::binary);

    dst << src.rdbuf();

    dst.close();
    src.close();

    m_fileName = copyName;
    m_archive = make_unique<ZipArchive>(m_fileName);
    m_archive->open(ZipArchive::WRITE);
}
