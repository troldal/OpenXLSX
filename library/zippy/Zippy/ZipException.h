//
// Created by Troldal on 2019-03-10.
//

#ifndef MINIZ_ZIPEXCEPTION_H
#define MINIZ_ZIPEXCEPTION_H

#include <stdexcept>

namespace Zippy
{
    class ZipException : public std::runtime_error
    {
    public:
        inline explicit ZipException(const std::string& err)
                : runtime_error(err) {
        }

        inline ~ZipException() override = default;
    };

    class ZipExceptionUndefined : public ZipException
    {
    public:
        inline explicit ZipExceptionUndefined(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionUndefined() override = default;
    };

    class ZipExceptionTooManyFiles : public ZipException
    {
    public:
        inline explicit ZipExceptionTooManyFiles(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionTooManyFiles() override = default;
    };

    class ZipExceptionFileTooLarge : public ZipException
    {
    public:
        inline explicit ZipExceptionFileTooLarge(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileTooLarge() override = default;
    };

    class ZipExceptionUnsupportedMethod : public ZipException
    {
    public:
        inline explicit ZipExceptionUnsupportedMethod(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionUnsupportedMethod() override = default;
    };

    class ZipExceptionUnsupportedEncryption : public ZipException
    {
    public:
        inline explicit ZipExceptionUnsupportedEncryption(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionUnsupportedEncryption() override = default;
    };

    class ZipExceptionUnsupportedFeature : public ZipException
    {
    public:
        inline explicit ZipExceptionUnsupportedFeature(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionUnsupportedFeature() override = default;
    };

    class ZipExceptionFailedFindingCentralDir : public ZipException
    {
    public:
        inline explicit ZipExceptionFailedFindingCentralDir(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFailedFindingCentralDir() override = default;
    };

    class ZipExceptionNotAnArchive : public ZipException
    {
    public:
        inline explicit ZipExceptionNotAnArchive(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionNotAnArchive() override = default;
    };

    class ZipExceptionInvalidHeader : public ZipException
    {
    public:
        inline explicit ZipExceptionInvalidHeader(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionInvalidHeader() override = default;
    };

    class ZipExceptionMultidiskUnsupported : public ZipException
    {
    public:
        inline explicit ZipExceptionMultidiskUnsupported(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionMultidiskUnsupported() override = default;
    };

    class ZipExceptionDecompressionFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionDecompressionFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionDecompressionFailed() override = default;
    };

    class ZipExceptionCompressionFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionCompressionFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionCompressionFailed() override = default;
    };

    class ZipExceptionUnexpectedDecompSize : public ZipException
    {
    public:
        inline explicit ZipExceptionUnexpectedDecompSize(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionUnexpectedDecompSize() override = default;
    };

    class ZipExceptionCrcCheckFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionCrcCheckFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionCrcCheckFailed() override = default;
    };

    class ZipExceptionUnsupportedCDirSize : public ZipException
    {
    public:
        inline explicit ZipExceptionUnsupportedCDirSize(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionUnsupportedCDirSize() override = default;
    };

    class ZipExceptionAllocFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionAllocFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionAllocFailed() override = default;
    };

    class ZipExceptionFileOpenFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionFileOpenFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileOpenFailed() override = default;
    };

    class ZipExceptionFileCreateFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionFileCreateFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileCreateFailed() override = default;
    };

    class ZipExceptionFileWriteFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionFileWriteFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileWriteFailed() override = default;
    };

    class ZipExceptionFileReadFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionFileReadFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileReadFailed() override = default;
    };

    class ZipExceptionFileCloseFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionFileCloseFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileCloseFailed() override = default;
    };

    class ZipExceptionFileSeekFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionFileSeekFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileSeekFailed() override = default;
    };

    class ZipExceptionFileStatFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionFileStatFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileStatFailed() override = default;
    };

    class ZipExceptionInvalidParameter : public ZipException
    {
    public:
        inline explicit ZipExceptionInvalidParameter(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionInvalidParameter() override = default;
    };

    class ZipExceptionInvalidFilename : public ZipException
    {
    public:
        inline explicit ZipExceptionInvalidFilename(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionInvalidFilename() override = default;
    };

    class ZipExceptionBufferTooSmall : public ZipException
    {
    public:
        inline explicit ZipExceptionBufferTooSmall(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionBufferTooSmall() override = default;
    };

    class ZipExceptionInternalError : public ZipException
    {
    public:
        inline explicit ZipExceptionInternalError(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionInternalError() override = default;
    };

    class ZipExceptionFileNotFound : public ZipException
    {
    public:
        inline explicit ZipExceptionFileNotFound(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionFileNotFound() override = default;
    };

    class ZipExceptionArchiveTooLarge : public ZipException
    {
    public:
        inline explicit ZipExceptionArchiveTooLarge(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionArchiveTooLarge() override = default;
    };

    class ZipExceptionValidationFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionValidationFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionValidationFailed() override = default;
    };

    class ZipExceptionWriteCallbackFailed : public ZipException
    {
    public:
        inline explicit ZipExceptionWriteCallbackFailed(const std::string& err)
                : ZipException(err) {
        }

        inline ~ZipExceptionWriteCallbackFailed() override = default;
    };

} // namespace Zippy


#endif //MINIZ_ZIPEXCEPTION_H
