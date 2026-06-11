#include "CxxOpenXLSXTestSupport.hpp"

#include <OpenXLSX.hpp>

#include <exception>
#include <string>

std::string cxx_readCellString(
    const std::string& filePath,
    const std::string& cellReference
)
{
    try {
        OpenXLSX::XLDocument document;
        document.open(filePath);

        auto worksheet = document.workbook().worksheet(1);
        return worksheet.cell(cellReference).value().getString();
    } catch (const std::exception& error) {
        return std::string("ERROR: ") + error.what();
    } catch (...) {
        return "ERROR: unknown C++ exception";
    }
}

std::string cxx_readExtList(const std::string& filePath)
{
    try {
        OpenXLSX::XLDocument document;
        document.open(filePath);

        auto worksheet = document.workbook().worksheet(1);
        return worksheet.extList();
    } catch (const std::exception& error) {
        return std::string("ERROR: ") + error.what();
    } catch (...) {
        return "ERROR: unknown C++ exception";
    }
}
