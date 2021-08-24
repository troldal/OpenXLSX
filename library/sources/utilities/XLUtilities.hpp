//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLUTILITIES_HPP
#define OPENXLSX_XLUTILITIES_HPP

#include <fstream>
#include <pugixml.hpp>

#include "XLCellReference.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    /**
     * @details
     */
    inline XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        // ===== If the requested node is beyond the current max node, append a new node to the end.
        auto result = XMLNode();
        if (!sheetDataNode.last_child() || rowNumber > sheetDataNode.last_child().attribute("r").as_ullong()) {
            result = sheetDataNode.append_child("row");

            result.append_attribute("r") = rowNumber;
            //            result.append_attribute("x14ac:dyDescent") = "0.2";
            //            result.append_attribute("spans")           = "1:1";
        }

        // ===== If the requested node is closest to the end, start from the end and search backwards
        else if (sheetDataNode.last_child().attribute("r").as_ullong() - rowNumber < rowNumber) {
            result = sheetDataNode.last_child();
            while (result.attribute("r").as_ullong() > rowNumber) result = result.previous_sibling();
            if (result.attribute("r").as_ullong() < rowNumber) {
                result = sheetDataNode.insert_child_after("row", result);

                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        // ===== Otherwise, start from the beginning
        else {
            result = sheetDataNode.first_child();
            while (result.attribute("r").as_ullong() < rowNumber) result = result.next_sibling();
            if (result.attribute("r").as_ullong() > rowNumber) {
                result = sheetDataNode.insert_child_before("row", result);

                result.append_attribute("r") = rowNumber;
                //                result.append_attribute("x14ac:dyDescent") = "0.2";
                //                result.append_attribute("spans")           = "1:1";
            }
        }

        return result;
    }

    /**
     * @brief Retrieve the xml node representing the cell at the given row and column. If the node doesn't
     * exist, it will be created.
     * @param rowNode The row node under which to find the cell.
     * @param columnNumber The column at which to find the cell.
     * @return The xml node representing the requested cell.
     */
    inline XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber)
    {
        auto cellNode = XMLNode();
        auto cellRef  = XLCellReference(rowNode.attribute("r").as_uint(), columnNumber);

        // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
        if (rowNode.last_child().empty() || XLCellReference(rowNode.last_child().attribute("r").value()).column() < columnNumber) {
            rowNode.append_child("c").append_attribute("r").set_value(cellRef.address().c_str());
            cellNode = rowNode.last_child();
        }
        // ===== If the requested node is closest to the end, start from the end and search backwards...
        else if (XLCellReference(rowNode.last_child().attribute("r").value()).column() - columnNumber < columnNumber) {
            cellNode = rowNode.last_child();
            while (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) cellNode = cellNode.previous_sibling();
            if (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber) {
                cellNode = rowNode.insert_child_after("c", cellNode);
                cellNode.append_attribute("r").set_value(cellRef.address().c_str());
            }
        }
        // ===== Otherwise, start from the beginning
        else {
            cellNode = rowNode.first_child();
            while (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber) cellNode = cellNode.next_sibling();
            if (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber) {
                cellNode = rowNode.insert_child_before("c", cellNode);
                cellNode.append_attribute("r").set_value(cellRef.address().c_str());
            }
        }
        return cellNode;
    }

    /**
     * @brief This function returns a std::ofstream with the given filename (must be UTF-8). However, implementation details depends on the
     * platform. If the platform is Windows, the filename has to be converted to a wide string (UTF-16 on windows); otherwise. the filename
     * string is used as is.
     * @param fileName A UTF-8 string with the filename
     * @return A std::ofstream
     * @note The UTF8->UTF16 conversion was implemented by WCIofQMandRA.
     * @todo Consider replacing the conversion algorithm by 3rd party library, such as miniutf.
     */
    inline std::ofstream getOutFileStream(const std::string& fileName)
    {
#ifdef _WIN32
        auto utf8_16 = [](const char* u8_str, wchar_t* u16_str)    // NOLINT
        {
            const unsigned char* pu      = (const unsigned char*)u8_str;
            wchar_t *            it      = u16_str, tmp;
            const wchar_t        invalid = 0xFFFDu;
            while (*pu) {
                if (*pu < 0x80u)
                    *it++ = (wchar_t)(*pu++);
                else if (*pu < 0xC2u)
                    *it++ = invalid, ++pu;
                else if (*pu < 0xE0u) {
                    /* uint32 is faster than uint16 */
                    uint32_t ch = *pu++ & 0x1Fu;
                    tmp         = *pu;
                    if (0x80u <= tmp && tmp < 0xC0u)
                        ++pu, *it++ = (wchar_t)(ch << 6 | tmp & 0x3Fu);
                    else
                        *it++ = invalid;
                }
                else if (*pu < 0xF0u) {
                    uint32_t ch = *pu++ & 0x0Fu;
                    tmp         = *pu;
                    if (0x80u <= tmp && tmp < 0xC0u) {
                        ++pu, ch = ch << 6 | tmp & 0x3Fu;
                        tmp = *pu;
                        if (0x80u <= tmp && tmp <= 0xC0u) {
                            ++pu, ch = ch << 6 | tmp & 0x3Fu;
                            /* isolated surrogate pair */
                            if (0xD800u <= ch && ch < 0xE000u)
                                *it++ = invalid;
                            else
                                *it++ = (wchar_t)ch;
                        }
                        else
                            *it++ = invalid;
                    }
                    else
                        *it++ = invalid;
                }
                else if (*pu < 0xF8u) {
                    uint32_t ch = *pu & 0x0Fu;
                    tmp         = *pu;
                    if (0x80u <= tmp && tmp < 0xC0u) {
                        ++pu, ch = ch << 6 | tmp & 0x3Fu;
                        tmp = *pu;
                        if (0x80u <= tmp && tmp < 0xC0u) {
                            ++pu, ch = ch << 6 | tmp & 0x3Fu;
                            tmp = *pu;
                            if (0x80u <= tmp && tmp <= 0xC0u) {
                                ++pu, ch = ch << 6 | tmp & 0x3Fu;
                                if (ch >= 0x110000u)
                                    *it++ = invalid;
                                else {
                                    ch -= 0x10000u;
                                    *it++ = (wchar_t)(0xD800u | ch >> 10);
                                    *it++ = (wchar_t)(0xDC00u | ch & 0x3FFu);
                                }
                            }
                            else
                                *it++ = invalid;
                        }
                        else
                            *it++ = invalid;
                    }
                    else
                        *it++ = invalid;
                }
                else
                    *it++ = invalid, ++pu;
            }
            *it = L'\0';
        };

        std::ofstream outfile;
        if (fileName.length() < 256u)    // very likely
        {
            wchar_t fileName_w[256];
            utf8_16(fileName.c_str(), fileName_w);
            outfile.open(fileName_w, std::ios::binary);
        }
        else {
            wchar_t* fileName_w = new wchar_t[fileName.length() + 1];
            utf8_16(fileName.c_str(), fileName_w);
            outfile.open(fileName_w, std::ios::binary);
            delete[] fileName_w;
        }
#else
        std::ofstream outfile(fileName, std::ios::binary);
#endif
        return outfile;
    }

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLUTILITIES_HPP
