/**
 * @file createTemplate.cpp
 * @author Oliver Ofenloch (57812959+ofenloch@users.noreply.github.com)
 * @brief
 * @version 0.1
 * @date 2022-08-03
 *
 * @copyright Copyright (c) 2022
 *
 */

#include <OpenXLSX.hpp>
#include <iostream>

int main(int argc, char* argv[])
{
    int returnValue { 0 };

    // OpenXLSX uses
    //          uint16_t for cell numbers
    // and
    //          uint32_t for row numbers

    for (auto cellName : { "A", "E", "Z", "AA", "AB", "AZ", "BA", "BZ", "CA", "CV", "GR" }) {
        const uint16_t cellNumber = OpenXLSX::XLWorksheet::columnNameToNumber(cellName);
        std::cout << "column name " << cellName << " is column number " << cellNumber << '\n';
        const std::string cellName2 = OpenXLSX::XLWorksheet::columnNumberToName(cellNumber);
        if (cellName2 != cellName) {
            std::cout << "ERROR: column number " << cellNumber << " is NOT column name " << cellName2 << '\n';
        }
    }

    const uint32_t nRows { 15 };
    const uint32_t nRowsOffset { 5 };

    const std::string fileName { "template.xlsx" };
    std::remove(fileName.c_str());

    OpenXLSX::XLDocument doc;
    doc.create(fileName);
    OpenXLSX::XLWorkbook  wbk = doc.workbook();
    OpenXLSX::XLWorksheet wks = wbk.worksheet("Sheet1");

    // just for testing:
    std::vector<uint16_t> numbers;
    std::vector<std::string> names;
    for (uint16_t i = 1; i <= 1024; i++) {
        numbers.push_back(i);
        names.push_back(OpenXLSX::XLWorksheet::columnNumberToName(i));
    }
    wks.row(1).values() = numbers;
    wks.row(2).values() = names;

    for (uint32_t iLogicalRow = 1; iLogicalRow <= nRows; iLogicalRow++) {
        const std::string siLogicalRow { std::to_string(iLogicalRow) };
        const uint32_t    iRealRow = nRowsOffset + iLogicalRow;
        const std::string siRealRow { std::to_string(iRealRow) };

        std::vector<std::string> cellContents;

        // cell A:
        cellContents.push_back(std::string("f(prod_massflow_" + siLogicalRow + ")"));
        // cell B:
        cellContents.push_back(std::string("f(prod_temperature_" + siLogicalRow + ")"));
        // cell C:
        cellContents.push_back(std::string("f(prod_pressure_" + siLogicalRow + ")"));
        // cell D:
        cellContents.push_back(std::string("f(prod_density_" + siLogicalRow + ")"));
        // cell E:
        cellContents.push_back(std::string("f(prod_viscosity_" + siLogicalRow + ")"));
        // cell F:
        cellContents.push_back(std::string("=A" + siRealRow + " / D" + siRealRow));
        // cell G (empty):
        cellContents.push_back(std::string(""));
        // cell H (empty):
        cellContents.push_back(std::string(""));
        // cell I:
        cellContents.push_back(std::string("f(prod_reynolds_" + siLogicalRow + ")"));
        wks.row(iRealRow).values() = cellContents;

        // convert formula string to real formula:
        const std::string cellName   = "F" + siRealRow;
        const std::string sFormula   = wks.cell(cellName).value();
        wks.cell(cellName).formula() = sFormula;
    }

    for (uint32_t iLogicalRow = 1; iLogicalRow <= nRows; iLogicalRow++) {
        const std::string siLogicalRow { std::to_string(iLogicalRow) };
        const uint32_t    iRealRow = nRowsOffset + nRows + nRowsOffset + iLogicalRow;
        const std::string siRealRow { std::to_string(iRealRow) };

        std::vector<std::string> cellContents;

        // cell A:
        cellContents.push_back(std::string("f(vap_massflow_" + siLogicalRow + ")"));
        // cell B:
        cellContents.push_back(std::string("f(vap_temperature_" + siLogicalRow + ")"));
        // cell C:
        cellContents.push_back(std::string("f(vap_pressure_" + siLogicalRow + ")"));
        // cell D:
        cellContents.push_back(std::string("f(vap_density_" + siLogicalRow + ")"));
        // cell E:
        cellContents.push_back(std::string("f(vap_viscosity_" + siLogicalRow + ")"));
        // cell F (empty):
        cellContents.push_back(std::string(""));
        // cell G:
        cellContents.push_back(std::string("=A" + siRealRow + " * B" + siRealRow));
        // cell H (empty):
        cellContents.push_back(std::string(""));
        // cell I:
        cellContents.push_back(std::string("f(vap_reynolds_" + siLogicalRow + ")"));
        wks.row(iRealRow).values() = cellContents;

        // convert formula string to real formula:
        const std::string cellName   = "G" + siRealRow;
        const std::string sFormula   = wks.cell(cellName).value();
        wks.cell(cellName).formula() = sFormula;
    }

    // Saving a large spreadsheet can take a while...
    std::cout << "Saving spreadsheet as " << fileName << " ..." << std::endl;
    doc.save();
    doc.close();
    return returnValue;
}