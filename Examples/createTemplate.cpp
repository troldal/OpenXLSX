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

#include <iostream>

#include <OpenXLSX.hpp>

int main(int argc, char* argv[])
{
    int returnValue { 0 };

    // OpenXLSX uses
    //     uint16_t for cell numbers
    // and
    //     uint32_t for row numbers
    const uint32_t nRows { 15 };
    const uint32_t nRowsOffset { 5 };

    const std::string fileName { "template.xlsx" };
    std::remove(fileName.c_str());

    OpenXLSX::XLDocument doc;
    doc.create(fileName);
    OpenXLSX::XLWorkbook wbk = doc.workbook();
    OpenXLSX::XLWorksheet wks = wbk.worksheet("Sheet1");

    // // just for testing:
    // std::vector<uint16_t> numbers;
    // std::vector<std::string> names;
    // for (uint16_t i = 1; i <= 1024; i++) {
    //     numbers.push_back(i);
    //     names.push_back(OpenXLSX::XLWorksheet::columnNumberToName(i));
    // }
    // wks.row(1).values() = numbers;
    // wks.row(2).values() = names;

    for (uint32_t iLogicalRow = 1; iLogicalRow <= nRows; iLogicalRow++) {
        const std::string siLogicalRow { std::to_string(iLogicalRow) };
        const uint32_t iRealRow = nRowsOffset + iLogicalRow;
        const std::string siRealRow { std::to_string(iRealRow) };
        std::vector<std::string> cellContents;        // cell A:
        cellContents.push_back(std::string("f(prod_name_" + siLogicalRow + ")"));
        // cell B:
        cellContents.push_back(std::string("f(prod_massflow_" + siLogicalRow + ")"));
        // cell C:
        cellContents.push_back(std::string("f(prod_temperature_" + siLogicalRow + ")"));
        // cell D:
        cellContents.push_back(std::string("f(prod_pressure_" + siLogicalRow + ")"));
        // cell E:
        cellContents.push_back(std::string("f(prod_density_" + siLogicalRow + ")"));
        // cell F:
        cellContents.push_back(std::string("f(prod_viscosity_" + siLogicalRow + ")"));
        // cell G:
        cellContents.push_back(std::string("f(prod_volumeflow_" + siLogicalRow + ")"));
        // cell H:
        cellContents.push_back(std::string("f(prod_nominal_diameter" + siLogicalRow + ")"));
        // cell I:
        cellContents.push_back(std::string("f(prod_length" + siLogicalRow + ")"));
        // cell J:
        cellContents.push_back(std::string("f(prod_velocity_" + siLogicalRow + ")"));
        // cell K:
        cellContents.push_back(std::string("f(prod_reynolds_" + siLogicalRow + ")"));
        // cell L:
        cellContents.push_back(std::string("f(prod_roughness_" + siLogicalRow + ")"));
        // cell M:
        // cellContents.push_back(std::string("f(prod_initial_value_" + siLogicalRow + ")"));
        // original sheet uses this:
        cellContents.push_back(std::string("=N" + siRealRow + ""));
        // cell N:
        // cellContents.push_back(std::string("=WENN(K11>2300;1/(2*(LOG(2,51/K11/(M11)^0,5+L11/H11/3,71)))^2;64/K11)"));
        cellContents.push_back(std::string("=WENN(K" + siRealRow + ">2300;1/(2*(LOG(2.51/K" + siRealRow + "/(M" + siRealRow + ")^0.5+L" + siRealRow + "/H" + siRealRow + "/3.71)))^2;64/K" + siRealRow + ")"));

        wks.row(iRealRow).values() = cellContents;

        // Excel does not like the reuslt of this:
        //////////////////////////////////////////
        // // convert formula string to real formula:
        // const std::string cellName = "M" + siRealRow;
        // const std::string sFormula = wks.cell(cellName).value();
        // OpenXLSX::XLFormula xlFormula(sFormula);
        // wks.cell(cellName).value() = "";
        // wks.cell(cellName).formula() = xlFormula;
    }

    for (uint32_t iLogicalRow = 1; iLogicalRow <= nRows; iLogicalRow++) {
        const std::string siLogicalRow { std::to_string(iLogicalRow) };
        const uint32_t iRealRow = nRowsOffset + nRows + nRowsOffset + iLogicalRow;
        const std::string siRealRow { std::to_string(iRealRow) };
        std::vector<std::string> cellContents;
        // cell A:
        cellContents.push_back(std::string("f(vap_name_" + siLogicalRow + ")"));
        // cell B:
        cellContents.push_back(std::string("f(vap_massflow_" + siLogicalRow + ")"));
        // cell C:
        cellContents.push_back(std::string("f(vap_temperature_" + siLogicalRow + ")"));
        // cell D:
        cellContents.push_back(std::string("f(vap_pressure_" + siLogicalRow + ")"));
        // cell E:
        cellContents.push_back(std::string("f(vap_density_" + siLogicalRow + ")"));
        // cell F:
        cellContents.push_back(std::string("f(vap_viscosity_" + siLogicalRow + ")"));
        // cell G:
        cellContents.push_back(std::string("f(vap_volumeflow_" + siLogicalRow + ")"));
        // cell H:
        cellContents.push_back(std::string("f(vap_nominal_diameter_" + siLogicalRow + ")"));
        // cell I:
        cellContents.push_back(std::string("f(vap_length" + siLogicalRow + ")"));
        // cell J:
        cellContents.push_back(std::string("f(vap_velocity_" + siLogicalRow + ")"));
        // cell K:
        cellContents.push_back(std::string("f(vap_reynolds_" + siLogicalRow + ")"));
        // cell L:
        cellContents.push_back(std::string("f(vap_roughness_" + siLogicalRow + ")"));
        // cell M:
        cellContents.push_back(std::string("f(vap_initial_value_" + siLogicalRow + ")"));
        // original sheet uses cellContents.push_back(std::string("=N" + siLogicalRow + ""));
        // cell N:
        // cellContents.push_back(std::string("=WENN(K11>2300;1/(2*(LOG(2,51/K11/(M11)^0,5+L11/H11/3,71)))^2;64/K11)"));
        cellContents.push_back(std::string("=WENN(K" + siRealRow + ">2300;1/(2*(LOG(2.51/K" + siRealRow + "/(M" + siRealRow + ")^0.5+L" + siRealRow + "/H" + siRealRow + "/3.71)))^2;64/K" + siRealRow + ")"));


        wks.row(iRealRow).values() = cellContents;
        // // convert formula string to real formula:
        // const std::string cellName = "G" + siRealRow;
        // const std::string sFormula = wks.cell(cellName).value();
        // wks.cell(cellName).formula() = sFormula;
    }
    // wks.setName("Leitungsliste");

    // Saving a large spreadsheet can take a while...
    std::cout << "Saving XLDocument as " << fileName << " ..." << std::endl;
    doc.save();
    doc.close();
    return returnValue;
}
