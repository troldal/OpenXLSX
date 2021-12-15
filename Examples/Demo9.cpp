#include <OpenXLSX.hpp>
#include <filesystem>
#include <iostream>

#ifdef WIN32
#include <stdio.h>
#include <windows.h>
#endif

using namespace std;
using namespace OpenXLSX;

void printType(XLNumberFormat fmt)
{
#ifdef WIN32
    SetConsoleOutputCP(CP_UTF8);
#endif

    switch (fmt.type()) {
        case XLNumberFormat::XLNumberType::kCurrency:
            cout << "\nkCurrency " + fmt.currencySumbol();
            break;
        case XLNumberFormat::XLNumberType::kDate:
            cout << "\nkDate ";
            break;
        case XLNumberFormat::XLNumberType::kPercent:
            cout << "\nkPercent ";
            break;
        case XLNumberFormat::XLNumberType::kUnkown:
            cout << "\nkUnkown ";
            break;
    }
}

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #09: Basic Usage\n";
    cout << "********************************************************************************\n";

    auto&                 current_path = std::filesystem::current_path();
    auto&                 parent_path  = current_path.parent_path();
    std::filesystem::path foundPath;

    while (parent_path.string().size() > 3) {
        if (parent_path.filename().string().compare("OpenXLSX") == 0) foundPath = parent_path;
        parent_path = parent_path.parent_path();
    }
    foundPath += "\\Tests\\NumberType.xlsx";
    if (!std::filesystem::exists(foundPath)) {
        cout << "File not found";
        return 0;
    }

    XLDocument doc;
    doc.open(foundPath.string());
    auto wks = doc.workbook().worksheet("Sheet1");

    {
        const auto     cell = wks.cell("A1");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }
    {
        const auto     cell = wks.cell("A2");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }
    {
        const auto     cell = wks.cell("A3");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }
    {
        const auto     cell = wks.cell("A4");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }
    {
        const auto     cell = wks.cell("A5");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }
    {
        const auto     cell = wks.cell("A6");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }
    {
        const auto     cell = wks.cell("A7");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }
    {
        const auto     cell = wks.cell("A8");
        XLNumberFormat fmt(cell.style());
        printType(fmt);
    }

    doc.close();

    return 0;
}