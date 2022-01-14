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


void testNumbers(const XLWorksheet& wks)
{
    // check if the value is a number before using XLNumberFormat

     cout << "\n";
    auto doNumFmt = [&](const XLCell& cell) {
        const auto& val = cell.value();
        if (XLValueType::Integer == val.type() || XLValueType::Float == val.type()) {
            const auto fmt = cell.style().numberFormat();
            printType(fmt);
        }
        else
            cout << "\nNot a number";
    };

    doNumFmt(wks.cell("A1"));
    doNumFmt(wks.cell("A2"));
    doNumFmt(wks.cell("A3"));
    doNumFmt(wks.cell("A4"));
    doNumFmt(wks.cell("A5"));
    doNumFmt(wks.cell("A6"));
    doNumFmt(wks.cell("A7"));
    doNumFmt(wks.cell("A8"));
    doNumFmt(wks.cell("A9"));
    cout << "\n";
}

void testFonts(const XLWorksheet& wks)
{
    cout << "\n";
    auto doFonstTest = [&](const XLCell& cell) {
        const auto& val   = cell.value();
        const auto& style = cell.style();
        const auto& font  = style.font();
        cout << "value= " << val.get<std::string>() << ", name= " << font.name() << ", size= " << font.size()
             << ", color= " << font.color().hex() << ", isBold " << font.bold() << ", isUnderline " << (font.underline() == 2)
             << ", double underline " << (font.underline() == -4119) << ", isStrike " << font.strikethrough() << ", isItalic "
             << font.italic() << '\n';
    };

    doFonstTest(wks.cell("A1"));
    doFonstTest(wks.cell("A2"));
    doFonstTest(wks.cell("A3"));
    doFonstTest(wks.cell("A4"));
    doFonstTest(wks.cell("A5"));
    doFonstTest(wks.cell("A6"));
    cout << "\n";
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

    testNumbers(doc.workbook().worksheet("Sheet1"));
    testFonts(doc.workbook().worksheet("Sheet2"));

    doc.close();

    return 0;
}