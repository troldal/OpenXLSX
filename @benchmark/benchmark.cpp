#include <iostream>
#include <chrono>
#include <iomanip>
#include "table_printer.h"
#include "../@source/OpenXLSX.h"

using namespace std;
using namespace OpenXLSX;
using bprinter::TablePrinter;

template<typename T>
unsigned long WriteTest(T value, unsigned long rows, unsigned long columns, int repetitions);

int main()
{
    WriteTest("Hello, OpenXLSX!", 1000000, 1, 10);
    WriteTest("Hello, OpenXLSX!", 100000, 10, 10);
    WriteTest("Hello, OpenXLSX!", 10000, 100, 10);
    WriteTest("Hello, OpenXLSX!", 1000, 1000, 10);
    WriteTest("Hello, OpenXLSX!", 100, 10000, 10);
    return 0;
}

template<typename T>
unsigned long WriteTest(T value, unsigned long rows, unsigned long columns, int repetitions)
{
    TablePrinter tp(&std::cout);
    tp.AddColumn("Run #", 10);
    tp.AddColumn("Opening (ms)", 20);
    tp.AddColumn("Writing (ms)", 20);
    tp.AddColumn("Writing (cells/s)", 20);
    tp.AddColumn("Saving (ms)", 20);

    tp.PrintTitle("Writing " + to_string(rows) + " x " + to_string(columns) + " cells");
    tp.PrintHeader();

    unsigned long TotalOpeningTime = 0;
    unsigned long TotalWritingTime = 0;
    unsigned long TotalSavingTime = 0;
    for (int i = 1; i <= repetitions; i++) {

        // Create new document
        auto startOpen = chrono::steady_clock::now();
        OpenXLSX::XLDocument doc;
        doc.CreateDocument("Benchmark.xlsx");
        auto endOpen = chrono::steady_clock::now();

        // Add data to document
        auto startWrite = chrono::steady_clock::now();
        auto wks = doc.Workbook()->Worksheet("Sheet1");
        wks->Cell(rows, columns)->Value()->Set(1);
        auto arange = wks->Range();
        for (auto &iter : arange) {
            iter.Value()->Set(value);
        }
        auto endWrite = chrono::steady_clock::now();

        // Save document
        auto startSave = chrono::steady_clock::now();
        doc.SaveDocument();
        doc.CloseDocument();
        auto endSave = chrono::steady_clock::now();

        TotalOpeningTime += std::chrono::duration_cast<std::chrono::milliseconds>(endOpen - startOpen).count();
        TotalWritingTime += std::chrono::duration_cast<std::chrono::milliseconds>(endWrite - startWrite).count();
        TotalSavingTime  += std::chrono::duration_cast<std::chrono::milliseconds>(endSave - startSave).count();

        tp << ("#" + to_string(i))
            << std::chrono::duration_cast<std::chrono::milliseconds>(endOpen - startOpen).count()
            << std::chrono::duration_cast<std::chrono::milliseconds>(endWrite - startWrite).count()
            << static_cast<unsigned long>(static_cast<double>(rows*columns)/(std::chrono::duration_cast<std::chrono::milliseconds>(endWrite - startWrite).count())*1000)
            << std::chrono::duration_cast<std::chrono::milliseconds>(endSave - startSave).count();

    }

    tp.PrintFooter();
    tp << "Average"
        << TotalOpeningTime/repetitions
        << TotalWritingTime/repetitions
        << static_cast<unsigned long>(static_cast<double>(rows*columns*repetitions)/(TotalWritingTime/1000))
        << TotalSavingTime/repetitions;
    tp.PrintFooter();

    cout << endl;

    return 0;
}

void myTest()
{
    XLDocument doc;
    doc.CreateDocument("./MyTest.xlsx");
    auto wks = doc.Workbook()->Worksheet("Sheet1");

    *wks->Cell("A1")->Value() = 3.14159;
    *wks->Cell("B1")->Value() = 42;
    *wks->Cell("C1")->Value() = "Hello OpenXLSX!";
    *wks->Cell("D1")->Value() = true;

    auto A1 = wks->Cell("A1")->Value()->Get<long double>();
    auto B1 = wks->Cell("B1")->Value()->Get<unsigned int>();
    auto C1 = wks->Cell("C1")->Value()->Get<std::string_view>();
    auto D1 = wks->Cell("D1")->Value()->Get<bool>();

    cout << "Cell A1: " << A1 << endl;
    cout << "Cell B1: " << B1 << endl;
    cout << "Cell C1: " << C1 << endl;
    cout << "Cell D1: " << D1 << endl;

    doc.SaveDocument();
}



void openLarge() {

    auto start = chrono::steady_clock::now();

    OpenXLSX::XLDocument doc;
    doc.OpenDocument("./Large.xlsx");
    auto wks = doc.Workbook()->Worksheet("Sheet1");
    wks->Export("./Profiles.csv");

    /*
    for (int i = 1; i <= 1000; ++i) {
        cout << setw(8) << i << ": ";
        cout << setw(10) << wks->Cell(i, 1)->Value()->AsString() << " ";
        cout << setw(38) << wks->Cell(i, 2)->Value()->AsString() << " ";
        cout << setw(38) << wks->Cell(i, 3)->Value()->AsString() << " ";
        cout << setw(38) << wks->Cell(i, 4)->Value()->AsString() << " ";
        cout << setw(38) << wks->Cell(i, 5)->Value()->AsString() << endl;
    } */

    doc.CloseDocument();

    auto end = chrono::steady_clock::now();
    cout << "Time to open file and export ~500k rows: " << std::chrono::duration_cast<std::chrono::milliseconds>(end - start).count() << endl;

}


