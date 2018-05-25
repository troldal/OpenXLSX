#include <iostream>
#include <chrono>
#include <iomanip>
#include <sstream>
#include <fstream>
#include "table_printer.h"
#include "../@source/OpenXLSX.h"

using namespace std;
using namespace OpenXLSX;
using bprinter::TablePrinter;

//----------------------------------------------------------------------------------------------------------------------
// Function declarations
//----------------------------------------------------------------------------------------------------------------------
template<typename T>
unsigned long WriteTest(T value,
                        unsigned long rows,
                        unsigned long columns,
                        int repetitions, const
                        std::string &fileName,
                        std::ostream &destination = std::cout);

//----------------------------------------------------------------------------------------------------------------------
// Main
//----------------------------------------------------------------------------------------------------------------------

int main()
{
    //ofstream ostr("Benchmark.out");

    auto &str = cout;

    str <<
    "          ____                               ____      ___ ____       ____  ____      ___\n"
    "         6MMMMb                              `MM(      )M' `MM'      6MMMMb\\`MM(      )M'\n"
    "        8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'\n"
    "       6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'\n"
    "       MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'\n"
    "       MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd\n"
    "       MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.\n"
    "       MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.\n"
    "       YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.\n"
    "        8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.\n"
    "         YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_\n"
    "                   MM\n"
    "                   MM\n"
    "                  _MM_" << endl << endl;

    str << "**************************************";
    str << " RUNNING BENCHMARKS ";
    str << "**************************************"<< endl << endl;

    //WriteTest("Hello, OpenXLSX!", 1000000, 1, 5, "BenchmarkString.xlsx");
    //WriteTest("Hello, OpenXLSX!", 100000, 10, 5, "BenchmarkString.xlsx");
    //WriteTest("Hello, OpenXLSX!", 10000, 100, 5, "BenchmarkString.xlsx");
    WriteTest("Hello, OpenXLSX! \xF0\x9F\x98\x81", 1000, 1000, 5, "BenchmarkString.xlsx", str);
    //WriteTest("Hello, OpenXLSX!", 100, 10000, 5, "BenchmarkString.xlsx");

    //WriteTest(-42, 1000000, 1, 5, "BenchmarkInteger.xlsx");
    //WriteTest(-42, 100000, 10, 5, "BenchmarkInteger.xlsx");
    //WriteTest(-42, 10000, 100, 5, "BenchmarkInteger.xlsx");
    WriteTest(-42, 1000, 1000, 5, "BenchmarkInteger.xlsx", str);
    //WriteTest(-42, 100, 10000, 5, "BenchmarkInteger.xlsx");

    //WriteTest(3.14159, 1000000, 1, 5, "BenchmarkFloat.xlsx");
    //WriteTest(3.14159, 100000, 10, 5, "BenchmarkFloat.xlsx");
    //WriteTest(3.14159, 10000, 100, 5, "BenchmarkFloat.xlsx");
    WriteTest(3.14159, 1000, 1000, 5, "BenchmarkFloat.xlsx", str);
    //WriteTest(3.14159, 100, 10000, 5, "BenchmarkFloat.xlsx");

    //WriteTest(true, 1000000, 1, 5, "BenchmarkBool.xlsx", str);
    //WriteTest(true, 100000, 10, 5, "BenchmarkBool.xlsx", str);
    //WriteTest(true, 10000, 100, 5, "BenchmarkBool.xlsx", str);
    WriteTest(true, 1000, 1000, 5, "BenchmarkBool.xlsx", str);
    //WriteTest(true, 100, 10000, 5, "BenchmarkBool.xlsx", str);

    //ostr.close();

    return 0;
}

//----------------------------------------------------------------------------------------------------------------------
// Function definitions
//----------------------------------------------------------------------------------------------------------------------

template<typename T>
unsigned long WriteTest(T value,
                        unsigned long rows,
                        unsigned long columns,
                        int repetitions,
                        const std::string &fileName,
                        std::ostream &destination)
{
    TablePrinter tp(&destination);
    tp.AddColumn("Run #", 10);
    tp.AddColumn("Opening (ms)", 20);
    tp.AddColumn("Writing (ms)", 20);
    tp.AddColumn("Writing (cells/s)", 20);
    tp.AddColumn("Saving (ms)", 20);

    stringstream ss;
    ss << "Writing \"" << value << "\" to " << to_string(rows) << " x " << to_string(columns) << " cells";

    tp.PrintTitle(ss.str());
    tp.PrintHeader();

    unsigned long TotalOpeningTime = 0;
    unsigned long TotalWritingTime = 0;
    unsigned long TotalSavingTime = 0;
    for (int i = 1; i <= repetitions; i++) {

        // Create new document
        auto startOpen = chrono::steady_clock::now();
        OpenXLSX::XLDocument doc;
        doc.CreateDocument(fileName);
        auto endOpen = chrono::steady_clock::now();

        // Add data to document
        auto startWrite = chrono::steady_clock::now();
        auto wks = doc.Workbook()->Worksheet("Sheet1");
        auto arange = wks->Range(XLCellReference("A1"), XLCellReference(rows, columns));
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

        double dims = static_cast<double>(rows*columns);
        double time = static_cast<double>(std::chrono::duration_cast<std::chrono::milliseconds>(endWrite - startWrite).count());

        tp << ("#" + to_string(i))
            << std::chrono::duration_cast<std::chrono::milliseconds>(endOpen - startOpen).count()
            << std::chrono::duration_cast<std::chrono::milliseconds>(endWrite - startWrite).count()
            << static_cast<unsigned long>(dims / (time / 1000))
            << std::chrono::duration_cast<std::chrono::milliseconds>(endSave - startSave).count();

    }

    double dims = static_cast<double>(rows*columns*repetitions);
    double time = static_cast<double>(TotalWritingTime);


    tp.PrintFooter();
    tp << "Average"
        << TotalOpeningTime/repetitions
        << TotalWritingTime/repetitions
        << static_cast<unsigned long>(dims / (time / 1000))
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


