#include <iostream>
#include <chrono>
#include <iomanip>
#include <sstream>
#include <fstream>
#include "table_printer.h"
#include <OpenXLSX/OpenXLSX.h>

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

unsigned long ReadTest(int repetitions,
                       const std::string &fileName,
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

    str << "***********************************";
    str << " RUNNING WRITE BENCHMARKS ";
    str << "***********************************"<< endl << endl;

    WriteTest("Hello, OpenXLSX!", 1000, 1000, 5, "BenchmarkString.xlsx", str);
    //WriteTest(-42, 1000, 1000, 10, "BenchmarkInteger.xlsx", str);
    //WriteTest(3.14159L, 1000, 1000, 10, "BenchmarkFloat.xlsx", str);
    //WriteTest(true, 1000, 1000, 10, "BenchmarkBool.xlsx", str);

    cout << endl << endl;
    str << "***********************************";
    str << " RUNNING READ BENCHMARKS ";
    str << "************************************"<< endl << endl;

    ReadTest(5, "BenchmarkString.xlsx", str);
    //ReadTest(5, "BenchmarkInteger.xlsx", str);
    //ReadTest(5, "BenchmarkFloat.xlsx", str);
    //ReadTest(5, "BenchmarkBool.xlsx", str);

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
        XLDocument doc;
        doc.CreateDocument(fileName);
        auto wks = doc.Workbook().Worksheet("Sheet1");
        auto endOpen = chrono::steady_clock::now();

        // Add data to document
        auto startWrite = chrono::steady_clock::now();
        auto arange = wks.Range(XLCellReference("A1"), XLCellReference(rows, columns));

        for (auto iter : arange) {
            iter.Value().Set(value);
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

unsigned long ReadTest(int repetitions, const std::string &fileName, ostream &destination)
{
    TablePrinter tp(&destination);
    tp.AddColumn("Run #", 10);
    tp.AddColumn("Opening (ms)", 27);
    tp.AddColumn("Reading (ms)", 27);
    tp.AddColumn("Reading (cells/s)", 27);

    XLDocument doc;
    doc.OpenDocument(fileName);
    auto rows = doc.Workbook().Worksheet("Sheet1").RowCount();
    auto cols = doc.Workbook().Worksheet("Sheet1").ColumnCount();
    stringstream ss;
    ss << "Reading \""
        << doc.Workbook().Worksheet("Sheet1").Cell("A1").Value().AsString()
        << "\" from "
        << to_string(rows)
        << " x "
        << to_string(cols)
        << " cells";

    doc.CloseDocument();
    tp.PrintTitle(ss.str());
    tp.PrintHeader();


    unsigned long TotalOpeningTime = 0;
    unsigned long TotalReadingTime = 0;
    for (int i = 1; i <= repetitions; i++) {

        // Create new document
        auto startOpen = chrono::steady_clock::now();
        OpenXLSX::XLDocument doc;
        doc.OpenDocument(fileName);
        auto wks = doc.Workbook().Worksheet("Sheet1");
        auto endOpen = chrono::steady_clock::now();

        // Add data to document
        auto startRead = chrono::steady_clock::now();
        int intVal;
        double doubleVal;
        std::string stringVal;
        bool boolVal;
        auto arange = wks.Range(XLCellReference("A1"), XLCellReference(rows, cols));
        for (const auto iter : arange) {
            switch(iter.Value().ValueType()) {
                case XLValueType::Integer :
                    intVal = iter.Value().Get<int>();
                    break;

                case XLValueType::Float :
                    doubleVal = iter.Value().Get<double>();
                    break;

                case XLValueType::String :
                    stringVal = iter.Value().Get<std::string>();
                    break;

                case XLValueType::Boolean :
                    boolVal = iter.Value().Get<bool>();
                    break;

                default:
                    break;

            }
        }
        auto endRead = chrono::steady_clock::now();
        doc.CloseDocument();

        TotalOpeningTime += std::chrono::duration_cast<std::chrono::milliseconds>(endOpen - startOpen).count();
        TotalReadingTime += std::chrono::duration_cast<std::chrono::milliseconds>(endRead - startRead).count();

        double dims = static_cast<double>(rows*cols);
        double time = static_cast<double>(std::chrono::duration_cast<std::chrono::milliseconds>(endRead - startRead).count());

        tp << ("#" + to_string(i))
           << std::chrono::duration_cast<std::chrono::milliseconds>(endOpen - startOpen).count()
           << std::chrono::duration_cast<std::chrono::milliseconds>(endRead - startRead).count()
           << static_cast<unsigned long>(dims / (time / 1000));

    }

    double dims = static_cast<double>(rows*cols*repetitions);
    double time = static_cast<double>(TotalReadingTime);


    tp.PrintFooter();
    tp << "Average"
       << TotalOpeningTime/repetitions
       << TotalReadingTime/repetitions
       << static_cast<unsigned long>(dims / (time / 1000));
    tp.PrintFooter();

    cout << endl;


    return 0;
}

void openLarge() {

    auto start = chrono::steady_clock::now();

    OpenXLSX::XLDocument doc;
    doc.OpenDocument("./Large.xlsx");
    auto wks = doc.Workbook().Worksheet("Sheet1");
    wks.Export("./Profiles.csv");

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


