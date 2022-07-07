#ifndef OPENXLSX_DEMO9_HPP
#define OPENXLSX_DEMO9_HPP

#include <OpenXLSX.hpp>

#include <chrono>
#include <deque>
#include <filesystem>
#include <future>
#include <iostream>
#include <numeric>
#include <random>
#include <sstream>

int demo9()
{
    using namespace std;
    using namespace OpenXLSX;

    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #09: Parallel Spreadsheet Handling\n";
    cout << "********************************************************************************\n";

    // OpenXLSX supports multi threading ... sort of. It is not possible to modify an Excel spreadsheet
    // from two or more threads simultaneously. Also, it is not possible to use the same XLDocument object
    // from two or more threads, even if it is only for reading. What you can do, is to simultaneously open
    // the same document from multiple XLDocument objects, which enables parallelism of sorts.
    // For large spreadsheets, however, this may take up a significant amount of memory.

    using std::chrono::high_resolution_clock;
    using std::chrono::duration_cast;
    using std::chrono::duration;
    using std::chrono::milliseconds;
    using std::chrono::seconds;

    decltype(high_resolution_clock::now()) t1;
    decltype(high_resolution_clock::now()) t2;
    duration<double> dur;

    // First, create a new document and access the sheet named 'Sheet1'.
    cout << "\nGenerating spreadsheet ... ";
    cout.flush();

    t1 = high_resolution_clock::now();
    XLDocument doc;
    std::string path = "./Demo09.xlsx";
    doc.create(path);
    auto wks = doc.workbook().worksheet("Sheet1");

    // Create a random-number generator (to be used later)
    std::random_device                 rand_dev;
    std::mt19937                       generator(rand_dev());
    std::uniform_int_distribution<int> distr(0, 99);

    // The XLWorksheet class has a 'rows()' method, that returns a XLRowRange
    // object. Similar to XLCellRange, XLRowRange provides begin and end iterators,
    // enabeling iteration through the individual XLRow objects.
    std::vector<XLCellValue> writeValues;
    for (auto& row : wks.rows(OpenXLSX::MAX_ROWS)) {
        writeValues.clear();
        for (int i = 0; i < 8; ++i) writeValues.emplace_back(distr(generator));

        // The XLRow class provides a 'values()' method. Similar to the 'value()'
        // method of the XLCell class, the 'values()' method returns a reference
        // to a proxy object that facilitates implicit conversions etc.
        // Using the 'values()' method, containers can be assigned to each row.
        // In this case, a std::vector is used, but any container supporting
        // bidirectional iterators (such as std::deque or std::list) can be
        // used.
        // In this example a vector of XLCellValue objects is used. This is
        // useful if you heterogeneous data (such as doubles, ints and strings)
        // in the same dataset. However, if you have homogeneous data, e.g. all
        // doubles, you don't have to use XLCellValue objects. You can, for
        // examlpe, assigne a std::vector of doubles directly to the row.
        row.values() = writeValues;
    }
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "(" << dur.count() << " seconds)\n";

    // Saving a large spreadsheet can take a while...
    cout << "Saving spreadsheet ... ";
    cout.flush();
    t1 = high_resolution_clock::now();
    doc.save();
    doc.close();
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "(" << dur.count() << " seconds)\n";

    // Create a lambda function, which will load the spreadsheet generated above.
    // This function will be used by the std::async objects below.
    auto load = []() -> double {
        auto t1 = high_resolution_clock::now();
        XLDocument doc;
        doc.open("./Demo09.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        std::vector<XLCellValue> readValues;
        uint64_t sum = 0;
        uint64_t count = 0;
        for (auto& row : wks.rows()) {
            readValues = row.values();

            // Count the number of cell values
            count += std::count_if(readValues.begin(), readValues.end(), [](const XLCellValue& v) {
                return v.type() != XLValueType::Empty;
            });

            // Sum the numbers in each cell.
            sum += std::accumulate(readValues.begin(),
                                   readValues.end(),
                                   0,
                                   [](uint64_t a, XLCellValue& b) {return a + b.get<uint64_t>(); });
        }
        auto t2 = high_resolution_clock::now();

        duration<double> dur = t2 - t1;
        doc.close();

        return dur.count();
    };

    // Launch two threads for simulteneous reading of the same sheet.
    cout << "Launching 2 threads ... " << endl;

    t1 = high_resolution_clock::now();
    std::future<double> answer0 = std::async(load);
    std::future<double> answer1 = std::async(load);

    cout << "Thread 0: " << answer0.get() << " seconds" << endl;
    cout << "Thread 1: " << answer1.get() << " seconds" << endl;

    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "Total execution time for all threads: " << dur.count() << " seconds\n";

    cout << "\nDEMO PROGRAM #09 COMPLETED\n\n";

    return 0;
}


#endif    // OPENXLSX_DEMO9_HPP
