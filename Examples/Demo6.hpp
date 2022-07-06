#ifndef OPENXLSX_DEMO6_HPP
#define OPENXLSX_DEMO6_HPP

#include <OpenXLSX.hpp>

#include <deque>
#include <iostream>
#include <numeric>
#include <random>

int demo6()
{
    using namespace std;
    using namespace OpenXLSX;

    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #06: Row Handling (using containers)\n";
    cout << "********************************************************************************\n";

    // As an alternative to using cell ranges, you can use row ranges, and extract row values to a standard container
    // This can be significantly faster (up to twice as fast in some cases, for both reading and
    // writing).

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
    std::string path = "./Demo06.xlsx";
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

    // Reopen the spreadsheet...
    cout << "Re-opening spreadsheet ... ";
    cout.flush();
    t1 = high_resolution_clock::now();
    doc.open(path);
    wks = doc.workbook().worksheet("Sheet1");
    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << "(" << dur.count() << " seconds)\n";

    // Prepare for reading...
    uint64_t sum = 0;
    uint64_t count = 0;
    std::vector<XLCellValue> readValues;
    cout << "Reading data from spreadsheet ...";

    // The 'values()' member function can be used for both reading and writing
    // of data. In this case the 'readValues' variable is a std::vector of
    // XLCellValues, and the 'values()' method will perform implicit conversion
    // to the correct type. Again, all containers supporting bi-directional
    // iterators are supported. Also, if you have a vector of doubles, the values
    // will be converted, if possible. Note, however, that if conversion is not
    // possible, an exception will be thrown. For that reason, it is often best
    // to use a container of XLCellValue objects, and then later determine the
    // data type for each object.
    t1 = high_resolution_clock::now();
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

    t2 = high_resolution_clock::now();
    dur = t2 - t1;
    std::cout << " (" << dur.count() << " seconds)\n";

    cout << "Cell count: " << count << endl;
    cout << "Sum of cell values: " << sum << endl;

    doc.close();

    cout << "\nDEMO PROGRAM #06 COMPLETED\n\n";

    return 0;
}


#endif    // OPENXLSX_DEMO6_HPP
