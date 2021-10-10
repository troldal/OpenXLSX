//
// Created by Kenneth Balslev on 27/08/2021.
//

#include <OpenXLSX.hpp>
#include <catch.hpp>
#include <fstream>
#include <vector>
#include <list>
#include <deque>

using namespace OpenXLSX;

TEST_CASE("XLRow Tests", "[XLRow]")
{
    SECTION("XLRow")
    {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto row_3 = wks.row(3);
        REQUIRE(row_3.rowNumber() == 3);

        XLRow copy1  = row_3;
        auto  height = copy1.height();
        REQUIRE(copy1.height() == height);

        XLRow copy2   = std::move(copy1);
        auto  descent = copy2.descent();
        REQUIRE(copy2.descent() == descent);

        XLRow copy3;
        copy3 = copy2;
        copy3.setHeight(height * 3);
        REQUIRE(copy3.height() == height * 3);
        copy3.setHeight(height * 4);
        REQUIRE(copy3.height() == height * 4);

        XLRow copy4;
        copy4 = std::move(copy3);
        copy4.setDescent(descent * 2);
        REQUIRE(copy4.descent() == descent * 2);
        copy4.setDescent(descent * 3);
        REQUIRE(copy4.descent() == descent * 3);

        REQUIRE(copy4.isHidden() == false);
        copy4.setHidden(true);
        REQUIRE(copy4.isHidden() == true);
        copy4.setHidden(false);
        REQUIRE(copy4.isHidden() == false);


        auto row1 = wks.row(11);
        auto row2 = wks.row(12);
        auto row3 = wks.row(13);
        auto row4 = wks.row(13);
        REQUIRE(row3 == row4);
        REQUIRE(row4 >= row3);
        REQUIRE(row3 <= row4);
        REQUIRE(row3 >= row2);
        REQUIRE(row2 <= row3);
        REQUIRE(row1 != row2);
        REQUIRE(row2 > row1);
        REQUIRE(row1 < row2);

        doc.save();
    }
}

TEST_CASE("XLRowData Tests", "[XLRowData]")
{

    SECTION("XLRowData (vector<int>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::vector<int> {1, 2, 3, 4, 5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::vector<int>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 15);

        val1sum = 0;
        const auto val1results2 = static_cast<std::vector<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<int>();
        REQUIRE(val1sum == 15);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (vector<double>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::vector<double> {1.1, 2.2, 3.3, 4.4, 5.5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0.0;
        const auto val1results1 = static_cast<std::vector<double>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 16.5);

        val1sum = 0;
        const auto val1results2 = static_cast<std::vector<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<double>();
        REQUIRE(val1sum == 16.5);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (vector<bool>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::vector<bool> {true, false, true, false, true};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::vector<bool>>(row1c.values());
        for (const auto v : val1results1)
            if (v) val1sum += 1;
        REQUIRE(val1sum == 3);

        val1sum = 0;
        const auto val1results2 = static_cast<std::vector<XLCellValue>>(row1c.values());
        for (const auto v : val1results2)
            if(v.get<bool>()) val1sum += 1;
        REQUIRE(val1sum == 3);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (vector<std::string>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::vector<std::string> {"This", "is", "a", "test."};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::vector<std::string>>(row1c.values());
        for (const auto v : val1results1) val1sum += v.size();
        REQUIRE(val1sum == 12);

        val1sum = 0;
        const auto val1results2 = static_cast<std::vector<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<std::string>().size();
        REQUIRE(val1sum == 12);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (vector<XLCellValue>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::vector<XLCellValue> {1, 2, 3, 4, 5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::vector<int>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 15);

        val1sum = 0;
        const auto val1results2 = static_cast<std::vector<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<int>();
        REQUIRE(val1sum == 15);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (list<int>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::list<int> {1, 2, 3, 4, 5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::list<int>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 15);

        val1sum = 0;
        const auto val1results2 = static_cast<std::list<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<int>();
        REQUIRE(val1sum == 15);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (list<bool>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::list<bool> {true, false, true, false, true};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::list<bool>>(row1c.values());
        for (const auto v : val1results1)
            if (v) val1sum += 1;
        REQUIRE(val1sum == 3);

        val1sum = 0;
        const auto val1results2 = static_cast<std::list<XLCellValue>>(row1c.values());
        for (const auto v : val1results2)
            if (v.get<bool>()) val1sum += 1;
        REQUIRE(val1sum == 3);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (list<double>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::list<double> {1.1, 2.2, 3.3, 4.4, 5.5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0.0;
        const auto val1results1 = static_cast<std::list<double>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 16.5);

        val1sum = 0;
        const auto val1results2 = static_cast<std::list<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<double>();
        REQUIRE(val1sum == 16.5);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (list<std::string>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::list<std::string> {"This", "is", "a", "test."};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::list<std::string>>(row1c.values());
        for (const auto v : val1results1) val1sum += v.size();
        REQUIRE(val1sum == 12);

        val1sum = 0;
        const auto val1results2 = static_cast<std::list<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<std::string>().size();
        REQUIRE(val1sum == 12);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (list<XLCellValue>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::list<XLCellValue> {1, 2, 3, 4, 5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::list<int>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 15);

        val1sum = 0;
        const auto val1results2 = static_cast<std::list<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<int>();
        REQUIRE(val1sum == 15);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (deque<int>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::deque<int> {1, 2, 3, 4, 5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::deque<int>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 15);

        val1sum = 0;
        const auto val1results2 = static_cast<std::deque<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<int>();
        REQUIRE(val1sum == 15);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (deque<bool>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::deque<bool> {true, false, true, false, true};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::deque<bool>>(row1c.values());
        for (const auto v : val1results1)
            if (v) val1sum += 1;
        REQUIRE(val1sum == 3);

        val1sum = 0;
        const auto val1results2 = static_cast<std::deque<XLCellValue>>(row1c.values());
        for (const auto v : val1results2)
            if (v.get<bool>()) val1sum += 1;
        REQUIRE(val1sum == 3);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (deque<double>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::deque<double> {1.1, 2.2, 3.3, 4.4, 5.5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0.0;
        const auto val1results1 = static_cast<std::deque<double>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 16.5);

        val1sum = 0;
        const auto val1results2 = static_cast<std::deque<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<double>();
        REQUIRE(val1sum == 16.5);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (deque<std::string>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::deque<std::string> {"This", "is", "a", "test."};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::deque<std::string>>(row1c.values());
        for (const auto v : val1results1) val1sum += v.size();
        REQUIRE(val1sum == 12);

        val1sum = 0;
        const auto val1results2 = static_cast<std::deque<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<std::string>().size();
        REQUIRE(val1sum == 12);

        row1.values().clear();

        doc.save();
    }

    SECTION("XLRowData (deque<XLCellValue>)") {
        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto val1 = std::deque<XLCellValue> {1, 2, 3, 4, 5};
        auto row1 = wks.row(1);
        row1.values() = val1;

        const auto row1c = row1;

        auto val1sum = 0;
        const auto val1results1 = static_cast<std::deque<int>>(row1c.values());
        for (const auto v : val1results1) val1sum += v;
        REQUIRE(val1sum == 15);

        val1sum = 0;
        const auto val1results2 = static_cast<std::deque<XLCellValue>>(row1c.values());
        for (const auto v : val1results2) val1sum += v.get<int>();
        REQUIRE(val1sum == 15);

        row1.values().clear();

        doc.save();
    }
}

TEST_CASE("XLRowDataRange Tests", "[XLRowDataRange]")
{
    SECTION("XLRowDataRange") {

        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto row = wks.row(1);
        auto range = row.cells();
        for (auto& cell : range) cell.value() = 1;

        auto sum = 0;
        for (const auto& cell : range) sum += cell.value().get<int>();
        REQUIRE(sum == 1);
        REQUIRE(range.size() == 1);

        auto range_copy = range;
        sum = 0;
        for (const auto& cell : range_copy) sum += cell.value().get<int>();
        REQUIRE(sum == 1);
        REQUIRE(range_copy.size() == 1);

        auto range_move = std::move(range_copy);
        sum = 0;
        for (const auto& cell : range_move) sum += cell.value().get<int>();
        REQUIRE(sum == 1);
        REQUIRE(range_move.size() == 1);

        auto range_copy2 = range_move;
        range_copy2 = range;
        sum = 0;
        for (const auto& cell : range_copy2) sum += cell.value().get<int>();
        REQUIRE(sum == 1);
        REQUIRE(range_copy2.size() == 1);

        auto range_move2 = range_move;
        range_move2 = std::move(range_copy2);
        sum = 0;
        for (const auto& cell : range_move2) sum += cell.value().get<int>();
        REQUIRE(sum == 1);
        REQUIRE(range_move2.size() == 1);

        auto row2 = wks.row(2);
        auto range2 = row2.cells(8);
        for (auto& cell : range2) cell.value() = 1;

        sum = 0;
        for (const auto& cell : row2.cells()) sum += cell.value().get<int>();
        REQUIRE(sum == 8);
        REQUIRE(range2.size() == 8);

        auto row3 = wks.row(3);
        auto range3 = row3.cells(3, 8);
        for (auto& cell : range3) cell.value() = 1;

        sum = 0;
        for (const auto& cell : row3.cells())
            if (cell.value().type() == XLValueType::Integer) sum += cell.value().get<int>();
        REQUIRE(sum == 6);
        REQUIRE(row3.cells().size() == 8);

        doc.save();
    }

    SECTION("XLRowDataIterator") {

        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto row = wks.row(1);
        auto range = row.cells(2);
        for (auto& cell : range) cell.value() = 1;
        auto begin = range.begin();
        auto end = range.end();
        auto other = begin;

        REQUIRE(begin == other);
        REQUIRE_FALSE(begin == end);
        REQUIRE(begin != end);
        REQUIRE_FALSE(begin != other);

        ++other;
        REQUIRE_FALSE(begin == other);

        ++other;
        REQUIRE(other == end);

        other = begin;
        REQUIRE(begin == other);

        other++;
        other = std::move(begin);
        begin = range.begin();
        REQUIRE(begin == other);

    }

}

TEST_CASE("XLRowIterator Tests", "[XLRowDataRange]")
{
    SECTION("XLRowIterator") {

        XLDocument doc;
        doc.create("./testXLRow.xlsx");
        auto wks = doc.workbook().worksheet("Sheet1");

        auto range = wks.rows(1,3);
        auto first = range.begin();
        auto second = first;
        second++;
        auto third = range.end();
        auto fourth = std::move(third);
        third = second;
        second = std::move(first);
        first = range.begin();

        REQUIRE(second == range.begin());
        REQUIRE_FALSE(second == third);
        REQUIRE_FALSE(second == fourth);
        REQUIRE(range.rowCount() == 3);
        REQUIRE(first->rowNumber() == 1);

    }
}