//
// Created by Kenneth Balslev on 19/06/2020.
//

#pragma warning(push)
#pragma warning(disable : 4244)

#include <OpenXLSX.hpp>
#include <benchmark/benchmark.h>
#include <numeric>
#include <cstdint>

using namespace OpenXLSX;

constexpr uint64_t rows    = 1048576;
constexpr uint8_t  columns = 8;

/**
 * @brief
 * @param state
 */
static void BM_WriteStrings(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.create("./benchmark_strings.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");

    std::vector<XLCellValue> values(columns, "OpenXLSX");

    for (auto _ : state)    // NOLINT
        for (auto& row : wks.rows())
            row.setValues(values);

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteStrings)->Unit(benchmark::kMillisecond);    // NOLINT

/**
 * @brief
 * @param state
 */
static void BM_WriteIntegers(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.create("./benchmark_integers.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");

    std::vector<XLCellValue> values(columns, 42);

    for (auto _ : state)    // NOLINT
        for (auto& row : wks.rows())
            row.setValues(values);

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteIntegers)->Unit(benchmark::kMillisecond);    // NOLINT

/**
 * @brief
 * @param state
 */
static void BM_WriteFloats(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.create("./benchmark_floats.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");

    std::vector<XLCellValue> values(columns, 3.14);

    for (auto _ : state)    // NOLINT
        for (auto& row : wks.rows())
            row.setValues(values);

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteFloats)->Unit(benchmark::kMillisecond);    // NOLINT

/**
 * @brief
 * @param state
 */
static void BM_WriteBools(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.create("./benchmark_bools.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");

    std::vector<XLCellValue> values(columns, true);

    for (auto _ : state)    // NOLINT
        for (auto& row : wks.rows())
            row.setValues(values);

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteBools)->Unit(benchmark::kMillisecond);    // NOLINT

/**
 * @brief
 * @param state
 */
static void BM_ReadStrings(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.open("./benchmark_strings.xlsx");
    auto     wks    = doc.workbook().worksheet("Sheet1");
    uint64_t result = 0;
    std::vector<XLCellValue> values;

    for (auto _ : state) {    // NOLINT
        for (auto& row : wks.rows()) {
            values = row.getValues(columns);
            result += std::count_if(values.begin(), values.end(), [](const XLCellValue& v) {
                return v.type() != XLValueType::Empty;
            });
        }
        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadStrings)->Unit(benchmark::kMillisecond);    // NOLINT

/**
 * @brief
 * @param state
 */
static void BM_ReadIntegers(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.open("./benchmark_integers.xlsx");
    auto     wks    = doc.workbook().worksheet("Sheet1");
    uint64_t result = 0;
    std::vector<XLCellValue> values;

    for (auto _ : state) {    // NOLINT
        for (auto& row : wks.rows()) {
            values = row.getValues(columns);
            result += std::accumulate(values.begin(),
                                      values.end(),
                                      0,
                                      [](uint64_t a, XLCellValue& b) { return a + b.get<uint64_t>(); });
        }

        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadIntegers)->Unit(benchmark::kMillisecond);    // NOLINT

/**
 * @brief
 * @param state
 */
static void BM_ReadFloats(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.open("./benchmark_floats.xlsx");
    auto        wks    = doc.workbook().worksheet("Sheet1");
    long double result = 0;
    std::vector<XLCellValue> values;

    for (auto _ : state) {    // NOLINT
        for (auto& row : wks.rows()) {
            values = row.getValues(columns);
            result += std::accumulate(values.begin(),
                                      values.end(),
                                      0,
                                      [](uint64_t a, XLCellValue& b) { return a + b.get<double>(); });
        }

        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadFloats)->Unit(benchmark::kMillisecond);    // NOLINT

/**
 * @brief
 * @param state
 */
static void BM_ReadBools(benchmark::State& state)    // NOLINT
{
    XLDocument doc;
    doc.open("./benchmark_bools.xlsx");
    auto     wks    = doc.workbook().worksheet("Sheet1");
    uint64_t result = 0;
    std::vector<XLCellValue> values;

    for (auto _ : state) {    // NOLINT
        for (auto& row : wks.rows()) {
            values = row.getValues(columns);
            result += std::accumulate(values.begin(),
                                      values.end(),
                                      0,
                                      [](uint64_t a, XLCellValue& b) { return a + b.get<bool>(); });
        }

        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(rows * columns);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadBools)->Unit(benchmark::kMillisecond);    // NOLINT

#pragma warning(pop)
