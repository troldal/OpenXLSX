//
// Created by Kenneth Balslev on 19/06/2020.
//

#pragma warning(push)
#pragma warning(disable : 4244)

#include <OpenXLSX.hpp>
#include <benchmark/benchmark.h>
#include <numeric>

using namespace OpenXLSX;

/**
 * @brief
 * @param state
 */
static void BM_WriteStrings(benchmark::State& state)
{
    XLDocument doc;
    doc.create("./benchmark_strings.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto arange = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    for (auto _ : state)
        for (auto& cell : arange) cell.value() = "OpenXLSX";

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteStrings)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_WriteIntegers(benchmark::State& state)
{
    XLDocument doc;
    doc.create("./benchmark_integers.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto arange = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    for (auto _ : state)
        for (auto& cell : arange) cell.value() = 42;

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteIntegers)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_WriteFloats(benchmark::State& state)
{
    XLDocument doc;
    doc.create("./benchmark_floats.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto arange = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    for (auto _ : state)
        for (auto& cell : arange) cell.value() = 3.14;

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteFloats)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_WriteBools(benchmark::State& state)
{
    XLDocument doc;
    doc.create("./benchmark_bools.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto arange = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));

    for (auto _ : state)
        for (auto& cell : arange) cell.value() = true;

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.save();
    doc.close();
}

BENCHMARK(BM_WriteBools)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_ReadStrings(benchmark::State& state)
{
    XLDocument doc;
    doc.open("./benchmark_strings.xlsx");
    auto     wks = doc.workbook().worksheet("Sheet1");
    auto     rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));
    uint64_t result {};

    for (auto _ : state) {
        result =
            std::accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<std::string_view>().size(); });
        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadStrings)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_ReadIntegers(benchmark::State& state)
{
    XLDocument doc;
    doc.open("./benchmark_integers.xlsx");
    auto     wks = doc.workbook().worksheet("Sheet1");
    auto     rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));
    uint64_t result;

    for (auto _ : state) {
        result = std::accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<uint64_t>(); });
        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadIntegers)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_ReadFloats(benchmark::State& state)
{
    XLDocument doc;
    doc.open("./benchmark_floats.xlsx");
    auto        wks = doc.workbook().worksheet("Sheet1");
    auto        rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));
    long double result;

    for (auto _ : state) {
        result = std::accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<double>(); });
        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadFloats)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_ReadBools(benchmark::State& state)
{
    XLDocument doc;
    doc.open("./benchmark_bools.xlsx");
    auto     wks = doc.workbook().worksheet("Sheet1");
    auto     rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, 8));
    uint64_t result;

    for (auto _ : state) {
        result = std::accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<bool>(); });
        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(1048576 * 8);
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadBools)->Unit(benchmark::kMillisecond);

#pragma warning(pop)
