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
static void BM_WriteColumns(benchmark::State& state)
{
    XLDocument doc;
    doc.create("./benchmark.xlsx");
    auto wks    = doc.workbook().worksheet("Sheet1");
    auto arange = wks.range(XLCellReference("A1"), XLCellReference(1048576, state.range(0)));

    for (auto _ : state)
        for (auto& cell : arange) cell.value() = 1;

    state.SetItemsProcessed(1048576 * state.range(0));
    state.counters["items"] = state.items_processed();

    doc.saveAs("./benchmark_out_" + std::to_string(state.range(0)) + ".xlsx");
    doc.close();
}

BENCHMARK(BM_WriteColumns)->RangeMultiplier(2)->Range(1, 1 << 10)->Unit(benchmark::kMillisecond);

/**
 * @brief
 * @param state
 */
static void BM_ReadColumns(benchmark::State& state)
{
    XLDocument doc;
    doc.open("./benchmark_out_" + std::to_string(state.range(0)) + ".xlsx");
    auto     wks = doc.workbook().worksheet("Sheet1");
    auto     rng = wks.range(XLCellReference("A1"), XLCellReference(1048576, state.range(0)));
    uint64_t result;

    for (auto _ : state) {
        result = std::accumulate(rng.begin(), rng.end(), 0, [](uint64_t a, XLCell& b) { return a + b.value().get<uint64_t>(); });
        benchmark::DoNotOptimize(result);
        benchmark::ClobberMemory();
    }

    state.SetItemsProcessed(1048576 * state.range(0));
    state.counters["items"] = state.items_processed();

    doc.close();
}

BENCHMARK(BM_ReadColumns)->RangeMultiplier(2)->Range(1, 1 << 10)->Unit(benchmark::kMillisecond);

#pragma warning(pop)
