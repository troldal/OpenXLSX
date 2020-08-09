//
// Created by Kenneth Balslev on 19/06/2020.
//

#pragma warning(push)
#pragma warning(disable : 4244)

#include <OpenXLSX.hpp>
#include <benchmark/benchmark.h>

using namespace OpenXLSX;

///**
// * @brief
// * @param state
// */
// static void BM_CreateCellReference(benchmark::State& state) {
//
//    auto cellRef = XLCellReference();
//
//    for (auto _ : state) {
//        cellRef = XLCellReference("A1");
//    }
//
//}

// BENCHMARK(BM_CreateCellReference);

///**
// * @brief
// * @param state
// */
// static void BM_GetCellCoordinates(benchmark::State& state) {
//
//    auto coordinates = XLCellReference::CoordinatesFromAddress("A1");
//
//    for (auto _ : state) {
//        coordinates = XLCellReference::CoordinatesFromAddress("A1");
//    }
//}
//
// BENCHMARK(BM_GetCellCoordinates);

///**
// * @brief
// * @param state
// */
// static void BM_WriteMatrix(benchmark::State& state)
//{
//    XLDocument doc;
//    doc.create("./benchmark.xlsx");
//    auto wks    = doc.workbook().worksheet("Sheet1");
//    auto arange = wks.range(XLCellReference("A1"), XLCellReference(state.range(0), state.range(0)));
//
//    for (auto _ : state)
//        for (auto& cell : arange) cell.value() = "OpenXLSX";
//
//    state.SetItemsProcessed(state.range(0) * state.range(0));
//    state.counters["items"] = state.items_processed();
//
//    doc.save();
//    doc.close();
//}
//
// BENCHMARK(BM_WriteMatrix)->RangeMultiplier(2)->Range(8, 8 << 9)->Unit(benchmark::kMillisecond);

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

    doc.saveAs("benchmark_out.xlsx");
    doc.close();
}

BENCHMARK(BM_WriteColumns)->RangeMultiplier(2)->Range(1, 1 << 8)->Unit(benchmark::kMillisecond);
//
///**
// * @brief
// * @param state
// */
// static void BM_WriteRows(benchmark::State& state) {
//    XLDocument doc;
//    doc.create("./benchmark.xlsx");
//    auto wks = doc.workbook().worksheet("Sheet1");
//    auto arange = wks.range(XLCellReference("A1"), XLCellReference(state.range(0), 16383));
//
//    for (auto _ : state)
//        for (auto& cell : arange) cell.value() = "OpenXLSX";
//
//    state.SetItemsProcessed(16384 * state.range(0));
//    state.counters["items"] = state.items_processed();
//
//    doc.save();
//    doc.close();
//}
//
// BENCHMARK(BM_WriteRows)->RangeMultiplier(2)->Range(8, 8 << 7)->Unit(benchmark::kMillisecond);

#pragma warning(pop)
