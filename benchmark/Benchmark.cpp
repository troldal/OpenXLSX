//
// Created by Kenneth Balslev on 19/06/2020.
//

#include <benchmark/benchmark.h>
#include <OpenXLSX.hpp>

using namespace OpenXLSX::Impl;

/**
 * @brief
 * @param state
 */
static void BM_WriteMatrix(benchmark::State& state) {
    XLDocument doc;
    doc.CreateDocument("./benchmark.xlsx");
    auto wks = doc.Workbook()->Worksheet("Sheet1");
    //auto arange = wks->Range(XLCellReference("A1"), XLCellReference(state.range(0), state.range(0)));

    for (auto _ : state) {
//        for (auto iter : arange) {
//            iter.Value().Set(3.1415);
//        }

        for (auto i = 1; i < state.range(0); ++i) {
            for (auto j = 1; j < state.range(0); ++j)
                wks->Cell(i, j).Value().Set(3.1415);
        }

    }

    state.SetItemsProcessed(state.range(0) * state.range(0));
    state.counters["items"] = state.items_processed();

    doc.SaveDocument();
    doc.CloseDocument();
}

BENCHMARK(BM_WriteMatrix)->RangeMultiplier(2)->Range(8, 8 << 9)->Unit(benchmark::kMillisecond);

///**
// * @brief
// * @param state
// */
//static void BM_WriteColumns(benchmark::State& state) {
//    XLDocument doc;
//    doc.CreateDocument("./benchmark.xlsx");
//    auto wks = doc.Workbook()->Worksheet("Sheet1");
//    auto arange = wks->Range(XLCellReference("A1"), XLCellReference(1048575, state.range(0)));
//
//    for (auto _ : state) {
//        for (auto iter : arange) {
//            iter.Value().Set(3.1415);
//        }
//    }
//
//    state.SetItemsProcessed(1048576 * state.range(0));
//    state.counters["items"] = state.items_processed();
//
//    doc.SaveDocument();
//    doc.CloseDocument();
//}
//
//BENCHMARK(BM_WriteColumns)->RangeMultiplier(2)->Range(1, 1 << 4)->Unit(benchmark::kMillisecond);
//
///**
// * @brief
// * @param state
// */
//static void BM_WriteRows(benchmark::State& state) {
//    XLDocument doc;
//    doc.CreateDocument("./benchmark.xlsx");
//    auto wks = doc.Workbook()->Worksheet("Sheet1");
//    auto arange = wks->Range(XLCellReference("A1"), XLCellReference(state.range(0), 16383));
//
//    for (auto _ : state) {
//        for (auto iter : arange) {
//            iter.Value().Set(3.1415);
//        }
//    }
//
//    state.SetItemsProcessed(16384 * state.range(0));
//    state.counters["items"] = state.items_processed();
//
//    doc.SaveDocument();
//    doc.CloseDocument();
//}
//
//BENCHMARK(BM_WriteRows)->RangeMultiplier(2)->Range(8, 8 << 7)->Unit(benchmark::kMillisecond);