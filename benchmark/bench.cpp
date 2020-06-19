//
// Created by Kenneth Balslev on 19/06/2020.
//

#include <benchmark/benchmark.h>
#include <OpenXLSX.hpp>
#include <cmath>

using namespace OpenXLSX;

static void BM_SomeFunction(benchmark::State& state) {
    XLDocument doc;
    doc.CreateDocument("./benchmark.xlsx");
    auto wks = doc.Workbook().Worksheet("Sheet1");
    auto arange = wks.Range(XLCellReference("A1"), XLCellReference(1000, 1000));

    for (auto _ : state) {
        for (auto iter : arange) {
            iter.Value().Set(3.1415);
        }
    }

    doc.SaveDocument();
    doc.CloseDocument();
}
// Register the function as a benchmark
BENCHMARK(BM_SomeFunction);
// Run the benchmark
BENCHMARK_MAIN();