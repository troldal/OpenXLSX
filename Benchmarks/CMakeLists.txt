cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Benchmark)

#=======================================================================================================================
# Define Benchmark targets
#=======================================================================================================================
set(BENCHMARK_ENABLE_TESTING OFF)
add_subdirectory(gbench)
add_executable(OpenXLSXBenchmark EXCLUDE_FROM_ALL Benchmark.cpp)
target_link_libraries(OpenXLSXBenchmark PRIVATE benchmark::benchmark benchmark::benchmark_main OpenXLSX::OpenXLSX)
