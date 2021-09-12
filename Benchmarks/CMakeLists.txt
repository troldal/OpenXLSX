#=======================================================================================================================
# Download Google Benchmark
#=======================================================================================================================
include(FetchContent)
FetchContent_Declare(gbench
                     GIT_REPOSITORY https://github.com/google/benchmark.git
                     GIT_TAG main
                     )

FetchContent_GetProperties(gbench)
if (NOT gbench_POPULATED)
    FetchContent_Populate(gbench)
    set(BENCHMARK_ENABLE_TESTING OFF)
    add_subdirectory(${gbench_SOURCE_DIR} ${gbench_BINARY_DIR} EXCLUDE_FROM_ALL)
endif ()


#=======================================================================================================================
# Define Benchmark targets
#=======================================================================================================================
add_executable(OpenXLSXBenchmark EXCLUDE_FROM_ALL Benchmark.cpp)
target_link_libraries(OpenXLSXBenchmark PRIVATE benchmark benchmark_main OpenXLSX::OpenXLSX)
