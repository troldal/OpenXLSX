#=======================================================================================================================
# Download Google Benchmark
#=======================================================================================================================
include(FetchContent)
FetchContent_Declare(gbench
                     GIT_REPOSITORY https://github.com/google/benchmark.git
                     GIT_TAG master
                     )

FetchContent_GetProperties(gbench)
if (NOT gbench_POPULATED)
    FetchContent_Populate(gbench)
    set(BENCHMARK_ENABLE_TESTING OFF)
    add_subdirectory(${gbench_SOURCE_DIR} ${gbench_BINARY_DIR})
endif ()


#=======================================================================================================================
# Define Benchmark targets
#=======================================================================================================================

# ===== Target using static library ===== #
add_executable(OpenXLSXBenchmark-static Benchmark.cpp)
target_link_libraries(OpenXLSXBenchmark-static PRIVATE benchmark benchmark_main OpenXLSX-static)
target_compile_definitions(OpenXLSXBenchmark-static PRIVATE OPENXLSX_STATIC_DEFINE)

# ===== Target using shared library ===== #
add_executable(OpenXLSXBenchmark-shared Benchmark.cpp)
target_link_libraries(OpenXLSXBenchmark-shared PRIVATE benchmark benchmark_main OpenXLSX-shared)

