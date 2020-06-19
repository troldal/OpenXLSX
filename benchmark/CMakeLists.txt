#=======================================================================================================================
# Download Google Benchmark
#=======================================================================================================================
include(ExternalProject)
ExternalProject_Add(GBench
        GIT_REPOSITORY https://github.com/google/benchmark.git
        GIT_TAG master
        EP_BASE ${CMAKE_CURRENT_BINARY_DIR}/GBench
        INSTALL_COMMAND ""
        CMAKE_ARGS  -D BENCHMARK_ENABLE_GTEST_TESTS=OFF
                    -D CMAKE_BUILD_TYPE=Release
        )

set_target_properties(GBench PROPERTIES EXCLUDE_FROM_ALL TRUE)
ExternalProject_Get_Property(GBench BINARY_DIR)
add_library(googlebench STATIC IMPORTED)
set_property(TARGET googlebench PROPERTY IMPORTED_LOCATION ${BINARY_DIR}/src/libbenchmark.a)
#=======================================================================================================================
# Define Benchmark target
#=======================================================================================================================




#=======================================================================================================================
# Define Benchmark target
#=======================================================================================================================
add_executable(Benchmark-Static Benchmark.cpp table_printer.h)
target_link_libraries(Benchmark-Static PRIVATE OpenXLSX-static)

add_executable(Benchmark-Shared Benchmark.cpp table_printer.h)
target_link_libraries(Benchmark-Shared PRIVATE OpenXLSX-shared)

add_executable(Benchmark-Obj Benchmark.cpp table_printer.h)
target_link_libraries(Benchmark-Obj PRIVATE OpenXLSX-objs)
if (CMAKE_BUILD_TYPE STREQUAL "Release")
    include(CheckIPOSupported)
    check_ipo_supported(RESULT result OUTPUT output)
    if (result)
        set_property(TARGET Benchmark-Obj PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
    endif ()
endif()

