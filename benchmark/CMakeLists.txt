#=======================================================================================================================
# Download Google Benchmark
#=======================================================================================================================
include(FetchContent)
FetchContent_Declare(gbench
        GIT_REPOSITORY https://github.com/google/benchmark.git
        GIT_TAG master
        )

FetchContent_GetProperties(gbench)
if(NOT gbench_POPULATED)
    FetchContent_Populate(gbench)
    set(BENCHMARK_ENABLE_TESTING OFF)
    add_subdirectory(${gbench_SOURCE_DIR} ${gbench_BINARY_DIR})
endif()


#=======================================================================================================================
# Define Benchmark target
#=======================================================================================================================
add_executable(TestBench bench.cpp)
target_link_libraries(TestBench PRIVATE benchmark OpenXLSX-static)



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

