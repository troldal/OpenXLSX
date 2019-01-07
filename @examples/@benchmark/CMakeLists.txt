cmake_minimum_required(VERSION 3.5)
project(OpenXLSX.Benchmark)

# C/C++ language version
set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(SOURCE_FILES benchmark.cpp table_printer.h)

add_executable(Benchmark ${SOURCE_FILES})
target_link_libraries(Benchmark PUBLIC -lOpenXLSX)
target_link_directories(Benchmark PUBLIC ${OPENXLSX_INSTALLDIR}/lib)
target_link_directories(Benchmark PUBLIC ${OPENXLSX_INSTALLDIR}/bin)
target_include_directories(Benchmark PUBLIC ${OPENXLSX_INSTALLDIR}/include)
target_include_directories(Benchmark PUBLIC ${OPENXLSX_DEPENDENCIES}/include)

set_target_properties(Benchmark PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
