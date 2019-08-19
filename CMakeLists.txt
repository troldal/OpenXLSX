#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.12 FATAL_ERROR)
project(OpenXLSX)
set(OPENXLSX_DEPENDENCIES ${CMAKE_CURRENT_BINARY_DIR}/dependencies)
set(OPENXLSX_INSTALLDIR ${CMAKE_CURRENT_BINARY_DIR}/install)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)

#=======================================================================================================================
# Set project metadata
#=======================================================================================================================
set(PROJECT_VENDOR "Kenneth Troldal Balslev")
set(PROJECT_CONTACT "kenneth.balslev@gmail.com")
set(PROJECT_URL "https://github.com/troldal/OpenXLSX")
set(PROJECT_DESCRIPTION "A C++17 library for reading, writing and modifying Excel spreadsheets")

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

#=======================================================================================================================
# Add build options
#=======================================================================================================================
option(STATIC "Set to ON to build a static library instance of a shared library" OFF)
option(CREATE_DOCS "Build library documentation (requires Doxygen and Graphviz/Dot to be installed)" ON)
option(BUILD_SAMPLES "Build sample programs" ON)
option(BUILD_BENCHMARKS "Build and run library benchmarks" ON)
option(BUILD_TESTS "Build and run library tests" ON)

#=======================================================================================================================
# Add project subdirectories
#=======================================================================================================================
add_subdirectory(library)

if(${CREATE_DOCS})
    add_subdirectory(doxy)
endif()

if(${BUILD_BENCHMARKS})
    add_subdirectory(benchmark)
    add_dependencies(Benchmark OpenXLSX)
endif()

if(${BUILD_SAMPLES})
    add_subdirectory(examples)
    add_dependencies(Demo1 OpenXLSX)
endif()

if(${BUILD_TESTS})
    add_subdirectory(tests)
    add_dependencies(InterfaceTest OpenXLSX)
endif()






