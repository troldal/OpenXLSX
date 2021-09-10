#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
set(OPENXLSX_MAJOR_VERSION 0)
set(OPENXLSX_MINOR_VERSION 3)
set(OPENXLSX_MICRO_VERSION 1)
project(OpenXLSX VERSION 0.3.1 LANGUAGES CXX)
set(CMAKE_DEBUG_POSTFIX d)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

#=======================================================================================================================
# Set project metadata
#=======================================================================================================================
set(OPENXLSX_PROJECT_VENDOR "Kenneth Troldal Balslev")
set(OPENXLSX_PROJECT_CONTACT "kenneth.balslev@gmail.com")
set(OPENXLSX_PROJECT_URL "https://github.com/troldal/OpenXLSX")
set(PROJECT_DESCRIPTION "A C++17 library for reading, writing and modifying Excel spreadsheets")

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(IGNORE_ME ${CMAKE_C_COMPILER}) # Suppress CMake warning message

#=======================================================================================================================
# Output Directories.
#=======================================================================================================================
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
#set(CMAKE_INSTALL_PREFIX ${CMAKE_CURRENT_LIST_DIR}/stage)

#=======================================================================================================================
# Add build options
#=======================================================================================================================
option(OPENXLSX_CREATE_DOCS "Build library documentation (requires Doxygen and Graphviz/Dot to be installed)" OFF)
option(OPENXLSX_COMPACT_MODE "Build library in compact mode (slower, but uses less memory)" OFF)
option(OPENXLSX_BUILD_SAMPLES "Build sample programs" ON)
option(OPENXLSX_BUILD_TESTS "Build and run library tests" OFF)
option(OPENXLSX_BUILD_BENCHMARKS "Build and run library benchmarks" OFF)
set(OPENXLSX_LIBRARY_TYPE "STATIC" CACHE STRING "Set the library type to SHARED or STATIC")

#=======================================================================================================================
# Add project subdirectories
#=======================================================================================================================
add_subdirectory(library)

if(${OPENXLSX_CREATE_DOCS})
    add_subdirectory(doxy)
endif()

if(${OPENXLSX_BUILD_SAMPLES})
    add_subdirectory(examples)
endif()

if(${OPENXLSX_BUILD_TESTS})
    add_subdirectory(tests)
endif()

if(${OPENXLSX_BUILD_BENCHMARKS})
    add_subdirectory(benchmark)
endif()

