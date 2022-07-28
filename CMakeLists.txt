#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX VERSION 0.4.1 LANGUAGES CXX)

set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)

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

#=======================================================================================================================
# Add build options
#=======================================================================================================================
if(NOT DEFINED OPENXLSX_MASTER_PROJECT)
    if(CMAKE_CURRENT_SOURCE_DIR STREQUAL CMAKE_SOURCE_DIR)
        set(OPENXLSX_MASTER_PROJECT ON)
    else()
        set(OPENXLSX_MASTER_PROJECT OFF)
    endif()
endif()

option(OPENXLSX_CREATE_DOCS "Build library documentation (requires Doxygen and Graphviz/Dot to be installed)" ${OPENXLSX_MASTER_PROJECT})
option(OPENXLSX_BUILD_SAMPLES "Build sample programs" ${OPENXLSX_MASTER_PROJECT})
option(OPENXLSX_BUILD_TESTS "Build and run library tests" ${OPENXLSX_MASTER_PROJECT})
option(OPENXLSX_BUILD_BENCHMARKS "Build and run library benchmarks" ${OPENXLSX_MASTER_PROJECT})
option(OPENXLSX_ENABLE_LIBZIP "Enable using libzip" OFF)

#=======================================================================================================================
# Add project subdirectories
#=======================================================================================================================
add_subdirectory(OpenXLSX)

if(${OPENXLSX_CREATE_DOCS})
    add_subdirectory(Documentation)
endif()

if(${OPENXLSX_BUILD_SAMPLES})
    add_subdirectory(Examples)
endif()

if(${OPENXLSX_BUILD_TESTS})
    add_subdirectory(Tests)
endif()

if(${OPENXLSX_BUILD_BENCHMARKS})
    add_subdirectory(Benchmarks)
endif()

