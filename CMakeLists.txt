#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.12 FATAL_ERROR)
project (OpenXLSX)
set(OPENXLSX_DEPENDENCIES ${CMAKE_CURRENT_BINARY_DIR}/dependencies)
set(OPENXLSX_INSTALLDIR ${CMAKE_CURRENT_BINARY_DIR}/install)
#set(CMAKE_CXX_VISIBILITY_PRESET hidden)
#set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)

#=======================================================================================================================
# Set project metadata
#=======================================================================================================================
set(PROJECT_VENDOR "Kenneth Troldal Balslev")
set(PROJECT_CONTACT "kenneth.balslev@gmail.com")
set(PROJECT_URL "https://github.com/troldal/OpenXLSX")
set(PROJECT_DESCRIPTION "A library for reading, writing and modifying Excel spreadsheets for C++17")

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

#=======================================================================================================================
# Add project subdirectories
#=======================================================================================================================
add_subdirectory(@library)
add_subdirectory(@examples)
add_subdirectory(@doxy)
add_subdirectory(@tests)
add_subdirectory(@benchmark)

#=======================================================================================================================
# Define internal dependencies
#=======================================================================================================================
add_dependencies(Demo1 OpenXLSX)
add_dependencies(Benchmark OpenXLSX)
add_dependencies(InterfaceTest OpenXLSX)


