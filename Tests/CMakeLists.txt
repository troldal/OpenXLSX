#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Tests)

#=======================================================================================================================
# Set up Catch2 library
#=======================================================================================================================
add_library(Catch INTERFACE IMPORTED)
target_include_directories(Catch SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/catch/>)

#=======================================================================================================================
# Define TEST target
#=======================================================================================================================
add_executable(OpenXLSXTests EXCLUDE_FROM_ALL "")
target_sources(OpenXLSXTests
        PRIVATE
        main.cpp
        testXLDocument.cpp
        testXLCell.cpp
        testXLCellRange.cpp
        testXLCellReference.cpp
        testXLCellValue.cpp
        testXLCellValueProxy.cpp
        testXLColor.cpp
        testXLDateTime.cpp
        testXLFormula.cpp
        testXLRow.cpp
        testXLSheet.cpp
        )

target_link_libraries(OpenXLSXTests
        PRIVATE
        OpenXLSX::OpenXLSX
        Catch
        )