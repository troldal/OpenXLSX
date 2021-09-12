#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Python)

#=======================================================================================================================
# Define Benchmark targets
#=======================================================================================================================
if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
    set(PYBIND11_CPP_STANDARD /std:c++17)
else()
    set(PYBIND11_CPP_STANDARD -std=c++1z)
endif()

add_subdirectory(pybind11)

set(PYOPENXLSX_SOURCES
    ${CMAKE_CURRENT_LIST_DIR}/pyOpenXLSX.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLCell.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLCellRange.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLCellReference.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLCellValue.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLColor.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLColumn.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLDocument.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLRow.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLSheet.cpp
    ${CMAKE_CURRENT_LIST_DIR}/pyXLWorkbook.cpp
    )

pybind11_add_module(pyOpenXLSX ${PYOPENXLSX_SOURCES})

# ===== Target using static library ===== #
target_link_libraries(pyOpenXLSX PRIVATE OpenXLSX-static)
target_compile_definitions(pyOpenXLSX PRIVATE OPENXLSX_STATIC_DEFINE)
set_target_properties(pyOpenXLSX PROPERTIES OUTPUT_NAME OpenXLSX)

get_property(pyOpenXLSXSuffix TARGET pyOpenXLSX PROPERTY SUFFIX)
add_custom_command(TARGET pyOpenXLSX POST_BUILD
                   COMMAND ${CMAKE_COMMAND} -E copy $<TARGET_FILE:pyOpenXLSX>
                   ${CMAKE_CURRENT_LIST_DIR}/../examples/OpenXLSX${pyOpenXLSXSuffix}
                   )


