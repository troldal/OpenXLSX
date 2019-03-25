#=======================================================================================================================
# Define library targets
#=======================================================================================================================
add_library(OpenXLSX SHARED "")
add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)

#=======================================================================================================================
# Set compiler flags
#=======================================================================================================================
include(CheckIPOSupported)
if (CMAKE_BUILD_TYPE STREQUAL "Debug")
    target_compile_options(OpenXLSX PRIVATE -Wall)
    target_compile_options(OpenXLSX PRIVATE -Wextra)
    target_compile_options(OpenXLSX PRIVATE -pedantic)
    #target_compile_options(OpenXLSX PRIVATE -fsanitize=leak -g)
    #target_link_options(OpenXLSX PRIVATE -fsanitize=leak -g)
    if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
        #target_compile_options(OpenXLSX PRIVATE -Weverything)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        target_compile_options(OpenXLSX PRIVATE -Weffc++)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Intel")
        # using Intel C++
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
        # using Visual Studio C++
    endif ()
elseif (CMAKE_BUILD_TYPE STREQUAL "Release")
    check_ipo_supported(RESULT result OUTPUT output)
    if (result)
        set_property(TARGET OpenXLSX PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
    endif ()
endif ()

#=======================================================================================================================
# Find project files
#=======================================================================================================================
set(OPENXLSX_IMPL_HEADERS
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLAbstractXMLFile_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLAppProperties_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLColor_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLContentTypes_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCoreProperties_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLDocument_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLFont_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLRelationships_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLSharedStrings_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLStyles_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLTokenizer_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLWorkbook_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLChartsheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLColumn_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLRow_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLSheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLTemplate_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLWorksheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCell_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellFormats_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellFormula_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellRange_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellReference_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellType_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellValue_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLXml_Impl.h
        )

set(OPENXLSX_IMPL_SOURCES
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLAbstractXMLFile_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLAppProperties_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLColor_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLContentTypes_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCoreProperties_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLDocument_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLFont_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLRelationships_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLSharedStrings_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLStyles_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLTokenizer_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLWorkbook_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLChartsheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLColumn_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLRow_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLSheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLWorksheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCell_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCellFormats_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCellFormula_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCellRange_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCellReference_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCellValue_Impl.cpp
        )

set(OPENXLSX_CXX_INTERFACE_HEADERS
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/OpenXLSX.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLDefinitions.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLDocument.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLWorkbook.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLSheet.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLWorksheet.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLCellReference.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLCell.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLCellValue.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLCellRange.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLChartsheet.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLRow.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLColumn.h
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/XLException.h)

set(OPENXLSX_CXX_INTERFACE_SOURCES
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLDocument.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLWorkbook.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLSheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLWorksheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLCellReference.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLCell.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLCellValue.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLCellRange.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLChartsheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLRow.cpp
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/sources/XLColumn.cpp)

#=======================================================================================================================
# Build OpenXLSX library
#=======================================================================================================================
target_sources(OpenXLSX
        PRIVATE
        ${OPENXLSX_IMPL_SOURCES}
        ${OPENXLSX_CXX_INTERFACE_SOURCES})
target_link_libraries(OpenXLSX
        PRIVATE
        PugiXML Zippy)
target_include_directories(OpenXLSX
        PRIVATE
        ${OPENXLSX_DEPENDENCIES}/include)
target_include_directories(OpenXLSX
        PRIVATE
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers
        PUBLIC
        ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers)
#SYSTEM INTERFACE
#    ${CMAKE_CURRENT_LIST_DIR}/implementation/headers
#    ${CMAKE_CURRENT_LIST_DIR}/interfaces/headers)

# ===== Set output directories for targets to (staged) installation directory ===== #
set_target_properties(OpenXLSX PROPERTIES LIBRARY_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/lib)
set_target_properties(OpenXLSX PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)

# ===== Copy header files to (staged) install directory ===== #
file(REMOVE_RECURSE ${OPENXLSX_INSTALLDIR}/include/)
file(COPY ${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers/ DESTINATION ${OPENXLSX_INSTALLDIR}/include/OpenXLSX)

#=======================================================================================================================
# Install OpenXLSX Library
#=======================================================================================================================
#include(GNUInstallDirs)
#install(TARGETS OpenXLSX
#        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}${INSTALL_SUFFIX}
#        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}${INSTALL_SUFFIX}
#        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR})
#install(FILES ${OPENXLSX_HEADERS} DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX)

