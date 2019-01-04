#=======================================================================================================================
# Define library targets
#=======================================================================================================================
add_library(OpenXLSX SHARED "")
add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)

#=======================================================================================================================
# Set compiler flags
#=======================================================================================================================
include(CheckIPOSupported)
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
    target_compile_options(OpenXLSX PRIVATE -Wall)
    target_compile_options(OpenXLSX PRIVATE -Wextra)
    target_compile_options(OpenXLSX PRIVATE -pedantic)
    if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
        target_compile_options(OpenXLSX PRIVATE -Weverything)
        target_compile_options(OpenXLSX PRIVATE -Wno-c++98-compat)
        target_compile_options(OpenXLSX PRIVATE -Wno-c++98-compat-pedantic)
        target_compile_options(OpenXLSX PRIVATE -Wno-documentation)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        target_compile_options(OpenXLSX PRIVATE -Weffc++)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Intel")
        # using Intel C++
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
        # using Visual Studio C++
    endif()
elseif(CMAKE_BUILD_TYPE STREQUAL "Release")
    check_ipo_supported(RESULT result OUTPUT output)
    if(result)
        set_property(TARGET OpenXLSX PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
    endif()
endif()

#=======================================================================================================================
# Find project files
#=======================================================================================================================
set(OPENXLSX_HEADERS
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLAbstractXMLFile.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLAppProperties.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLColor.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLContentTypes.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCoreProperties.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLDocument.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLException.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLFont.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLRelationships.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSharedStrings.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSpreadsheetElement.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLStyles.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLTokenizer.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLWorkbook.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLChartsheet.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLColumn.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLRow.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSheet.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLTemplate.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLWorksheet.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCell.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellFormats.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellFormula.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellIterator.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellRange.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellReference.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellType.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellValue.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLXml.h)

set(OPENXLSX_SOURCES
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLAbstractXMLFile.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLAppProperties.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColor.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLContentTypes.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCoreProperties.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLDocument.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLFont.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRelationships.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSharedStrings.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSpreadsheetElement.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLStyles.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLTokenizer.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorkbook.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLChartsheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColumn.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRow.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorksheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCell.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellFormats.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellFormula.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellIterator.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellRange.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellReference.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellValue.cpp)

#=======================================================================================================================
# Build OpenXLSX library
#=======================================================================================================================
target_sources(OpenXLSX PUBLIC ${OPENXLSX_SOURCES})
target_link_libraries(OpenXLSX PUBLIC z bz2 zip zip++ pugixml)
target_link_directories(OpenXLSX PUBLIC ${OPENXLSX_DEPENDENCIES}/lib)
target_include_directories(OpenXLSX SYSTEM PUBLIC ${OPENXLSX_DEPENDENCIES}/include)
target_include_directories(OpenXLSX PRIVATE ${CMAKE_CURRENT_LIST_DIR}/headers
                                    SYSTEM INTERFACE ${CMAKE_CURRENT_LIST_DIR}/headers)

set_target_properties(OpenXLSX PROPERTIES LIBRARY_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/lib)
file(COPY ${CMAKE_CURRENT_LIST_DIR}/headers/ DESTINATION ${OPENXLSX_INSTALLDIR}/include/OpenXLSX)


#=======================================================================================================================
# Install OpenXLSX Library
#=======================================================================================================================
include(GNUInstallDirs)
install(TARGETS OpenXLSX
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}${INSTALL_SUFFIX}
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}${INSTALL_SUFFIX}
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR})
install(FILES ${OPENXLSX_HEADERS} DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX)

