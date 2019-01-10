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
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLAbstractXMLFile_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLAppProperties_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLColor_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLContentTypes_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCoreProperties_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLDocument_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLException_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLFont_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLRelationships_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSharedStrings_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSpreadsheetElement_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLStyles_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLTokenizer_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLWorkbook_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLChartsheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLColumn_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLRow_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLTemplate_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLWorksheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCell_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellFormats_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellFormula_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellIterator_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellRange_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellReference_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellType_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellValue_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLXml_Impl.h)

set(OPENXLSX_SOURCES
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLAbstractXMLFile_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLAppProperties_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColor_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLContentTypes_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCoreProperties_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLDocument_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLFont_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRelationships_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSharedStrings_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSpreadsheetElement_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLStyles_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLTokenizer_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorkbook_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLChartsheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColumn_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRow_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorksheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCell_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellFormats_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellFormula_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellIterator_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellRange_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellReference_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellValue_Impl.cpp)

#=======================================================================================================================
# Build OpenXLSX library
#=======================================================================================================================
target_sources(OpenXLSX PUBLIC ${OPENXLSX_HEADERS} ${OPENXLSX_SOURCES})
target_link_libraries(OpenXLSX PUBLIC pugixml zip++ zip z bz2)
target_link_directories(OpenXLSX PUBLIC ${OPENXLSX_DEPENDENCIES}/lib)
target_include_directories(OpenXLSX SYSTEM PUBLIC ${OPENXLSX_DEPENDENCIES}/include)
target_include_directories(OpenXLSX PRIVATE ${CMAKE_CURRENT_LIST_DIR}/headers
                                    SYSTEM INTERFACE ${CMAKE_CURRENT_LIST_DIR}/headers)

set_target_properties(OpenXLSX PROPERTIES LIBRARY_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/lib)
set_target_properties(OpenXLSX PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${OPENXLSX_INSTALLDIR}/bin)
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

