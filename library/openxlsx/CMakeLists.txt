#=======================================================================================================================
# Download and define PugiXML library target
#=======================================================================================================================
file(MAKE_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML)
if(NOT EXISTS ${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML/pugiconfig.hpp)
    message(STATUS "PugiXML config file not found. Downloading...")
    file(DOWNLOAD
            https://raw.githubusercontent.com/zeux/pugixml/master/src/pugiconfig.hpp
            ${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML/pugiconfig.hpp
            TIMEOUT 60)
endif()

if(NOT EXISTS ${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML/pugixml.hpp)
    message(STATUS "PugiXML header file not found. Downloading...")
    file(DOWNLOAD
            https://raw.githubusercontent.com/zeux/pugixml/master/src/pugixml.hpp
            ${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML/pugixml.hpp
            TIMEOUT 60)
endif()

if(NOT EXISTS ${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML/pugixml.cpp)
    message(STATUS "PugiXML source file not found. Downloading...")
    file(DOWNLOAD
            https://raw.githubusercontent.com/zeux/pugixml/master/src/pugixml.cpp
            ${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML/pugixml.cpp
            TIMEOUT 60)
endif()

add_library(PugiXML INTERFACE)
add_library(PugiXML::PugiXML ALIAS PugiXML)
target_include_directories(PugiXML SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}/_deps/PugiXML>)
target_compile_definitions(PugiXML INTERFACE PUGIXML_NO_XPATH PUGIXML_HEADER_ONLY)

#=======================================================================================================================
# Define library targets
#=======================================================================================================================
add_library(OpenXLSX-objs OBJECT "")
add_library(OpenXLSX-static STATIC "")
add_library(OpenXLSX-shared SHARED "")

#=======================================================================================================================
# Feature Checks
#=======================================================================================================================
include(CheckCXXSourceCompiles)
check_cxx_source_compiles("
                          #include <array>
                          #include <string>
                          #include <charconv>
                          
                          int main() {
                                  std::array<char, 7> str {};
                                  auto p = std::to_chars(str.data(), str.data() + str.size(), 12345).ptr;
                                  auto strResult = std::string(str.data(), p - str.data());
                                  
                                  unsigned long value = 0;
                                  std::from_chars(strResult.data(), strResult.data() + strResult.size(), value);

                                  return 0;
                          }" CHARCONV_RESULT)

if (CHARCONV_RESULT)
    target_compile_definitions(OpenXLSX-objs PRIVATE CHARCONV_ENABLED)
endif ()

#=======================================================================================================================
# Set compiler flags
#=======================================================================================================================
if (CMAKE_BUILD_TYPE STREQUAL "Debug")
    if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
        #target_compile_options(OpenXLSX PRIVATE -Wlifetime)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        target_compile_options(OpenXLSX-objs PRIVATE
                -Wmisleading-indentation
                -Wduplicated-cond
                -Wduplicated-branches
                -Wlogical-op
                -Wnull-dereference
                -Wuseless-cast)
    elseif (("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang") OR ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU"))
        target_compile_options(OpenXLSX-objs PRIVATE
                -Wall
                -Wextra
                -Wshadow
                -Wnon-virtual-dtor
                -Wold-style-cast
                -Wcast-align
                -Wunused
                -Woverloaded-virtual
                -Wpedantic
                -Wconversion
                -Wsign-conversion
                -Wdouble-promotion
                -Wformat=2
                -Weffc++)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Intel")
        # using Intel C++
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
        target_compile_options(OpenXLSX-objs PRIVATE
                /permissive
                /W4
                /w14242
                /w14254
                /w14263
                /w14265
                /w14287
                /we4289
                /w14296
                /w14311
                /w14545
                /w14546
                /w14547
                /w14549
                /w14555
                /w14619
                /w14640
                /w14826
                /w14905
                /w14906
                /w14928)
    endif ()
endif ()

#=======================================================================================================================
# Find project files
#=======================================================================================================================
set(OPENXLSX_IMPL_HEADERS
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLAbstractXMLFile_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLAppProperties_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLContentTypes_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCoreProperties_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLDocument_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLRelationships_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLSharedStrings_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLTokenizer_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLWorkbook_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLChartsheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLColumn_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLColor_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLRow_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLSheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLTemplate_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLWorksheet_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCell_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellRange_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellReference_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellType_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLCellValue_Impl.h
        ${CMAKE_CURRENT_LIST_DIR}/implementation/headers/XLXml_Impl.h
        )

set(OPENXLSX_IMPL_SOURCES
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLAbstractXMLFile_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLAppProperties_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLContentTypes_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCoreProperties_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLDocument_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLRelationships_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLSharedStrings_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLTokenizer_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLWorkbook_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLChartsheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLColumn_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLColor_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLRow_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLSheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLWorksheet_Impl.cpp
        ${CMAKE_CURRENT_LIST_DIR}/implementation/sources/XLCell_Impl.cpp
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
# Generate object library
#=======================================================================================================================
target_sources(OpenXLSX-objs
        PRIVATE
        ${OPENXLSX_IMPL_SOURCES}
        ${OPENXLSX_CXX_INTERFACE_SOURCES})
target_link_libraries(OpenXLSX-objs
        PRIVATE
        PugiXML Zippy)
target_include_directories(OpenXLSX-objs
        PRIVATE
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/implementation/headers>
        PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/interfaces/c++/headers>
        $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>      # For export header
        )

target_compile_definitions(OpenXLSX-objs
        PRIVATE
        OpenXLSX_shared_EXPORTS
        )

#=======================================================================================================================
# Generate static library
#=======================================================================================================================
target_link_libraries(OpenXLSX-static
        PUBLIC
        OpenXLSX-objs
        )
target_compile_definitions(OpenXLSX-static
        PRIVATE
        OPENXLSX_STATIC_DEFINE
        )

#=======================================================================================================================
# Generate shared library
#=======================================================================================================================
target_link_libraries(OpenXLSX-shared
        PUBLIC
        OpenXLSX-objs
        )

#=======================================================================================================================
# Enable Link-Time Optimization
#=======================================================================================================================
if (CMAKE_BUILD_TYPE STREQUAL "Release")
    include(CheckIPOSupported)
    check_ipo_supported(RESULT result OUTPUT output)
    if (result)
        set_property(TARGET OpenXLSX-static PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
        set_property(TARGET OpenXLSX-shared PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
    endif ()
endif()


#=======================================================================================================================
# Enable code coverage
#=======================================================================================================================
if(ENABLE_COVERAGE AND CMAKE_BUILD_TYPE STREQUAL "Debug")
    if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang" AND NOT WIN32)
        target_compile_options(OpenXLSX-objs PRIVATE -fprofile-instr-generate -fcoverage-mapping)
        target_link_options(OpenXLSX-objs PRIVATE -fprofile-instr-generate)
        #Uncomment in case of linker errors
        #link_libraries(clang_rt.profile-x86_64)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        target_compile_options(OpenXLSX-objs PRIVATE --coverage)
        target_link_options(OpenXLSX-objs PRIVATE --coverage)
        #Uncomment in case of linker errors
        #link_libraries(gcov)
    endif ()
endif()

#=======================================================================================================================
# Generate shared library export header with compiler-specific keywords
#=======================================================================================================================
include(GenerateExportHeader)
generate_export_header(OpenXLSX-shared
        BASE_NAME openxlsx
        EXPORT_FILE_NAME openxlsx_export.h
        )

#=======================================================================================================================
# Install
#=======================================================================================================================
# Some basic stuff we'll need in this section
include(GNUInstallDirs)
include(CMakePackageConfigHelpers)
set(ConfigPackageLocation ${CMAKE_INSTALL_LIBDIR}/cmake/openxlsx)

# Install interface headers
install(
        FILES ${OPENXLSX_CXX_INTERFACE_HEADERS}
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/openxlsx/${dir}
)

# Install export header
install(
        FILES ${CMAKE_CURRENT_BINARY_DIR}/openxlsx_export.h
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/openxlsx
)

# Targets
install(
        TARGETS
        OpenXLSX-objs
        OpenXLSX-static
        OpenXLSX-shared
        PugiXML
        Zippy
        EXPORT OpenXLSXTargets
        LIBRARY
        DESTINATION ${CMAKE_INSTALL_LIBDIR}
        COMPONENT lib
        ARCHIVE
        DESTINATION ${CMAKE_INSTALL_LIBDIR}
        COMPONENT lib
        RUNTIME
        DESTINATION ${CMAKE_INSTALL_BINDIR}
        COMPONENT bin
        INCLUDES
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX
)

# Package version
write_basic_package_version_file(
        "${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX/OpenXLSXConfigVersion.cmake"
        VERSION ${OpenXLSX_VERSION}
        COMPATIBILITY AnyNewerVersion
)
install(
        FILES
        OpenXLSXConfig.cmake
        "${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX/OpenXLSXConfigVersion.cmake"
        DESTINATION ${ConfigPackageLocation}
)

# Package configuration
configure_file(OpenXLSXConfig.cmake
        "${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX/OpenXLSXConfig.cmake"
        COPYONLY
        )

# Package export targets
export(
        EXPORT OpenXLSXTargets
        FILE "${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX/OpenXLSXTargets.cmake"
        NAMESPACE OpenXLSX::
)
install(
        EXPORT OpenXLSXTargets
        FILE OpenXLSXTargets.cmake
        NAMESPACE OpenXLSX::
        DESTINATION ${ConfigPackageLocation}
)



