#target_compile_definitions(Zippy INTERFACE MINIZ_HEADER_FILE_ONLY)

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
set(OPENXLSX_SOURCES
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlFile.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLAppProperties.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLContentTypes.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCoreProperties.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLDocument.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRelationships.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSharedStrings.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorkbook.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLChartsheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColumn.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColor.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRow.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorksheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCell.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellRange.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellValue.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellReference.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlParser.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlData.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellIterator.cpp
        )

#=======================================================================================================================
# Generate object library
#=======================================================================================================================
target_sources(OpenXLSX-objs
        PRIVATE
        ${OPENXLSX_SOURCES})
target_include_directories(OpenXLSX-objs
        PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
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
endif ()


#=======================================================================================================================
# Enable code coverage
#=======================================================================================================================
if (ENABLE_COVERAGE AND CMAKE_BUILD_TYPE STREQUAL "Debug")
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
endif ()

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



