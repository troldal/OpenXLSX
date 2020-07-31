#=======================================================================================================================
# COMPILER FEATURES
#   Some older C++17 compilers don't support the char_conv features. If the compiler doesn't support it,
#   a less optimal work-around will be used.
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
    add_compile_definitions(CHARCONV_ENABLED)
endif ()

#=======================================================================================================================
# COMPILER FLAGS
#   Set compiler debug flags for GCC, Clang and MSVC for all targets.
#=======================================================================================================================
if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    add_compile_options(
            $<$<CONFIG:Debug>:-Wmisleading-indentation>
            $<$<CONFIG:Debug>:-Wduplicated-cond>
            $<$<CONFIG:Debug>:-Wduplicated-branches>
            $<$<CONFIG:Debug>:-Wlogical-op>
            $<$<CONFIG:Debug>:-Wnull-dereference>
            $<$<CONFIG:Debug>:-Wuseless-cast>)

elseif (("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang") OR
("${CMAKE_CXX_COMPILER_ID}" STREQUAL "AppleClang") OR
("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU"))
    add_compile_options(
            $<$<CONFIG:Debug>:-Wall>
            $<$<CONFIG:Debug>:-Wextra>
            $<$<CONFIG:Debug>:-Wshadow>
            $<$<CONFIG:Debug>:-Wnon-virtual-dtor>
            $<$<CONFIG:Debug>:-Wold-style-cast>
            $<$<CONFIG:Debug>:-Wcast-align>
            $<$<CONFIG:Debug>:-Wunused>
            $<$<CONFIG:Debug>:-Woverloaded-virtual>
            $<$<CONFIG:Debug>:-Wpedantic>
            $<$<CONFIG:Debug>:-Wconversion>
            $<$<CONFIG:Debug>:-Wsign-conversion>
            $<$<CONFIG:Debug>:-Wdouble-promotion>
            $<$<CONFIG:Debug>:-Wformat=2>
            $<$<CONFIG:Debug>:-Weffc++>)

elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
    target_compile_options(
            $<$<CONFIG:Debug>:/permissive>
            $<$<CONFIG:Debug>:/W4>
            $<$<CONFIG:Debug>:/w14242>
            $<$<CONFIG:Debug>:/w14254>
            $<$<CONFIG:Debug>:/w14263>
            $<$<CONFIG:Debug>:/w14265>
            $<$<CONFIG:Debug>:/w14287>
            $<$<CONFIG:Debug>:/we4289>
            $<$<CONFIG:Debug>:/w14296>
            $<$<CONFIG:Debug>:/w14311>
            $<$<CONFIG:Debug>:/w14545>
            $<$<CONFIG:Debug>:/w14546>
            $<$<CONFIG:Debug>:/w14547>
            $<$<CONFIG:Debug>:/w14549>
            $<$<CONFIG:Debug>:/w14555>
            $<$<CONFIG:Debug>:/w14619>
            $<$<CONFIG:Debug>:/w14640>
            $<$<CONFIG:Debug>:/w14826>
            $<$<CONFIG:Debug>:/w14905>
            $<$<CONFIG:Debug>:/w14906>
            $<$<CONFIG:Debug>:/w14928>)
endif ()

#=======================================================================================================================
# PROJECT FILES
#   List of project source files
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
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlData.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellIterator.cpp
    )

#=======================================================================================================================
# OBJECT LIBRARY
#   Define the object library.
#   Object provide a simpler way of integrating a library into an application. However, not all build systems (e.g.
#   Xcode) play nice with object libraries. Therefore, an object library will only be built if supported.
#=======================================================================================================================
if (${BUILD_OBJECT_LIB} AND (NOT CMAKE_GENERATOR STREQUAL Xcode))
    add_library(OpenXLSX-object OBJECT "")
    target_sources(OpenXLSX-object PRIVATE ${OPENXLSX_SOURCES})
    target_include_directories(OpenXLSX-object
                               PRIVATE
                               $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                               $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>     # For export header
                               SYSTEM INTERFACE
                               $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                               $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)      # For export header
    #target_compile_definitions(OpenXLSX-object PRIVATE OpenXLSX_shared_EXPORTS)
endif ()

#=======================================================================================================================
# STATIC LIBRARY
#   Define the static library
#=======================================================================================================================
add_library(OpenXLSX-static STATIC "")
target_sources(OpenXLSX-static PRIVATE ${OPENXLSX_SOURCES})
target_include_directories(OpenXLSX-static
                           PRIVATE
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>     # For export header
                           SYSTEM INTERFACE
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)      # For export header
target_compile_definitions(OpenXLSX-static PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# SHARED LIBRARY
#   Define the shared library
#=======================================================================================================================
add_library(OpenXLSX-shared SHARED "")
target_sources(OpenXLSX-shared PRIVATE ${OPENXLSX_SOURCES})
target_include_directories(OpenXLSX-shared
                           PRIVATE
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>     # For export header
                           SYSTEM INTERFACE
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)      # For export header

#=======================================================================================================================
# GENERATE EXPORT HEADER
#   Generate preprocessor definitions for setting symbol visibility.
#=======================================================================================================================
include(GenerateExportHeader)
generate_export_header(OpenXLSX-shared
                       BASE_NAME openxlsx
                       EXPORT_FILE_NAME OpenXLSX-Exports.hpp
                       EXPORT_MACRO_NAME OPENXLSX_EXPORT
                       NO_EXPORT_MACRO_NAME OPENXLSX_HIDDEN)

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
        target_compile_options(OpenXLSX-shared PRIVATE -fprofile-instr-generate -fcoverage-mapping)
        target_link_options(OpenXLSX-shared PRIVATE -fprofile-instr-generate)
    elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
        target_compile_options(OpenXLSX-shared PRIVATE --coverage)
        target_link_options(OpenXLSX-shared PRIVATE --coverage)
    endif ()
endif ()


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
        #        OpenXLSX-object
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



