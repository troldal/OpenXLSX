#=======================================================================================================================
# EXTERNAL LIBRARIES
#   Define external libraries used by OpenXLSX. The libraries (Zippy and PugiXML) are header-only, so
#   INTERFACE libraries should be defined.
#=======================================================================================================================
add_library(Zippy INTERFACE IMPORTED)
target_include_directories(Zippy SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/external/zippy/>)

add_library(PugiXML INTERFACE IMPORTED)
target_include_directories(PugiXML SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/external/pugixml/>)
if (${ENABLE_COMPACT_MODE})
    target_compile_definitions(PugiXML INTERFACE PUGIXML_COMPACT)
endif ()

add_library(NoWide INTERFACE IMPORTED)
target_include_directories(NoWide SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/external/nowide/>)

if (${ENABLE_UNICODE_FILENAMES})
    add_compile_definitions(UNICODE_FILENAMES_ENABLED)
endif ()


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
# PROJECT FILES
#   List of project source files
#=======================================================================================================================
set(OPENXLSX_SOURCES
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlFile.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLContentTypes.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLProperties.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLDocument.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLRelationships.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLSharedStrings.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorkbook.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLColumn.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLColor.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLRow.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLSheet.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLCell.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellRange.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellValue.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellReference.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlData.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellIterator.cpp
    ${CMAKE_CURRENT_LIST_DIR}/sources/XLZipArchive.cpp
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
                               PUBLIC
                               $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
                               $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                               $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)     # For export header
    target_link_libraries(OpenXLSX-object PRIVATE Zippy PugiXML NoWide)
    target_compile_definitions(OpenXLSX-object PRIVATE OpenXLSX_shared_EXPORTS)
endif ()

#=======================================================================================================================
# STATIC LIBRARY
#   Define the static library
#=======================================================================================================================
add_library(OpenXLSX-static STATIC "")
target_sources(OpenXLSX-static PRIVATE ${OPENXLSX_SOURCES})
target_include_directories(OpenXLSX-static
                           PUBLIC
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)     # For export header
target_link_libraries(OpenXLSX-static PRIVATE Zippy PugiXML NoWide)
target_compile_definitions(OpenXLSX-static PRIVATE OPENXLSX_STATIC_DEFINE)

#=======================================================================================================================
# SHARED LIBRARY
#   Define the shared library
#=======================================================================================================================
add_library(OpenXLSX-shared SHARED "")
target_sources(OpenXLSX-shared PRIVATE ${OPENXLSX_SOURCES})
target_include_directories(OpenXLSX-shared
                           PUBLIC
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
                           $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)     # For export header
target_link_libraries(OpenXLSX-shared PRIVATE Zippy PugiXML NoWide)

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
# COMPILER FLAGS
#   Set compiler debug flags for GCC, Clang and MSVC.
#=======================================================================================================================
list(APPEND OPENXLSX_DEBUG_FLAGS_GNU
     "-Wmisleading-indentation"
     "-Wduplicated-cond"
     "-Wduplicated-branches"
     "-Wlogical-op"
     "-Wnull-dereference"
     "-Wuseless-cast")
list(APPEND OPENXLSX_DEBUG_FLAGS_GNUCLANG
     "-Wall"
     "-Wextra"
     "-Wshadow"
     "-Wnon-virtual-dtor"
     "-Wold-style-cast"
     "-Wcast-align"
     "-Wunused"
     "-Woverloaded-virtual"
     "-Wpedantic"
     "-Wconversion"
     "-Wsign-conversion"
     "-Wdouble-promotion"
     "-Wformat=2"
     "-Weffc++"
     "-Wno-unknown-pragmas")
list(APPEND OPENXLSX_DEBUG_FLAGS_MSVC
     "/permissive"
     "/W4"
     "/w14242"
     "/w14254"
     "/w14263"
     "/w14265"
     "/w14287"
     "/we4289"
     "/w14296"
     "/w14311"
     "/w14545"
     "/w14546"
     "/w14547"
     "/w14549"
     "/w14555"
     "/w14619"
     "/w14640"
     "/w14826"
     "/w14905"
     "/w14906"
     "/w14928"
     "/wd4251"
     "/wd4275")
list(APPEND OPENXLSX_RELEASE_FLAGS_MSVC
     "/wd4251"
     "/wd4275")

if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    target_compile_options(OpenXLSX-shared PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNU}>)
    target_compile_options(OpenXLSX-static PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNU}>)
    if (${BUILD_OBJECT_LIB} AND (NOT CMAKE_GENERATOR STREQUAL Xcode))
        target_compile_options(OpenXLSX-object PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNU}>)
    endif ()

elseif (("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang") OR
("${CMAKE_CXX_COMPILER_ID}" STREQUAL "AppleClang") OR
("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU"))
    target_compile_options(OpenXLSX-shared PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNUCLANG}>)
    target_compile_options(OpenXLSX-static PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNUCLANG}>)
    if (${BUILD_OBJECT_LIB} AND (NOT CMAKE_GENERATOR STREQUAL Xcode))
        target_compile_options(OpenXLSX-object PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNUCLANG}>)
    endif ()

elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
    string(REGEX REPLACE "/W[3|4]" "" CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS}")
    target_compile_options(OpenXLSX-shared PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_MSVC}>)
    target_compile_options(OpenXLSX-static PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_MSVC}>)
    target_compile_options(OpenXLSX-shared PRIVATE $<$<CONFIG:Release>:${OPENXLSX_RELEASE_FLAGS_MSVC}>)
    target_compile_options(OpenXLSX-static PRIVATE $<$<CONFIG:Release>:${OPENXLSX_RELEASE_FLAGS_MSVC}>)
    if (${BUILD_OBJECT_LIB} AND (NOT CMAKE_GENERATOR STREQUAL Xcode))
        target_compile_options(OpenXLSX-object PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_MSVC}>)
        target_compile_options(OpenXLSX-object PRIVATE $<$<CONFIG:Release>:${OPENXLSX_RELEASE_FLAGS_MSVC}>)
    endif ()
endif ()

#=======================================================================================================================
# Enable Link-Time Optimization
#=======================================================================================================================
if (CMAKE_BUILD_TYPE STREQUAL "Release")
    include(CheckIPOSupported)
    check_ipo_supported(RESULT result OUTPUT output)
    if (result)
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
        TARGETS OpenXLSX-static OpenXLSX-shared
        EXPORT OpenXLSXTargets
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
        COMPONENT lib
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
        COMPONENT lib
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
        COMPONENT bin
        INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX
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



