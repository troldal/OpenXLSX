#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
set(OPENXLSX_MAJOR_VERSION 0)
set(OPENXLSX_MINOR_VERSION 4)
set(OPENXLSX_MICRO_VERSION 2)
project(OpenXLSX.Library VERSION 0.4.2 LANGUAGES CXX)
set(CMAKE_DEBUG_POSTFIX d)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

#=======================================================================================================================
# Set project metadata
#=======================================================================================================================
set(OPENXLSX_PROJECT_VENDOR "Kenneth Troldal Balslev")
set(OPENXLSX_PROJECT_CONTACT "kenneth.balslev@gmail.com")
set(OPENXLSX_PROJECT_URL "https://github.com/troldal/OpenXLSX")
set(PROJECT_DESCRIPTION "A C++17 library for reading, writing and modifying Excel spreadsheets")

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(IGNORE_ME ${CMAKE_C_COMPILER}) # Suppress CMake warning message

#=======================================================================================================================
# Add build options
#=======================================================================================================================
option(OPENXLSX_ENABLE_LTO "Enables Link-Time Optimization (LTO)" ON)
if (BUILD_SHARED_LIBS)
    set(OPENXLSX_LIBRARY_TYPE "SHARED" CACHE STRING "Set the library type to SHARED")
else()
    set(OPENXLSX_LIBRARY_TYPE "STATIC" CACHE STRING "Set the library type to STATIC")
endif()

#=======================================================================================================================
# EXTERNAL LIBRARIES
#   Define external libraries used by OpenXLSX. The libraries (Zippy, PugiXML, and NoWide) are header-only, so
#   INTERFACE libraries should be defined.
#=======================================================================================================================

#if (OPENXLSX_ENABLE_NOWIDE)
#    if (OPENXLSX_ENABLE_LIBBOOST_NOWIDE)
#        find_package(Boost 1.74 REQUIRED COMPONENTS nowide) # on Linux at least, this requires libboost-nowide-dev to be installed
#    else()
#        add_library(NoWide INTERFACE IMPORTED)
#        target_include_directories(NoWide SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/external/nowide/>)
#    endif()
#endif()

# if (OPENXLSX_ENABLE_LIBZIP)
#     find_package(LibZip 1.7
#         CONFIG
#         COMPONENTS zip
#         QUIET OPTIONAL_COMPONENTS zipcmp zipmerge ziptool # requires the following tools to be installed: zipcmp zipmerge ziptool
#     )
# else()
#     CPMFindPackage(
#             NAME miniz
#             GITHUB_REPOSITORY richgel999/miniz
#             GIT_TAG 3.0.2
#     )
# endif ()

#if (OPENXLSX_ENABLE_LIBPUGIXML)
#    find_package(PugiXML 1.14 REQUIRED)
#else()
#    add_library(PugiXML INTERFACE IMPORTED)
#    target_include_directories(PugiXML SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/external/pugixml/>)
#endif()

#if (${OPENXLSX_COMPACT_MODE})
#    # NOTE: compact mode currently makes no sense if system library pugixml is used
#    if (NOT OPENXLSX_ENABLE_LIBPUGIXML)
#        target_compile_definitions(PugiXML INTERFACE PUGIXML_COMPACT)
#    endif()
#endif ()

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
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCell.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellIterator.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellRange.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellReference.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellValue.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColor.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColumn.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLComments.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLContentTypes.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLDateTime.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLDocument.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLDrawing.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLFormula.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLMergeCells.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLProperties.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRelationships.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRow.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRowData.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSharedStrings.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLStyles.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLTables.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorkbook.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlData.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlFile.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLXmlParser.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLZipArchive.cpp
        )


#=======================================================================================================================
# STATIC AND SHARED LIBRARY
#   Check that the input is valid
#=======================================================================================================================
if(NOT "${OPENXLSX_LIBRARY_TYPE}" STREQUAL "STATIC" AND NOT "${OPENXLSX_LIBRARY_TYPE}" STREQUAL "SHARED")
    message( FATAL_ERROR "Invalid library type. Must be SHARED or STATIC." )
endif()

#=======================================================================================================================
# STATIC LIBRARY
#   Define the static library
#=======================================================================================================================
if ("${OPENXLSX_LIBRARY_TYPE}" STREQUAL "STATIC")
    add_library(OpenXLSX STATIC "")
    add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)
    target_sources(OpenXLSX PRIVATE ${OPENXLSX_SOURCES})
    target_include_directories(OpenXLSX
            PUBLIC
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
            ${Boost_INCLUDE_DIR}    # just in case
            ${NOWIDE_INCLUDE_DIR}   # includes {Boost_INCLUDE_DIR}/boost
            $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)     # For export header
#    if (OPENXLSX_ENABLE_LIBPUGIXML)
#        target_link_libraries(OpenXLSX
#                PRIVATE
#                pugixml::pugixml)
#    else()
        # pugixml is used in header only mode - so whatever this does...
        target_link_libraries(OpenXLSX
                PRIVATE
                pugixml::static)
#    endif()
    if (OPENXLSX_ENABLE_LIBZIP)
        target_link_libraries(OpenXLSX
                PRIVATE
                libzip::zip)
    else ()
        target_link_libraries(OpenXLSX
                PRIVATE
#                $<BUILD_INTERFACE:Zippy>
                miniz)
    endif ()

    if (OPENXLSX_ENABLE_NOWIDE)
#        if (OPENXLSX_ENABLE_LIBBOOST_NOWIDE)
#            target_link_libraries(OpenXLSX
#                    PRIVATE
#                    Boost::nowide)
#        else()
            # boost nowide is header-only, a "library" only makes sense in the context of the nowide header files shipped with OpenXLSX
#            target_link_libraries(OpenXLSX
#                    PRIVATE
#                    $<BUILD_INTERFACE:NoWide>)
            target_link_libraries(OpenXLSX
                    PRIVATE
                    ${NOWIDE_TARGET})
                    # nowide::nowide)
#        endif()
    endif ()

    target_compile_definitions(OpenXLSX PUBLIC OPENXLSX_STATIC_DEFINE)

endif ()

#=======================================================================================================================
# SHARED LIBRARY
#   Define the shared library
#=======================================================================================================================
if ("${OPENXLSX_LIBRARY_TYPE}" STREQUAL "SHARED")
    add_library(OpenXLSX SHARED "")
    add_library(OpenXLSX::OpenXLSX ALIAS OpenXLSX)
    target_sources(OpenXLSX PRIVATE ${OPENXLSX_SOURCES})
    target_include_directories(OpenXLSX
            PUBLIC
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
            ${Boost_INCLUDE_DIR}    # just in case
            ${NOWIDE_INCLUDE_DIR}   # includes {Boost_INCLUDE_DIR}/boost
            $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)     # For export header
#    if (OPENXLSX_ENABLE_LIBPUGIXML)
#        target_link_libraries(OpenXLSX
#                PRIVATE
#                pugixml::pugixml)
#    else()
        # pugixml is used in header only mode - so whatever this does...
        target_link_libraries(OpenXLSX
                PRIVATE
                pugixml::shared)
#    endif()
    if (OPENXLSX_ENABLE_LIBZIP)
        target_link_libraries(OpenXLSX
                PRIVATE
                libzip::zip)
    else ()
        target_link_libraries(OpenXLSX
                PRIVATE
#                $<BUILD_INTERFACE:Zippy>
                miniz)
    endif ()

    if (OPENXLSX_ENABLE_NOWIDE)
#        if (OPENXLSX_ENABLE_LIBBOOST_NOWIDE)
#            target_link_libraries(OpenXLSX
#                    PRIVATE
#                    Boost::nowide)
#        else()
            # boost nowide is header-only, a "library" only makes sense in the context of the nowide header files shipped with OpenXLSX
#            target_link_libraries(OpenXLSX
#                    PRIVATE
#                    $<BUILD_INTERFACE:NoWide>)
            target_link_libraries(OpenXLSX
                    PRIVATE
                    ${NOWIDE_TARGET})
                    # nowide::nowide)
#        endif()
    endif ()

    # Enable Link-Time Optimization (LTO)
    if(${OPENXLSX_ENABLE_LTO})
        include(CheckIPOSupported)
        check_ipo_supported(RESULT result OUTPUT output)
        if (result)
            set_property(TARGET OpenXLSX PROPERTY INTERPROCEDURAL_OPTIMIZATION TRUE)
        endif ()
    endif()

endif()


if (OPENXLSX_ENABLE_NOWIDE)
    target_compile_definitions(OpenXLSX PRIVATE ENABLE_NOWIDE)
#    if (OPENXLSX_ENABLE_LIBBOOST_NOWIDE)
#        # when libboost-nowide is used, a preprocessor directive is needed to change include path and namespace
#        target_compile_definitions(OpenXLSX PRIVATE ENABLE_LIBBOOST_NOWIDE)
#    endif ()
endif ()
#if (OPENXLSX_ENABLE_LIBPUGIXML)
#    target_compile_definitions(OpenXLSX PRIVATE USE_LIBPUGIXML)
#endif ()
if (OPENXLSX_ENABLE_LIBZIP)
    target_compile_definitions(OpenXLSX PRIVATE USE_LIBZIP)
endif ()

# Generate export header
include(GenerateExportHeader)
generate_export_header(OpenXLSX
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
        "-Wnull-dereference")
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
        "-Wdouble-promotion"
        "-Wformat=2"
        "-Weffc++"
        "-Wno-unknown-pragmas")
list(APPEND OPENXLSX_DEBUG_FLAGS_MSVC
        "/permissive-"
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
        "/permissive-"
        "/wd4251"
        "/wd4275")

if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    target_compile_options(OpenXLSX PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNU}>)

elseif (("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang") OR
("${CMAKE_CXX_COMPILER_ID}" STREQUAL "AppleClang") OR
("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU"))
    target_compile_options(OpenXLSX PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_GNUCLANG}>)

elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
    string(REGEX REPLACE "/W[3|4]" "" CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS}")
    target_compile_options(OpenXLSX PRIVATE $<$<CONFIG:Debug>:${OPENXLSX_DEBUG_FLAGS_MSVC}>)
    target_compile_options(OpenXLSX PRIVATE $<$<CONFIG:Release>:${OPENXLSX_RELEASE_FLAGS_MSVC}>)
endif ()

#=======================================================================================================================
# Install
#=======================================================================================================================
# Some basic stuff we'll need in this section
include(GNUInstallDirs)
include(CMakePackageConfigHelpers)
set(ConfigPackageLocation ${CMAKE_INSTALL_LIBDIR}/cmake/OpenXLSX)

# Install interface headers
install(
        FILES ${OPENXLSX_CXX_INTERFACE_HEADERS}
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX/${dir}
)

# Install export header
install(
        FILES ${CMAKE_CURRENT_BINARY_DIR}/OpenXLSX-Exports.hpp
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX/headers
)

file(GLOB OPENXLSX_HEADER_LIST ${CMAKE_CURRENT_LIST_DIR}/headers/*.hpp)
install(
        FILES ${OPENXLSX_HEADER_LIST}
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX/headers
)

install(
        FILES ${CMAKE_CURRENT_LIST_DIR}/OpenXLSX.hpp
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX
)


# Targets
set(CMAKE_INSTALL_TARGETS OpenXLSX ${PUGIXML_INSTALL_TARGET})

if (NOT OPENXLSX_ENABLE_LIBZIP)
    set(CMAKE_INSTALL_TARGETS ${CMAKE_INSTALL_TARGETS} ${ZIP_LIBRARY_INSTALL_TARGET})
endif ()

if (OPENXLSX_ENABLE_NOWIDE)
    set(CMAKE_INSTALL_TARGETS ${CMAKE_INSTALL_TARGETS} ${NOWIDE_INSTALL_TARGET})
endif ()

install(
        TARGETS ${CMAKE_INSTALL_TARGETS}
        EXPORT OpenXLSXTargets
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
        COMPONENT lib
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
        COMPONENT lib
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
        COMPONENT bin
        INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX
)

set_target_properties(OpenXLSX PROPERTIES FOLDER libs VERSION ${OPENXLSX_MAJOR_VERSION}.${OPENXLSX_MINOR_VERSION}.${OPENXLSX_MICRO_VERSION} SOVERSION ${OPENXLSX_MAJOR_VERSION})

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
