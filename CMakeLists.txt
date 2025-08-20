#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX VERSION 0.4.2 LANGUAGES CXX)

set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(IGNORE_ME ${CMAKE_C_COMPILER}) # Suppress CMake warning message

#=======================================================================================================================
# Include CPM package manager
#=======================================================================================================================
include(${CMAKE_CURRENT_LIST_DIR}/cmake/CPM.cmake)

#=======================================================================================================================
# Output Directories.
#=======================================================================================================================
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)

#=======================================================================================================================
# Add build options
#=======================================================================================================================
option(OPENXLSX_CREATE_DOCS           "Build library documentation (requires Doxygen and Graphviz/Dot to be installed)" ON)
option(OPENXLSX_BUILD_SAMPLES         "Build sample programs" ON)
option(OPENXLSX_BUILD_TESTS           "Build and run library tests" ON) # TBD: are these still functional?
option(OPENXLSX_BUILD_BENCHMARKS      "Build and run library benchmarks" OFF)
option(OPENXLSX_ENABLE_LIBZIP         "Enable using libzip" OFF) # when OFF, default zip library is miniz. NOTE: libzip currently does NOT support CPM_DOWNLOAD_ALL
option(OPENXLSX_FORCE_NOWIDE          "Force use of boost nowide, even on non-Windows systems (testing)" OFF)
option(OPENXLSX_ENABLE_LIBZIP_EXAMPLE "Build Demo1A example using libzip - requires libbz2-dev" OFF)
option(         BUILD_SHARED_LIBS     "Builds the shared library" OFF) # replaced OPENXLSX_SHARED_LIBRARY with cmake option reserved for that purpose
option(OPENXLSX_COMPACT_MODE          "Build library in compact mode (slower, but uses less memory)" OFF)

# workaround option for CPM_LOCAL_PACKAGES_ONLY
option(OPENXLSX_CPM_LOCAL_PACKAGES_ONLY "CPM shall only use packages already installed on the host OS, otherwise fail" OFF)
if(${OPENXLSX_CPM_LOCAL_PACKAGES_ONLY})
    set(CPM_LOCAL_PACKAGES_ONLY ON) # Defining this directly as an option in this file doesn't seem to work, as stupid as it is
    set(NOWIDE_TARGET Boost::nowide)
    set(NOWIDE_INSTALL_TARGET "")  # TBD: none needed?
    set(PUGIXML_INSTALL_TARGET "") # TBD: none needed?
else()
    set(CPM_DOWNLOAD_ALL ON) # Force download so that a local package is not used while a component is missing (e.g. Boost::nowide)
    set(NOWIDE_TARGET nowide::nowide)
    set(NOWIDE_INSTALL_TARGET nowide)
    set(PUGIXML_INSTALL_TARGET pugixml)
endif()

#=======================================================================================================================
# Add project configurations (#ifdef / #ifndef testable)
#=======================================================================================================================
if(WIN32 OR ${OPENXLSX_FORCE_NOWIDE})
    set(OPENXLSX_ENABLE_NOWIDE ON)
endif()

if(OPENXLSX_ENABLE_NOWIDE)
    CPMFindPackage(
            NAME Boost
            VERSION 1.73.0
            GITHUB_REPOSITORY boostorg/nowide
            # NAME nowide
            GIT_TAG standalone
            OPTIONS
            "NOWIDE_INSTALL ON"
            FIND_PACKAGE_ARGUMENTS "COMPONENTS nowide"
    )
    if(NOT "${Boost_INCLUDE_DIR}" STREQUAL "")
        set(NOWIDE_INCLUDE_DIR "${Boost_INCLUDE_DIR}/boost")
    endif()
endif()

if (OPENXLSX_ENABLE_LIBZIP)
    # NOTE: libzip 1.11.4 is the newest version as of 2025-08-16 and not available on most host systems, so a lower version tag should be used.
    #        However, 1.11.4 is the first version that allows building libzip separately from the (unneeded) components zipcmp zipmerge ziptool
    #        Maybe check for available version & if 1.11.4 is available, exclude unneeded components - or does this happen automatically?
    find_package(LibZip 1.7
        CONFIG
        COMPONENTS zip
        QUIET OPTIONAL_COMPONENTS zipcmp zipmerge ziptool # requires the following tools to be installed: zipcmp zipmerge ziptool
    )
    set(ZIP_LIBRARY_INSTALL_TARGET "") # TBD: none needed?
    # TODO: Help needed - how to make CPMFindPackage work with the libzip repo?
    # CPMFindPackage(
    #         NAME libzip
    #         GITHUB_REPOSITORY nih-at/libzip
    #         GIT_TAG 1.11.4
    # )
    # set(ZIP_LIBRARY_INSTALL_TARGET libzip)
else()
    CPMFindPackage(
            NAME miniz
            GITHUB_REPOSITORY richgel999/miniz
            GIT_TAG 3.0.2
            # use of reserved BUILD_SHARED_LIBS automatically passes this option through
    )
    set(ZIP_LIBRARY_INSTALL_TARGET miniz)
endif ()

CPMFindPackage(
        NAME pugixml
        GITHUB_REPOSITORY zeux/pugixml
        GIT_TAG v1.15
        OPTIONS
        "PUGIXML_COMPACT ${OPENXLSX_COMPACT_MODE}"
        # use of reserved BUILD_SHARED_LIBS automatically passes this option through
)


#=======================================================================================================================
# Add project subdirectories
#=======================================================================================================================
add_subdirectory(OpenXLSX)

if(${OPENXLSX_CREATE_DOCS})
    add_subdirectory(Documentation)
endif()

if(${OPENXLSX_BUILD_SAMPLES})
    add_subdirectory(Examples)
endif()

if(${OPENXLSX_BUILD_TESTS})
    add_subdirectory(Tests)
endif()

if(${OPENXLSX_BUILD_BENCHMARKS})
    add_subdirectory(Benchmarks)
endif()


#=======================================================================================================================
# Summarize OpenXLSX configuration
#=======================================================================================================================
message(     NOTICE     "-- OpenXLSX Configuration:" )
message(     NOTICE     "   ++ Build library documentation:       ${OPENXLSX_CREATE_DOCS}" )
message(     NOTICE     "   ++ Build sample programs:             ${OPENXLSX_BUILD_SAMPLES}" )
message(     NOTICE     "   ++ Build and run library tests:       ${OPENXLSX_BUILD_TESTS}" )
message(     NOTICE     "   ++ Build and run library benchmarks:  ${OPENXLSX_BUILD_BENCHMARKS}" )
message(     NOTICE     "   ++ Build Demo1A example using libzip: ${OPENXLSX_ENABLE_LIBZIP_EXAMPLE}" )
if(OPENXLSX_ENABLE_LIBZIP_EXAMPLE AND NOT OPENXLSX_BUILD_SAMPLES)
    message( WARNING    "Demo1A will not be built because sample programs are disabled!" )
endif()
    message( NOTICE     "   ++ Build library in compact mode:     ${OPENXLSX_COMPACT_MODE}" )
    message( NOTICE     "   ++" )

if(BUILD_SHARED_LIBS)
    message( NOTICE     "   ++ Library type:                shared" )
else()
    message( NOTICE     "   ++ Library type:                static" )
endif()

if(OPENXLSX_ENABLE_NOWIDE)
    message( NOTICE     "   ++ Unicode support:             Boost nowide" )
    if("${Boost_INCLUDE_DIR}" STREQUAL "")
        message( NOTICE "   ++ Boost_INCLUDE_DIR:           (empty)" )
    else()
        message( NOTICE "   ++ Boost_INCLUDE_DIR:           ${Boost_INCLUDE_DIR}" )
        message( NOTICE "   ++ NOWIDE_INCLUDE_DIR:          ${NOWIDE_INCLUDE_DIR}" )
    endif()
else()
    message( NOTICE     "   ++ Unicode support:             (OS) native" )
endif()

if (OPENXLSX_ENABLE_LIBZIP)
    message( NOTICE     "   ++ zip library:                 libzip" )
else()
    message( NOTICE     "   ++ zip library:                 miniz" )
endif()
if(CPM_LOCAL_PACKAGES_ONLY)
    message( NOTICE     "   ++ CPM package source:          CPM_LOCAL_PACKAGES_ONLY" )
endif()
if(CPM_DOWNLOAD_ALL)
    message( NOTICE     "   ++ CPM package source:          CPM_DOWNLOAD_ALL" )
endif()

# NOTE: for some reason, cmake with --log-level VERBOSE (or DEBUG) prefixes message output with "-- "
message( VERBOSE        "   ++ pugixml_SOURCE_DIR:          ${pugixml_SOURCE_DIR}" )
message( VERBOSE        "   ++ NOWIDE_INSTALL_TARGET:       ${NOWIDE_INSTALL_TARGET}" )
message( VERBOSE        "   ++ ZIP_LIBRARY_INSTALL_TARGET:  ${ZIP_LIBRARY_INSTALL_TARGET}" )
message( VERBOSE        "   ++ PUGIXML_INSTALL_TARGET:      ${PUGIXML_INSTALL_TARGET}" )
