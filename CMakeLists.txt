#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.26 FATAL_ERROR)
project(OpenXLSX VERSION 0.4.2 LANGUAGES C CXX)

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
include(${CMAKE_CURRENT_LIST_DIR}/cmake/MakeObjectLib.cmake)
include(${CMAKE_CURRENT_LIST_DIR}/cmake/AddDependency.cmake)
include(${CMAKE_CURRENT_LIST_DIR}/cmake/MessageColor.cmake)

#=======================================================================================================================
# Output Directories.
#=======================================================================================================================
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)

#=======================================================================================================================
# Add build options
#=======================================================================================================================
option(BUILD_SHARED_LIBS "Builds the shared library" OFF) # replaced OPENXLSX_SHARED_LIBRARY with cmake option reserved for that purpose
option(OPENXLSX_BUILD_SAMPLES         "Build sample programs" ON)
option(OPENXLSX_BUILD_TESTS           "Build and run library tests" ON) # TBD: are these still functional?
option(OPENXLSX_CREATE_DOCS           "Build library documentation (requires Doxygen and Graphviz/Dot to be installed)" ON)
option(OPENXLSX_BUILD_BENCHMARKS      "Build and run library benchmarks" OFF)
option(OPENXLSX_FORCE_NOWIDE          "Force use of boost nowide, even on non-Windows systems (testing)" OFF)
#option(OPENXLSX_ENABLE_LIBZIP_EXAMPLE "Build Demo1A example using libzip - requires libbz2-dev" ON)
option(OPENXLSX_COMPACT_MODE          "Build library in compact mode (slower, but uses less memory)" OFF)

# Dependency resolution mode: AUTO | LOCAL | REMOTE
set(DEP_MODE "AUTO" CACHE STRING "Dependency resolution mode: AUTO | LOCAL | REMOTE")
set_property(CACHE DEP_MODE PROPERTY STRINGS AUTO LOCAL REMOTE)

# Dependency inclusion options: BUNDLE_REMOTE_DEPS only makes sense if DEP_MODE is AUTO or REMOTE
if(DEP_MODE STREQUAL "AUTO" OR DEP_MODE STREQUAL "REMOTE")
    option(BUNDLE_REMOTE_DEPS "Build REMOTE dependencies as object libraries and bundle into main library" OFF)
endif()

# zip library selection: libzip (requires additional dependencies) or miniz (default, header-only, no additional dependencies)
if(NOT WIN32)
    option(OPENXLSX_ENABLE_LIBZIP "Enable using libzip" OFF) # when OFF, default zip library is miniz. NOTE: libzip currently does NOT support CPM_DOWNLOAD_ALL
endif()

#=======================================================================================================================
# Initialize dependency accumulators
#=======================================================================================================================

set(OPENXLSX_DEPENDENCY_PACKAGES "")
set(OPENXLSX_DEPENDENCY_OBJECTS "")
set(OPENXLSX_DEPENDENCY_HEADERS "")
set(OPENXLSX_DEPENDENCY_LIBS "")

if(BUILD_SHARED_LIBS)
    set(CMAKE_POSITION_INDEPENDENT_CODE ON) # needed for shared libraries
endif()

#=======================================================================================================================
# Add Boost.Nowide dependency (if needed)
#=======================================================================================================================
if(WIN32 OR ${OPENXLSX_FORCE_NOWIDE})
    set(OPENXLSX_ENABLE_NOWIDE ON)
endif()

if(OPENXLSX_ENABLE_NOWIDE)

    # Check if package exists without errors or full loading
    find_package(Boost 1.73 COMPONENTS nowide QUIET)
    if (Boost_FOUND)
        list(APPEND OPENXLSX_DEPENDENCY_LIBS "$<BUILD_LOCAL_INTERFACE:Boost::nowide>")
        set(NOWIDE_TARGET Boost::nowide)
    else ()

        set(INSTALL_NOWIDE_TARGET OFF)
        if(NOT BUNDLE_REMOTE_DEPS)
            set(INSTALL_NOWIDE_TARGET ON)
        endif()

        add_dependency(
                NAME nowide
                TARGET nowide::nowide
                FIND_PACKAGE_NAME nowide
                VERSION 11.3.0
                GIT_REPOSITORY https://github.com/boostorg/nowide.git
                GIT_TAG standalone
                OPTIONS "NOWIDE_INSTALL ${INSTALL_NOWIDE_TARGET}"
                PACKAGE_LIST OPENXLSX_DEPENDENCY_PACKAGES
                OBJECT_LIST OPENXLSX_DEPENDENCY_OBJECTS
                HEADER_LIST OPENXLSX_DEPENDENCY_HEADERS
                LIBRARY_LIST OPENXLSX_DEPENDENCY_LIBS
        )
        set(NOWIDE_TARGET nowide::nowide)
    endif ()
endif()

#=======================================================================================================================
# Add zip library dependency
#=======================================================================================================================
if (OPENXLSX_ENABLE_LIBZIP)
    add_dependency(
            NAME libzip
            TARGET libzip::zip
            VERSION 1.11.3
            GIT_REPOSITORY https://github.com/nih-at/libzip.git
            GIT_TAG v1.11.3
            OPTIONS "BUILD_SHARED_LIBS ON LIBZIP_DO_INSTALL ON"
            PACKAGE_LIST OPENXLSX_DEPENDENCY_PACKAGES
            OBJECT_LIST OPENXLSX_DEPENDENCY_OBJECTS
            HEADER_LIST OPENXLSX_DEPENDENCY_HEADERS
            LIBRARY_LIST OPENXLSX_DEPENDENCY_LIBS
    )
else()
    find_package(miniz CONFIG QUIET)
    if (miniz_FOUND AND NOT ${DEP_MODE} STREQUAL "REMOTE")
        set(MINIZ_TARGET miniz::miniz)
    else()
        set(MINIZ_TARGET miniz)
    endif()

    add_dependency(
            NAME miniz
            TARGET ${MINIZ_TARGET}
            FIND_PACKAGE_NAME miniz
            FIND_PACKAGE_OPTIONS CONFIG
            VERSION 3.0.2
            GIT_REPOSITORY https://github.com/richgel999/miniz.git
            GIT_TAG 3.0.2
            OPTIONS "AMALGAMATE_SOURCES ON BUILD_HEADER_ONLY OFF"
            PACKAGE_LIST OPENXLSX_DEPENDENCY_PACKAGES
            OBJECT_LIST OPENXLSX_DEPENDENCY_OBJECTS
            HEADER_LIST OPENXLSX_DEPENDENCY_HEADERS
            LIBRARY_LIST OPENXLSX_DEPENDENCY_LIBS
    )

endif ()

#=======================================================================================================================
# Add pugixml dependency
#=======================================================================================================================
if (NOT ${DEP_MODE} STREQUAL "REMOTE")
    find_package(pugixml CONFIG QUIET)
endif()

if (pugixml_FOUND)
    set(PUGIXML_TARGET pugixml)
elseif(BUILD_SHARED_LIBS)
    set(PUGIXML_TARGET pugixml-shared)
else()
    set(PUGIXML_TARGET pugixml-static)
endif()

add_dependency(
        NAME pugixml
        TARGET ${PUGIXML_TARGET}
        FIND_PACKAGE_NAME pugixml
        FIND_PACKAGE_OPTIONS CONFIG
        VERSION 1.14
        GIT_REPOSITORY https://github.com/zeux/pugixml.git
        GIT_TAG v1.14
        OPTIONS "PUGIXML_COMPACT ${OPENXLSX_COMPACT_MODE}"
        PACKAGE_LIST OPENXLSX_DEPENDENCY_PACKAGES
        OBJECT_LIST OPENXLSX_DEPENDENCY_OBJECTS
        HEADER_LIST OPENXLSX_DEPENDENCY_HEADERS
        LIBRARY_LIST OPENXLSX_DEPENDENCY_LIBS
)

#message(AUTHOR_WARNING "Dependency packages: ${OPENXLSX_DEPENDENCY_PACKAGES}")
#message(AUTHOR_WARNING "Dependency objects: ${OPENXLSX_DEPENDENCY_OBJECTS}")
#message(AUTHOR_WARNING "Dependency headers: ${OPENXLSX_DEPENDENCY_HEADERS}")
#message(AUTHOR_WARNING "Dependency libraries: ${OPENXLSX_DEPENDENCY_LIBS}")


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
#message( VERBOSE        "   ++ NOWIDE_INSTALL_TARGET:       ${NOWIDE_INSTALL_TARGET}" )
message( VERBOSE        "   ++ ZIP_LIBRARY_INSTALL_TARGET:  ${ZIP_LIBRARY_INSTALL_TARGET}" )
#message( VERBOSE        "   ++ PUGIXML_INSTALL_TARGET:      ${PUGIXML_INSTALL_TARGET}" )
