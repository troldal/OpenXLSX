#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
set(OPENXLSX_MAJOR_VERSION 0)
set(OPENXLSX_MINOR_VERSION 5)
set(OPENXLSX_MICRO_VERSION 0)
project(OpenXLSX VERSION 0.5.0 LANGUAGES CXX)

set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)
set(IGNORE_ME ${CMAKE_C_COMPILER}) # Suppress CMake warning message

# ============================================================================
# Policy Settings
# ============================================================================
set(CMAKE_POLICY_DEFAULT_CMP0077 NEW) # Respect BUILD_SHARED_LIBS / suppress warning in nowide - TBD if this is unproblematic
if(POLICY CMP0077)
    cmake_policy(SET CMP0077 NEW)  # Respect BUILD_SHARED_LIBS
endif()

set(CMAKE_POLICY_DEFAULT_CMP0167 NEW)
if(POLICY CMP0167)
    cmake_policy(SET CMP0167 NEW) # find the upstream BoostConfig.cmake directly
endif()

#=======================================================================================================================
# Build options
#=======================================================================================================================
# 2026-04-19 working configurations:
# - static library build (BUILD_SHARED_LIBS OFF)
# - shared library build (BUILD_SHARED_LIBS ON)
# - either setting for OPENXLSX_LOCAL_PACKAGES_ONLY (OFF = automatically fetch dependencies, ON = error out if dependency is not available locally)
# - OpenXLSX and dependencies are installed into /usr/local/lib/OpenXLSX subfolder to avoid version conflicts for parallel installations of zip library, pugixml, nowide
# - pugixml, libzip & miniz package config (.pc) files are now part of the installation
# - dependency package config files installed by OpenXLSX will be installed as pugixml-OpenXLSX.pc, libzip-OpenXLSX.pc, miniz-OpenXLSX.pc respectively to avoid conflicts
# - dependency package config files should behave well with existing versions of the dependencies (OpenXLSX package config file gives precedence to the self-installed dependencies)
# - OpenXLSX package config file now provides the runtime path for an executable linked against OpenXLSX shared libraries in /usr/local/lib/OpenXLSX
# - vcpkg configuration / install via vcpkg
# - TBD: static/shared library build on Windows 11 with clang_cl, MSVC, gcc

# 2026-04-19 remaining issues:
# - TBD for OPENXLSX_MONOLITHIC_LIBRARY=ON, use target_link_interface instead of target_link_library for remaining dependencies (libzip/miniz, pugixml) and modify installation files
#   NOTE: the workaround of target_link_interface is currently unable to link in subdependencies and accordingly fails for libzip::zip. Accordingly, use of
#          target_link_interface (for anything but nowide) is descoped now. The only downside (which isn't a problem) is that dependencies are also install targets
#
option(OPENXLSX_CREATE_DOCS           "Build library documentation (requires Doxygen and Graphviz/Dot to be installed)" ON)
option(OPENXLSX_BUILD_SAMPLES         "Build sample programs" ON)
option(OPENXLSX_BUILD_TESTS           "Build and run library tests - currently not functional" OFF)
option(OPENXLSX_BUILD_BENCHMARKS      "Build and run library benchmarks - currently not functional" OFF)
option(OPENXLSX_ENABLE_LIBZIP         "Enable using libzip" OFF) # when OFF, default zip library is miniz
option(OPENXLSX_ENABLE_LIBZIP_EXAMPLE "Build Demo1A example using libzip - requires libbz2-dev and liblzma-dev" OFF)
option(OPENXLSX_FORCE_NOWIDE          "Force use of boost nowide, even on non-Windows systems (testing)" OFF)
option(OPENXLSX_NOWIDE_STANDALONE     "use nowide standalone (default source: github)" ON)
# Generic library build settings
option(         BUILD_SHARED_LIBS     "Builds the shared library" OFF) # cmake option reserved for that purpose
option(OPENXLSX_MONOLITHIC_LIBRARY    "(static library only) Link library dependencies into a bundled OpenXLSX library file" OFF)
option(OPENXLSX_LOCAL_PACKAGES_ONLY   "only use dependencies already installed on the host OS, otherwise fail" OFF)
option(OPENXLSX_COMPACT_MODE          "Build library in compact mode (slower, but uses less memory)" OFF)

#=======================================================================================================================
# Options for tool manage_dependency
#=======================================================================================================================
# These options should not be manipulated directly, but rather handled through OpenXLSX options
#                                                       # Default: all OFF
option(PREFER_STATIC   "Prefer static linking over shared libraries"   OFF)     # Prefer static linking over shared libraries
option(USE_SYSTEM_LIBS "Use system-installed libraries when available" OFF)     # Use system-installed libraries when available
option(FETCH_DEPS_AUTO "Automatically fetch missing dependencies"      OFF)     # Automatically fetch missing dependencies
# NOTE: with default settings, FETCH_DEPS_AUTO will get enabled because USE_SYSTEM_LIBS is OFF
option(FORCE_FETCH_ALL "Ignore system libraries and fetch all deps"    OFF)     # Ignore system libraries and fetch all deps

#=======================================================================================================================
# Manage configurations
#=======================================================================================================================
if(OPENXLSX_COMPACT_MODE)
    message( WARNING "setting OPENXLSX_COMPACT_MODE is currently not communicated to dependencies (pugixml) - they will be built normally - this is on the TO-DO list" )
endif()


if(OPENXLSX_MONOLITHIC_LIBRARY)
    # message( FATAL_ERROR "OPENXLSX_MONOLITHIC_LIBRARY is currently not supported - this is a TBD feature" )

    set(PREFER_STATIC ON CACHE BOOL "" FORCE)
    if(BUILD_SHARED_LIBS)
        message( FATAL_ERROR "Can not use OPENXLSX_MONOLITHIC_LIBRARY together with BUILD_SHARED_LIBS" )
    endif()
endif()

if(FORCE_FETCH_ALL)
    if(OPENXLSX_LOCAL_PACKAGES_ONLY)
        message( FATAL_ERROR "Can not use OPENXLSX_LOCAL_PACKAGES_ONLY together with FORCE_FETCH_ALL" )
    endif()
    set(USE_SYSTEM_LIBS OFF CACHE BOOL "" FORCE)
    set(FETCH_DEPS_AUTO ON  CACHE BOOL "" FORCE)
    message(STATUS "FORCE_FETCH_ALL enabled: ignoring system libraries")
else()
    if(OPENXLSX_LOCAL_PACKAGES_ONLY)
        if(FETCH_DEPS_AUTO)
            message( WARNING "Use of OPENXLSX_LOCAL_PACKAGES_ONLY will override FETCH_DEPS_AUTO" )
        endif()
        set(USE_SYSTEM_LIBS ON  CACHE BOOL "" FORCE)
        set(FETCH_DEPS_AUTO OFF CACHE BOOL "" FORCE)
        set(FORCE_FETCH_ALL OFF CACHE BOOL "" FORCE)
    else()
        set(USE_SYSTEM_LIBS ON  CACHE BOOL "" FORCE)
        set(FETCH_DEPS_AUTO ON  CACHE BOOL "" FORCE)
        set(FORCE_FETCH_ALL OFF CACHE BOOL "" FORCE)
    endif()
endif()

if(OPENXLSX_NOWIDE_STANDALONE AND NOT FETCH_DEPS_AUTO)
    if(WIN32 OR ${OPENXLSX_FORCE_NOWIDE})
        # message(FATAL_ERROR "OPENXLSX_NOWIDE_STANDALONE selected but FETCH_DEPS_AUTO is disabled - can not pull in standalone repository")
        message(WARNING "OPENXLSX_NOWIDE_STANDALONE selected but FETCH_DEPS_AUTO is disabled - if the installed version of nowide is *not* the standalone version, this will result in errors")
    else()
        # suppress error if nowide is not going to be used
    endif()
    # TODO TBD: accept standalone local installation
endif()


#  NOTE: when building an executable against the *shared object* OpenXLSX library, /usr/local/lib will not be in the (ldconfig) default libraries search path
#        and needs to be configured, e.g. by LD_LIBRARY_PATH=/usr/local/lib ./a.out
if(${OPENXLSX_MONOLITHIC_LIBRARY} AND ${BUILD_SHARED_LIBS})
    message( FATAL_ERROR "Can not use option OPENXLSX_MONOLITHIC_LIBRARY with shared library" )
endif()

if(${OPENXLSX_MONOLITHIC_LIBRARY} AND ${OPENXLSX_LOCAL_PACKAGES_ONLY})
    message( WARNING "option OPENXLSX_MONOLITHIC_LIBRARY with OPENXLSX_LOCAL_PACKAGES_ONLY may fail when locally a installed dependency is not available as a static library" )
endif()

if(${OPENXLSX_LOCAL_PACKAGES_ONLY} AND NOT ${OPENXLSX_ENABLE_LIBZIP})
    message( WARNING "option OPENXLSX_LOCAL_PACKAGES_ONLY without option OPENXLSX_ENABLE_LIBZIP may fail unless libminiz is preinstalled on the system" )
endif()

if(${OPENXLSX_MONOLITHIC_LIBRARY})
    message( NOTICE "NOTE: bundled static library libOpenXLSX-bundled.a will NOT be installed to the target system by this build configuration" )
endif()


#=======================================================================================================================
# Configure a clean OPENXLSX library name available to the subfolder CMakeLists.txt
#=======================================================================================================================
set(OPENXLSX_LIBRARY_NAME ${PROJECT_NAME})

#=======================================================================================================================
# Attempt to locate the git executable
#=======================================================================================================================
include( FindGit )
if(NOT "${GIT_EXECUTABLE}" STREQUAL "")
    message( DEBUG "GIT_EXECUTABLE is ${GIT_EXECUTABLE}" )
else()
    message( WARNING "could not locate GIT_EXECUTABLE - will be unable to fetch missing dependencies" )
endif()

#=======================================================================================================================
# Include tools to manage dependencies
#=======================================================================================================================
include(${CMAKE_CURRENT_LIST_DIR}/cmake/manage_dependency.cmake)

#=======================================================================================================================
# Include tool to get git commit metadata
#=======================================================================================================================
include(${CMAKE_CURRENT_LIST_DIR}/cmake/get-git-HEAD-ID.cmake)

#=======================================================================================================================
# Output Directories.
#=======================================================================================================================
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/output)

#=======================================================================================================================
# Add project configurations (#ifdef / #ifndef testable)
#=======================================================================================================================
if(WIN32 OR ${OPENXLSX_FORCE_NOWIDE})
    set(OPENXLSX_ENABLE_NOWIDE ON)
endif()

# Handle support for Boost::nowide
if(OPENXLSX_ENABLE_NOWIDE)

    # If nowide is installed locally:
    set( NOWIDE_TYPICAL_NAMES )                                       # to be used with find_library in determining whether a static library is installed locally
    list( APPEND NOWIDE_TYPICAL_NAMES "boost_nowide" )                # CAUTION: NOWIDE_TYPICAL_NAMES *must* be a list (semicolon-separated)
    list( APPEND NOWIDE_TYPICAL_NAMES "libboost_nowide" )             #
    list( APPEND NOWIDE_TYPICAL_NAMES "nowide" )                      #

    if(OPENXLSX_NOWIDE_STANDALONE)
        set( NOWIDE_PACKAGE_NAME "nowide" )
        set( NOWIDE_COMPONENTS "" )
        #
        set( NOWIDE_ALIAS "nowide::nowide" )
        set( NOWIDE_TARGET "nowide" )
        set( NOWIDE_TARGET_SYSTEM "nowide::nowide" )    # this is the target expected from find_package
        # set( NOWIDE_GIT_REPO_PARAM "https://codeberg.org/lars_uffmann/nowide_standalone.git" )
        # set( NOWIDE_GIT_TAG_PARAM  "v11.3.1-fixed" )
        set( NOWIDE_URL_PARAM "https://github.com/boostorg/nowide/releases/download/v11.3.1/nowide_standalone_v11.3.1.tar.gz" )
        set( NOWIDE_URL_HASH_PARAM "SHA256=eaec4d331e3961f5eeb10c46a11691d62047900a7a40765b0f23cdd3181e6ca6" )
    else()
        set( NOWIDE_PACKAGE_NAME "Boost" )
        set( NOWIDE_COMPONENTS "COMPONENTS;nowide" )
        #
        set( NOWIDE_ALIAS "Boost::nowide" )
        set( NOWIDE_TARGET "boost_nowide" )
        set( NOWIDE_TARGET_SYSTEM "Boost::nowide" )     # this is the target expected from find_package
        set( NOWIDE_GIT_REPO_PARAM "https://github.com/boostorg/nowide.git" )
        set( NOWIDE_GIT_TAG_PARAM  "v11.3.1" )
    endif()

    if(PREFER_STATIC)
        set(Boost_USE_STATIC_LIBS ON)   # prefer/require static boost libraries
    endif()

    if(OPENXLSX_MONOLITHIC_LIBRARY)
        set(Boost_NOWIDE_INSTALL OFF CACHE BOOL "") # this CACHE instruction appears to be the only way to set that option for the subproject
        set(      NOWIDE_INSTALL OFF CACHE BOOL "") # standalone version has a different option name
    else()
        set(Boost_NOWIDE_INSTALL  ON CACHE BOOL "") #
        set(      NOWIDE_INSTALL  ON CACHE BOOL "") #
    endif()

    manage_dependency(
        LIB_NAME       Nowide
        PACKAGE_NAME   ${NOWIDE_PACKAGE_NAME}
        COMPONENTS     ${NOWIDE_COMPONENTS}
        TYPICAL_NAMES  ${NOWIDE_TYPICAL_NAMES}

        TARGET_NAME        ${NOWIDE_TARGET}
        TARGET_NAME_SYSTEM ${NOWIDE_TARGET_SYSTEM}

        GIT_REPOSITORY ${NOWIDE_GIT_REPO_PARAM}
        GIT_TAG        ${NOWIDE_GIT_TAG_PARAM}

        # a provided URL & URL_HASH takes precedence over GITHUB_REPO or GIT_REPOSITORY:
        URL            ${NOWIDE_URL_PARAM}
        URL_HASH       ${NOWIDE_URL_HASH_PARAM}
    )
    message( STATUS "manage_dependency reported Nowide_PROVIDED_TARGET as ${Nowide_PROVIDED_TARGET}, Nowide_INSTALL_TARGET as ${Nowide_INSTALL_TARGET}" )
    set(NOWIDE_TARGET ${Nowide_PROVIDED_TARGET} CACHE STRING "" FORCE)
endif()

if (OPENXLSX_ENABLE_LIBZIP)
    # NOTE: libzip 1.11.4 is the newest version as of 2025-08-16 and not available on most host systems, so a lower version tag should be used.
    #        However, 1.11.4 is the first version that allows building libzip separately from the (unneeded) components zipcmp zipmerge ziptool
    #        Maybe check for available version & if 1.11.4 is available, exclude unneeded components - or does this happen automatically?

    set(ZIP_LIBRARY libzip) # becomes a dependency in package config
    set(ZIP_LIBRARY_ALIAS           "libzip::zip")
    set(ZIP_LIBRARY_TARGET          "libzip::zip")
    set(ZIP_LIBRARY_TARGET_SYSTEM   "libzip::zip")
    manage_dependency(
        LIB_NAME       libzip
        PACKAGE_NAME   LibZip
        # COMPONENTS     zip    # 2026-04-04: it appears libzip does not actually expose COMPONENTS for find_package
        VERSION        1.11.3   # TBD what version is needed
        TYPICAL_NAMES  "libzip" # TBD what these can be

        TARGET_NAME         ${ZIP_LIBRARY_TARGET}
        TARGET_NAME_SYSTEM  ${ZIP_LIBRARY_TARGET_SYSTEM}

        GIT_REPOSITORY  https://github.com/nih-at/libzip.git
        GIT_TAG         v1.11.3
        # OPTIONS "QUIET OPTIONAL_COMPONENTS zipcmp zipmerge ziptool" # "
        # OPTIONS "BUILD_TOOLS OFF"                                   # only relevant for CPMFindPackage (i.e. OPENXLSX_LOCAL_PACKAGES_ONLY=OFF)
        # OPTIONS "BUILD_REGRESS OFF"                                 # "
        # OPTIONS "BUILD_OSSFUZZ OFF"                                 # "
        # OPTIONS "BUILD_EXAMPLES OFF"                                # "
        # OPTIONS "BUILD_DOC OFF"                                     # "
    )
    message( STATUS "manage_dependency reported libzip_PROVIDED_TARGET as ${libzip_PROVIDED_TARGET}, libzip_INSTALL_TARGET as ${libzip_INSTALL_TARGET}" )
    set(ZIP_LIBRARY_PROVIDED_TARGET "${libzip_PROVIDED_TARGET}")
    set(ZIP_LIBRARY_INSTALL_TARGET "${libzip_INSTALL_TARGET}")
else()
    if(WIN32)
        set(OPENXLSX_MINIZ_BUILD_MODE "OFF") # for miniz on windows, the shared library fails to build - so build the shared OpenXLSX library against the static miniz lib in that case
    else()
        set(OPENXLSX_MINIZ_BUILD_MODE "${BUILD_SHARED_LIBS}")
    endif()

    set(ZIP_LIBRARY miniz) # becomes a dependency in package config
    set(ZIP_LIBRARY_ALIAS miniz::miniz)
    set(ZIP_LIBRARY_TARGET miniz)
    set(ZIP_LIBRARY_TARGET_SYSTEM miniz::miniz)
    manage_dependency(
        LIB_NAME      miniz
        PACKAGE_NAME  miniz
        VERSION       3.0.2
        TYPICAL_NAMES "miniz"

        TARGET_NAME         ${ZIP_LIBRARY_TARGET}
        TARGET_NAME_SYSTEM  ${ZIP_LIBRARY_TARGET_SYSTEM}
        GIT_REPOSITORY  https://github.com/richgel999/miniz.git
        GIT_TAG         3.0.2
    )
    message( STATUS "manage_dependency reported miniz_PROVIDED_TARGET as ${miniz_PROVIDED_TARGET}, miniz_INSTALL_TARGET as ${miniz_INSTALL_TARGET}" )
    set(ZIP_LIBRARY_PROVIDED_TARGET "${miniz_PROVIDED_TARGET}")
    set(ZIP_LIBRARY_INSTALL_TARGET "${miniz_INSTALL_TARGET}")
endif ()


if(${BUILD_SHARED_LIBS})
    set(PUGIXML_TARGET pugixml::shared) # actually, just pugixml seems to work fine for a target - but let's be explicit
else()
    set(PUGIXML_TARGET pugixml::static) # see above
endif()

manage_dependency(
    LIB_NAME      pugixml
    PACKAGE_NAME  pugixml
    VERSION       1.15
    TYPICAL_NAMES "pugixml" # TBD what these can be

    TARGET_NAME         ${PUGIXML_TARGET}
    TARGET_NAME_SYSTEM  pugixml
    GIT_REPOSITORY  https://github.com/zeux/pugixml.git
    GIT_TAG         v1.15
    # OPTIONS         "PUGIXML_COMPACT ${OPENXLSX_COMPACT_MODE}" # TBD how to pass on PUGIXML options
)
message( STATUS "manage_dependency reported puxixml_PROVIDED_TARGET as ${pugixml_PROVIDED_TARGET}, pugixml_INSTALL_TARGET as ${pugixml_INSTALL_TARGET}" )
set(PUGIXML_INSTALL_TARGET "${pugixml_INSTALL_TARGET}")
if(pugixml_FETCHED)
    set(PUGIXML_STATIC_LIBRARY "pugixml-static")     # used for OPENXLSX_BUNDLED_STATIC_LIB at the moment
else()
    set(PUGIXML_STATIC_LIBRARY "pugixml")            # used for OPENXLSX_BUNDLED_STATIC_LIB at the moment

    set(PUGIXML_TARGET "pugixml")
    message( STATUS "pugixml was found locally, using PUGIXML_TARGET ${PUGIXML_TARGET}" )
endif()


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
message(     NOTICE     "-- OpenXLSX configuration:" )
message(     NOTICE     "   ++ git HEAD (short):                      ${GIT_COMMIT_ID}" )
message(     NOTICE     "   ++ git HEAD commit timestamp:             ${GIT_COMMIT_ID_TIME}" )
message(     NOTICE     "   ++ Build library documentation:           ${OPENXLSX_CREATE_DOCS}" )
message(     NOTICE     "   ++ Build sample programs:                 ${OPENXLSX_BUILD_SAMPLES}" )
message(     NOTICE     "   ++ Build and run library tests:           ${OPENXLSX_BUILD_TESTS}" )
message(     NOTICE     "   ++ Build and run library benchmarks:      ${OPENXLSX_BUILD_BENCHMARKS}" )
message(     NOTICE     "   ++ Enable using libzip:                   ${OPENXLSX_ENABLE_LIBZIP}" )
message(     NOTICE     "   ++ Build Demo1A example using libzip:     ${OPENXLSX_ENABLE_LIBZIP_EXAMPLE}" )
if(OPENXLSX_ENABLE_LIBZIP_EXAMPLE AND NOT OPENXLSX_BUILD_SAMPLES)
    message( WARNING    "Demo1A will not be built because sample programs are disabled!" )
endif()
message(     NOTICE     "   ++ Force use of boost nowide:             ${OPENXLSX_FORCE_NOWIDE}" )
message(     NOTICE     "   ++ Use nowide standalone library:         ${OPENXLSX_NOWIDE_STANDALONE}" )
message(     NOTICE     "   ++ Build the shared library:              ${BUILD_SHARED_LIBS}" )
message(     NOTICE     "   ++ Make a bundled static library file:    ${OPENXLSX_MONOLITHIC_LIBRARY}" )
message(     NOTICE     "   ++ Only use installed packages:           ${OPENXLSX_LOCAL_PACKAGES_ONLY}" )
message(     NOTICE     "   ++ Build library in compact mode:         ${OPENXLSX_COMPACT_MODE}" )
message(     NOTICE     "   ++" )

if(BUILD_SHARED_LIBS)
    message( NOTICE     "   ++ Library type:                          shared" )
else()
    message( NOTICE     "   ++ Library type:                          static" )
endif()

if(OPENXLSX_ENABLE_NOWIDE)
    message( NOTICE     "   ++ Unicode support:                       Boost nowide" )
    # if("${Boost_INCLUDE_DIR}" STREQUAL "")
    #     message( NOTICE "   ++ Boost_INCLUDE_DIR:                     (empty)" )
    # else()
    #     message( NOTICE "   ++ Boost_INCLUDE_DIR:                     ${Boost_INCLUDE_DIR}" )
    # endif()
    if("${EXAMPLES_NOWIDE_INCLUDE}" STREQUAL "")
        message( NOTICE "   ++ EXAMPLES_NOWIDE_INCLUDE:               (empty)" )
    else()
        message( NOTICE "   ++ EXAMPLES_NOWIDE_INCLUDE:               ${EXAMPLES_NOWIDE_INCLUDE}" )
    endif()
else()
    message( NOTICE     "   ++ Unicode support:                       (OS) native" )
endif()

if (OPENXLSX_ENABLE_LIBZIP)
    message( NOTICE     "   ++ zip library:                           libzip" )
else()
    message( NOTICE     "   ++ zip library:                           miniz" )
endif()

message(     NOTICE     "   manage_dependency configuration:" )
message(     NOTICE     "   ++ PREFER_STATIC:                         ${PREFER_STATIC}" )
message(     NOTICE     "   ++ USE_SYSTEM_LIBS:                       ${USE_SYSTEM_LIBS}" )
message(     NOTICE     "   ++ FETCH_DEPS_AUTO:                       ${FETCH_DEPS_AUTO}" )
message(     NOTICE     "   ++ FORCE_FETCH_ALL:                       ${FORCE_FETCH_ALL}" )

# NOTE: for some reason, cmake with --log-level VERBOSE (or DEBUG) prefixes message output with "-- "
message( VERBOSE        "   VERBOSE configuration information:" )
message( VERBOSE        "   ++ ZIP_LIBRARY_INSTALL_TARGET:            ${ZIP_LIBRARY_INSTALL_TARGET}" )
message( VERBOSE        "   ++ PUGIXML_INSTALL_TARGET:                ${PUGIXML_INSTALL_TARGET}" )


#=======================================================================================================================
# Configure CPack for debian package generation - thank you @ https://github.com/Jacob-Burckhardt for the suggestions
#=======================================================================================================================
configure_file ("${CMAKE_CURRENT_LIST_DIR}/cmake/libopenxlsxCPackOptions.cmake.in"
                "${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/libopenxlsxCPackOptions.cmake"
                @ONLY)

set (CPACK_PROJECT_CONFIG_FILE
     "${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/libopenxlsxCPackOptions.cmake")

# must include the CPack project config file here or the settings won't be used
include ("${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/libopenxlsxCPackOptions.cmake")

# include( CPack ) must always be invoked after all CPACK\_\* variables are defined
include( CPack )
message(     NOTICE     "   ++ CPACK_PACKAGE_VERSION:                 ${CPACK_PACKAGE_VERSION}" )

message( NOTICE     "   ++ Prefer static linking over shared libraries: ${PREFER_STATIC}" )
message( NOTICE     "   ++ Use system-installed libraries when available: ${USE_SYSTEM_LIBS}" )
message( NOTICE     "   ++ Automatically fetch missing dependencies: ${FETCH_DEPS_AUTO}" )
message( NOTICE     "   ++ Ignore system libraries and fetch all deps: ${FORCE_FETCH_ALL}" )

# TODO: how to trigger cpack -G DEB from cmake
