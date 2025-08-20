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
option(OPENXLSX_CREATE_DOCS "Build library documentation (requires Doxygen and Graphviz/Dot to be installed)" ON)
option(OPENXLSX_BUILD_SAMPLES "Build sample programs" ON)
option(OPENXLSX_BUILD_TESTS "Build and run library tests" ON)
option(OPENXLSX_BUILD_BENCHMARKS "Build and run library benchmarks" OFF)
#option(OPENXLSX_ENABLE_LIBPUGIXML "Enable using libpugixml" OFF)
option(OPENXLSX_ENABLE_LIBZIP "Enable using libzip" OFF)
option(OPENXLSX_FORCE_NOWIDE "Force use of boost nowide, even on non-Windows systems (testing)" OFF)
#option(OPENXLSX_ENABLE_LIBBOOST_NOWIDE "Enable using libboost-nowide" OFF)
option(OPENXLSX_ENABLE_LIBZIP_EXAMPLE "Build Demo1A example using libzip - requires libbz2-dev" OFF)
option(OPENXLSX_SHARED_LIBRARY "Builds the shared library" OFF)
option(OPENXLSX_COMPACT_MODE "Build library in compact mode (slower, but uses less memory)" OFF)
#=======================================================================================================================
# Add project configurations (#ifdef / #ifndef testable)
#=======================================================================================================================
if(WIN32 OR ${OPENXLSX_FORCE_NOWIDE})
    set(OPENXLSX_ENABLE_NOWIDE ON)
endif()


# Requires CMake 3.16+ (for robust target property usage); 3.21+ recommended.
function(make_object_facade ORIG_TARGET OUT_OBJ_TARGET)
    if (NOT TARGET ${ORIG_TARGET})
        message(FATAL_ERROR "make_object_facade: target '${ORIG_TARGET}' not found")
    endif()

    set(OBJ "${OUT_OBJ_TARGET}")
    add_library(${OBJ} OBJECT)  # empty for now

    # 1) Reuse the upstream target's sources.
    # This picks up sources added via add_library()/target_sources().
    # (Generated sources still work if the subproject defines proper custom commands.)
    target_sources(${OBJ} PRIVATE "$<TARGET_PROPERTY:${ORIG_TARGET},SOURCES>")

    # 2) Mimic compile features/options/defs and include dirs (both PRIVATE & INTERFACE).
    target_compile_features(${OBJ} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},COMPILE_FEATURES>")
    target_compile_options(${OBJ} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},COMPILE_OPTIONS>"
            "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_COMPILE_OPTIONS>")
    target_compile_definitions(${OBJ} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},COMPILE_DEFINITIONS>"
            "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_COMPILE_DEFINITIONS>")
    target_include_directories(${OBJ} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},INCLUDE_DIRECTORIES>"
            "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")

    # 3) Position-independent code: reflect upstream if set, or respect your global policy.
    get_target_property(_pic ${ORIG_TARGET} POSITION_INDEPENDENT_CODE)
    if (DEFINED _pic)
        set_property(TARGET ${OBJ} PROPERTY POSITION_INDEPENDENT_CODE ${_pic})
    endif()

    # 4) Name it for the caller.
    set(${OUT_OBJ_TARGET} ${OBJ} PARENT_SCOPE)
endfunction()



#if(OPENXLSX_ENABLE_LIBZIP)
##    CPMFindPackage(
##            NAME ZLIB
##            GITHUB_REPOSITORY madler/zlib
##            GIT_TAG v1.3.1
##    )
##    message(zlib: ${ZLIB_SOURCE_DIR})
##    add_library (ZLIB::ZLIB ALIAS zlibstatic)
#
#find_package(zlib REQUIRED)
#
##
##    if(ZLIB_ADDED)
##        target_include_directories(zlib
##                PUBLIC $<BUILD_INTERFACE:${zlib_BINARY_DIR}>
##                PUBLIC $<INSTALL_INTERFACE:include>)
##        target_include_directories(zlibstatic
##                PUBLIC $<BUILD_INTERFACE:${zlib_BINARY_DIR}>
##                PUBLIC $<INSTALL_INTERFACE:include>)
##
##        message(STATUS "Cmake added local zlib: ${zlib_SOURCE_DIR}")
##    else()
##        message(FATAL_ERROR "Cmake failed to add local zlib")
##    endif()
#
#    CPMFindPackage(
#            NAME libzip
#            GITHUB_REPOSITORY nih-at/libzip
#            GIT_TAG v1.11.4
#    )
#endif()

if(OPENXLSX_ENABLE_NOWIDE)
    CPMFindPackage(
            NAME nowide
            GITHUB_REPOSITORY boostorg/nowide
            GIT_TAG standalone
            OPTIONS
            "NOWIDE_INSTALL ON"
    )
    message(nowide: ${nowide_SOURCE_DIR})
endif()

CPMFindPackage(
        NAME pugixml
        GITHUB_REPOSITORY zeux/pugixml
        GIT_TAG v1.15
        OPTIONS
        "PUGIXML_COMPACT ${OPENXLSX_COMPACT_MODE}"
)

## Suppose the zlib target is 'zlibstatic' (common on Windows builds) or 'zlib'
if (TARGET pugixml-static)
    make_object_facade(pugixml-static pugixml_obj)
else()
    message(FATAL_ERROR "Couldn't find a pugixml target to wrap.")
endif()



message(pugixml: ${pugixml_SOURCE_DIR})



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


