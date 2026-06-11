#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Library VERSION 0.5.0 LANGUAGES CXX)
set(CMAKE_DEBUG_POSTFIX d)
set(CMAKE_CXX_VISIBILITY_PRESET hidden)
set(CMAKE_VISIBILITY_INLINES_HIDDEN YES)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

# Include utility functions
include( ${CMAKE_CURRENT_LIST_DIR}/cmake/bundle_static_library.cmake )
include( ${CMAKE_CURRENT_LIST_DIR}/cmake/patch_pkg_config.cmake )

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

# message( NOTICE "CMAKE_CURRENT_SOURCE_DIR is ${CMAKE_CURRENT_SOURCE_DIR}" )
    target_include_directories(OpenXLSX
            PUBLIC
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>
            $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)     # For export header
    target_include_directories(OpenXLSX
            INTERFACE
            $<INSTALL_INTERFACE:include/OpenXLSX>)              # get install-safe include directories

    target_link_libraries(OpenXLSX
            PRIVATE
            ${PUGIXML_TARGET})

    target_link_libraries(OpenXLSX
            PRIVATE
            ${ZIP_LIBRARY_PROVIDED_TARGET})

    if (OPENXLSX_ENABLE_NOWIDE)
        if(OPENXLSX_MONOLITHIC_LIBRARY)
            # nominal way to link via target_link_libraries: triggers unwanted installation of dependencies
            # accordingly, for monolithic library, use target_link_interface instead of target_link_libraries
            #  to link only the include directories and compile definitions
            target_link_interface(
                LIB_NAME           OpenXLSX
                PRIVATE
                INTERFACE_NAME     nowide_interface
                DEPENDENCY_ALIAS   ${NOWIDE_ALIAS}
                DEPENDENCY_TARGET  ${NOWIDE_TARGET}
            )
            set(NOWIDE_INSTALL_TARGET nowide_interface)
        else()
            target_link_libraries(OpenXLSX
                    PRIVATE
                    ${NOWIDE_ALIAS})
            if(TARGET ${NOWIDE_TARGET})
                # 2026-04-04 get_target_property now accesses NOWIDE_TARGET, was wrongly accessing NOWIDE_ALIAS
                get_target_property(NOWIDE_IS_IMPORTED "${NOWIDE_TARGET}" IMPORTED)
                if(NOT NOWIDE_IS_IMPORTED)
                    set(NOWIDE_INSTALL_TARGET ${NOWIDE_TARGET})
                endif()
            endif()
        endif()
    endif ()

    target_compile_definitions(OpenXLSX PUBLIC OPENXLSX_STATIC_DEFINE)

    # if static combined library shall be built:
    if (OPENXLSX_MONOLITHIC_LIBRARY)
        list(APPEND static_libs "$<TARGET_FILE:OpenXLSX>")
        list(APPEND static_libs_system "")

        # message( NOTICE "ZIP_LIBRARY is ${ZIP_LIBRARY}" )
        if("${ZIP_LIBRARY}" STREQUAL "libzip")
            message( WARNING "building a monolithic OpenXLSX library with libzip is not recommended - dependencies of libzip (-lbz2 -llzma -lzstd -lssl -lcrypto -lz) will not be included. Use miniz for a truly standalone OpenXLSX library" )
        endif()

        list(APPEND static_libs "${pugixml_TARGET_STATIC}")
        list(APPEND static_libs "${${ZIP_LIBRARY}_TARGET_STATIC}")
        if(OPENXLSX_ENABLE_NOWIDE)
            list(APPEND static_libs "${Nowide_TARGET_STATIC}")
        endif()

        # # Output the status of the various library targets:
        # foreach(lib IN LISTS static_libs)
        #     message( NOTICE "OpenXLSX static_libs entry is ${lib}" )
        # endforeach()
        bundle_static_library_simplified( OpenXLSX OpenXLSX-bundled "${static_libs}" "${static_libs_system}" )

        # target_link_libraries(OpenXLSX-bundled INTERFACE OpenXLSX) # NOTE: this might not be needed at all - the bundled library is just an extra build target
        # # TBD: how to correctly define link and include directories associated with OpenXLSX-bundled? The below does not work :(
        # set_target_properties(OpenXLSX-bundled
        #   PROPERTIES
        #     IMPORTED_LOCATION "$<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>/libOpenXLSX-bundled.a"
        #     # INTERFACE_INCLUDE_DIRECTORIES "${CMAKE_SOURCE_DIR} ${CMAKE_SOURCE_DIR}/headers"
        #     INTERFACE_INCLUDE_DIRECTORIES "${CMAKE_CURRENT_LIST_DIR} ${CMAKE_CURRENT_LIST_DIR}/headers"
        #     # INTERFACE_INCLUDE_DIRECTORIES "$<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}> $<BUILD_INTERFACE:${CMAKE_CURRENT_LIST_DIR}/headers>"
        # )
        # get_target_property( includeDirs OpenXLSX-bundled INTERFACE_INCLUDE_DIRECTORIES)
        # message( NOTICE "OpenXLSX-bundled INTERFACE_INCLUDE_DIRECTORIES is ${includeDirs}" )

  endif ()
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
            $<BUILD_INTERFACE:${CMAKE_CURRENT_BINARY_DIR}>)     # For export header
    target_include_directories(OpenXLSX
            INTERFACE
            $<INSTALL_INTERFACE:include/OpenXLSX>)              # get install-safe include directories

    target_link_libraries(OpenXLSX
            PRIVATE
            ${PUGIXML_TARGET})

    target_link_libraries(OpenXLSX
            PRIVATE
            ${ZIP_LIBRARY_PROVIDED_TARGET})

    if (OPENXLSX_ENABLE_NOWIDE)
        target_link_libraries(OpenXLSX
                PRIVATE
                ${NOWIDE_ALIAS})

        if(TARGET ${NOWIDE_TARGET})
            # 2026-04-04 get_target_property now accesses NOWIDE_TARGET, was wrongly accessing NOWIDE_ALIAS
            get_target_property(NOWIDE_IS_IMPORTED "${NOWIDE_TARGET}" IMPORTED)
            if(NOT NOWIDE_IS_IMPORTED)
                set(NOWIDE_INSTALL_TARGET ${NOWIDE_TARGET})
            endif()
        endif()

        # # link to the built archive file so that OpenXLSX does not require NOWIDE_INSTALL_TARGET to be in the export set
        # target_link_libraries(OpenXLSX PRIVATE $<TARGET_FILE:${NOWIDE_INSTALL_TARGET}>)

        # ensure nowide is built before OpenXLSX (so that archive exists at link time)
        add_dependencies(OpenXLSX ${NOWIDE_ALIAS})

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

    # ### ensure transparent include paths for the source code
    #     2026-04-04 get_target_property now accesses NOWIDE_TARGET, was wrongly accessing NOWIDE_ALIAS
    get_target_property(NOWIDE_INCLUDES ${NOWIDE_TARGET} INTERFACE_INCLUDE_DIRECTORIES)     # get nowide include paths
    set(NOWIDE_EXTRA_INCLUDE) # empty initializer
    if(USE_SYSTEM_LIBS OR NOT OPENXLSX_NOWIDE_STANDALONE) # no need to add NOWIDE_EXTRA_INCLUDE for standalone downloads
        list( GET NOWIDE_INCLUDES 0 NOWIDE_INCLUDE )                                        # ensure that NOWIDE_INCLUDE only contains a single directory entry
        set(NOWIDE_EXTRA_INCLUDE "${NOWIDE_INCLUDE}/boost")                                 # and add NOWIDE_EXTRA_INCLUDE with an appended ./boost subdirectory
        target_include_directories(OpenXLSX PRIVATE ${NOWIDE_EXTRA_INCLUDE})                # should be PRIVATE for include directories that users linking against OpenXLSX shall NOT need
        message( NOTICE "NOWIDE_EXTRA_INCLUDE is ${NOWIDE_EXTRA_INCLUDE}" )
    endif()

    # Make determined nowide include path (if any) available to OpenXLSX Examples as explicitly named variable
    set(EXAMPLES_NOWIDE_INCLUDE ${NOWIDE_EXTRA_INCLUDE} CACHE STRING "" FORCE)
endif ()

if (OPENXLSX_ENABLE_LIBZIP)
    target_compile_definitions(OpenXLSX PRIVATE USE_LIBZIP)
else()
    # ### ensure transparent include paths for the source code
    get_target_property(MINIZ_IS_IMPORTED ${ZIP_LIBRARY_PROVIDED_TARGET} IMPORTED)
    if(MINIZ_IS_IMPORTED)
        get_target_property(MINIZ_INCLUDES ${ZIP_LIBRARY_PROVIDED_TARGET} INTERFACE_INCLUDE_DIRECTORIES)     # get miniz include paths
        set(MINIZ_EXTRA_INCLUDE) # empty initializer
        # list( GET MINIZ_INCLUDES 0 MINIZ_INCLUDE )          # ensure that MINIZ_INCLUDE only contains a single directory entry
        # set(MINIZ_EXTRA_INCLUDE "${MINIZ_INCLUDE}/miniz")   # and add MINIZ_EXTRA_INCLUDE with an appended ./miniz subdirectory
        set(MINIZ_EXTRA_INCLUDE "${MINIZ_INCLUDES}")
        target_include_directories(OpenXLSX PRIVATE ${MINIZ_EXTRA_INCLUDE})     # should be PRIVATE for include directories that users linking against OpenXLSX shall NOT need
        message( NOTICE "MINIZ_EXTRA_INCLUDE is ${MINIZ_EXTRA_INCLUDE}" )
    endif()
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

# Install OpenXLSX header files
install(DIRECTORY ${CMAKE_CURRENT_LIST_DIR}/headers/
  DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX/headers
  FILES_MATCHING
    PATTERN "headers/*.hpp"
    PATTERN "headers/detail" EXCLUDE
)

install(
        FILES ${CMAKE_CURRENT_LIST_DIR}/OpenXLSX.hpp
        DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX
)


# Determine OpenXLSX dependency targets, if any
set(CMAKE_INSTALL_TARGETS_DEPS "")
if(NOT "${PUGIXML_INSTALL_TARGET}" STREQUAL "")
    set(CMAKE_INSTALL_TARGETS_DEPS ${CMAKE_INSTALL_TARGETS_DEPS} ${PUGIXML_INSTALL_TARGET})
endif()
if(NOT "${ZIP_LIBRARY_INSTALL_TARGET}" STREQUAL "")
    set(CMAKE_INSTALL_TARGETS_DEPS ${CMAKE_INSTALL_TARGETS_DEPS} ${ZIP_LIBRARY_INSTALL_TARGET})
endif()
if(OPENXLSX_ENABLE_NOWIDE AND NOT "${NOWIDE_INSTALL_TARGET}" STREQUAL "")
    set(CMAKE_INSTALL_TARGETS_DEPS ${CMAKE_INSTALL_TARGETS_DEPS} ${NOWIDE_INSTALL_TARGET})
endif()
string(STRIP "${CMAKE_INSTALL_TARGETS_DEPS}" CMAKE_INSTALL_TARGETS_DEPS)

# Install OpenXLSX dependencies, if any
if(NOT "${CMAKE_INSTALL_TARGETS_DEPS}" STREQUAL "") # TODO: test when all dependencies are available locally
    install(
        TARGETS ${CMAKE_INSTALL_TARGETS_DEPS}
        EXPORT OpenXLSXTargets
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}/OpenXLSX    # shared objects
        COMPONENT lib
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}/OpenXLSX    # static libs
        COMPONENT lib
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}/OpenXLSX    # executables, if any
        COMPONENT bin
        # INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}/OpenXLSX   # not needed due to use of target_include_directories with BUILD_INTERFACE for dependencies
    )
endif()

# OpenXLSX install targets
install(
        TARGETS OpenXLSX
        EXPORT OpenXLSXTargets
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}             # shared objects
        COMPONENT lib
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}             # static libs
        COMPONENT lib
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}             # executables, if any
        COMPONENT bin
        # INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}            # not needed due to use of target_include_directories with BUILD_INTERFACE and INSTALL_INTERFACE
)

set_target_properties(OpenXLSX PROPERTIES FOLDER libs VERSION ${OPENXLSX_MAJOR_VERSION}.${OPENXLSX_MINOR_VERSION}.${OPENXLSX_MICRO_VERSION} SOVERSION ${OPENXLSX_MAJOR_VERSION})
# Ensure that an installed version of OpenXLSX refers to a relative runpath
set_target_properties(OpenXLSX PROPERTIES
  INSTALL_RPATH "\$ORIGIN:\$ORIGIN/OpenXLSX"
  BUILD_RPATH "\$ORIGIN"
  INSTALL_RPATH_USE_LINK_PATH TRUE
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


#=======================================================================================================================
# PkgConfig: Establish a pkg-config file for the library
#=======================================================================================================================
# message( NOTICE "CMAKE_CURRENT_SOURCE_DIR is ${CMAKE_CURRENT_SOURCE_DIR}" )
# message( NOTICE "CMAKE_LIBRARY_OUTPUT_DIRECTORY is ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}" )

# message( NOTICE "PROJECT_NAME is ${PROJECT_NAME}" )
# message( NOTICE "CMAKE_PROJECT_NAME is ${CMAKE_PROJECT_NAME}" )
# message( NOTICE "CMAKE_CURRENT_BINARY_DIR is ${CMAKE_CURRENT_BINARY_DIR}" )
#
# message( NOTICE "CMAKE_INSTALL_PREFIX is ${CMAKE_INSTALL_PREFIX}" )
# message( NOTICE "CMAKE_INSTALL_INCLUDEDIR is ${CMAKE_INSTALL_INCLUDEDIR}" )
# message( NOTICE "PROJECT_DESCRIPTION is ${PROJECT_DESCRIPTION}" )
# message( NOTICE "PROJECT_HOMEPAGE_URL is ${PROJECT_HOMEPAGE_URL}" )
# message( NOTICE "PROJECT_VERSION is ${PROJECT_VERSION}" )
# message( NOTICE "CMAKE_INSTALL_DATAROOTDIR is ${CMAKE_INSTALL_DATAROOTDIR}" )
#
# message( NOTICE "CMAKE_INSTALL_LIBDIR is ${CMAKE_INSTALL_LIBDIR}" )
# message( NOTICE "CMAKE_INSTALL_BINDIR is ${CMAKE_INSTALL_BINDIR}" )

# message( NOTICE ">> CMAKE_CURRENT_LIST_DIR is ${CMAKE_CURRENT_LIST_DIR}" )

# Write the template for the package config file
set(PUGIXML_PKGCONFIG_FILE "pugixml")
if(TARGET ${PUGIXML_INSTALL_TARGET})
    # pugixml is installed by OpenXLSX
    set(PUGIXML_PKGCONFIG_FILE "pugixml-OpenXLSX")
endif()
set(ZIP_LIBRARY_PKGCONFIG_FILE "${ZIP_LIBRARY}")
if(TARGET ${ZIP_LIBRARY_INSTALL_TARGET})
    set(ZIP_LIBRARY_PKGCONFIG_FILE "${ZIP_LIBRARY}-OpenXLSX")
endif()

file(WRITE ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${OPENXLSX_LIBRARY_NAME}.pc.in.in
    "prefix=@DEFERRED@CMAKE_INSTALL_PREFIX@DEFERRED@\n"
    "exec_prefix=\${prefix}\n"
    "includedir=\${prefix}/@CMAKE_INSTALL_INCLUDEDIR@/@OPENXLSX_LIBRARY_NAME@\n" # includedir=${prefix}/include
    "libdir=\${exec_prefix}/lib\n"
    "\n"
    "Name: @OPENXLSX_LIBRARY_NAME@\n"
    "Description: @PROJECT_DESCRIPTION@\n"
    "URL: @PROJECT_HOMEPAGE_URL@\n"
    "Version: @PROJECT_VERSION@\n"
    "Cflags: -I\"\${includedir}\"\n"   # Cflags: -I${includedir}
    "Requires.private: ${PUGIXML_PKGCONFIG_FILE} >= 1.15 ${ZIP_LIBRARY_PKGCONFIG_FILE}\n"
    "Libs: -L\${libdir} -Wl,-rpath,\${libdir} -lOpenXLSX\n"
    "Libs.private: \n")

if(NOT OPENXLSX_ENABLE_LIBZIP)
    file(WRITE ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/miniz.pc.in.in
        "prefix=@DEFERRED@CMAKE_INSTALL_PREFIX@DEFERRED@\n"
        "exec_prefix=\${prefix}\n"
        "includedir=\${prefix}/include\n"
        "libdir=\${exec_prefix}/lib/@OPENXLSX_LIBRARY_NAME@\n"
        "\n"
        "Name: miniz\n"
        "Description: zip implementation\n"
        "URL: https://github.com/richgel999/miniz\n"
        "Version: 3.0.2\n" # TBD how to determine actual version
        "Cflags: -I\${includedir}\n"
        "Libs: -L\${libdir} -lminiz\n")
endif()

# Let CMake fill in the variables in the package config template
set( DEFERRED "@" )
configure_file(
    ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${OPENXLSX_LIBRARY_NAME}.pc.in.in
    ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${OPENXLSX_LIBRARY_NAME}.pc.in
    @ONLY
)
if(NOT OPENXLSX_ENABLE_LIBZIP)
    configure_file(
        ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/miniz.pc.in.in
        ${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/miniz.pc.in
        @ONLY
    )
endif()

# Fix @DEFERRED@ variables that are not known until install time (CMAKE_INSTALL_PREFIX)
install(
    CODE "configure_file(\"${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${OPENXLSX_LIBRARY_NAME}.pc.in\" \"${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${OPENXLSX_LIBRARY_NAME}.pc\" @ONLY)"
)
set(ZIP_LIBRARY_PKGCONFIG_SHOULD_EXIST FALSE)
if(NOT OPENXLSX_ENABLE_LIBZIP)
    install(
        CODE "configure_file(\"${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/miniz.pc.in\" \"${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/miniz.pc\" @ONLY)"
    )
    set(ZIP_LIBRARY_PKGCONFIG_SHOULD_EXIST TRUE) # For a file generated at build time, EXISTS can not be trivially verified, so use this boolean as a crutch
endif()


# Finally: install the package config files
set(PKGCONFIG_SUBDIR "pkgconfig")

install(
    FILES "${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/${OPENXLSX_LIBRARY_NAME}.pc"
    DESTINATION "${CMAKE_INSTALL_LIBDIR}/${PKGCONFIG_SUBDIR}"
)


# Install the pugixml package config file
if(TARGET ${PUGIXML_INSTALL_TARGET})
    message( DEBUG "${PUGIXML_TARGET} is pulled in from an external source" )
    # get_target_property(PUGIXML_IS_IMPORTED "${PUGIXML_INSTALL_TARGET}" IMPORTED)
    message( DEBUG "pugixml_BINARY_DIR is ${pugixml_BINARY_DIR}" )
    set(LIBPUGIXML_PC "${pugixml_BINARY_DIR}/pugixml.pc")
    if(EXISTS "${LIBPUGIXML_PC}")
        install(CODE "message(NOTICE \"Installing package config file for ${PUGIXML_INSTALL_TARGET}...\")")
        set(PUGIXML_PKGCONFIG_PATCHED "${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/pugixml-patched.pc")
        pkg_config_add_library_subdir("${LIBPUGIXML_PC}" "${PUGIXML_PKGCONFIG_PATCHED}" OpenXLSX) # prefix "Libs:" config with add -L{exec_prefix}/lib/OpenXLSX
        install(
            FILES "${PUGIXML_PKGCONFIG_PATCHED}"
            DESTINATION "${CMAKE_INSTALL_LIBDIR}/${PKGCONFIG_SUBDIR}"
            RENAME "${PUGIXML_PKGCONFIG_FILE}.pc"
        )
    endif()
endif()

# Install the zip library package config file, if any
if(TARGET ${ZIP_LIBRARY_INSTALL_TARGET})
    message( DEBUG "${ZIP_LIBRARY} is pulled in from an external source" )
    if(OPENXLSX_ENABLE_LIBZIP)
        message( DEBUG "libzip_BINARY_DIR is ${libzip_BINARY_DIR}" )
        set(ZIP_LIBRARY_PKGCONFIG_SOURCE "${libzip_BINARY_DIR}/libzip.pc")
        set(ZIP_LIBRARY_PKGCONFIG_PATCHED "${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/libzip-patched.pc")
        pkg_config_add_library_subdir("${ZIP_LIBRARY_PKGCONFIG_SOURCE}" "${ZIP_LIBRARY_PKGCONFIG_PATCHED}" OpenXLSX) # prefix "Libs:" config with add -L{exec_prefix}/lib/OpenXLSX
    else()
        set(ZIP_LIBRARY_PKGCONFIG_PATCHED "${CMAKE_LIBRARY_OUTPUT_DIRECTORY}/miniz.pc")
    endif()

    if(ZIP_LIBRARY_PKGCONFIG_SHOULD_EXIST OR EXISTS "${ZIP_LIBRARY_PKGCONFIG_SOURCE}")
        install(CODE "message(NOTICE \"Installing package config file for ${ZIP_LIBRARY_INSTALL_TARGET}...\")")
        install(
            FILES "${ZIP_LIBRARY_PKGCONFIG_PATCHED}"
            DESTINATION "${CMAKE_INSTALL_LIBDIR}/${PKGCONFIG_SUBDIR}"
            RENAME "${ZIP_LIBRARY_PKGCONFIG_FILE}.pc"
        )
    endif()
endif()
