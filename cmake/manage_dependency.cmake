# ============================================================================
# Author
# ============================================================================
# The function manage_dependency is a friendly contribution by NLATP (openxlsx.genetics016@passinbox.com)
# The rest of the module is written by Lars Uffmann (coding23@uffmann.name) in an attempt to better understand CMake platform independency
#  (yes, this is still giving me headaches)

# ============================================================================
# Policy Settings - to be controlled by the main project file
# ============================================================================
# set(CMAKE_POLICY_DEFAULT_CMP0077 NEW) # Respect BUILD_SHARED_LIBS / suppress warning in nowide - TBD if this is unproblematic
# if(POLICY CMP0077)
#     cmake_policy(SET CMP0077 NEW)  # Respect BUILD_SHARED_LIBS
# endif()
#
# set(CMAKE_POLICY_DEFAULT_CMP0167 NEW)
# if(POLICY CMP0167)
#     cmake_policy(SET CMP0167 NEW) # find the upstream BoostConfig.cmake directly
# endif()

# ============================================================================
# Configuration Options, to be configured in caller CMakeLists.txt
# ============================================================================
# option(PREFER_STATIC "Prefer static linking over shared libraries" ON)
# option(USE_SYSTEM_LIBS "Use system-installed libraries when available" ON)
# option(FETCH_DEPS_AUTO "Automatically fetch missing dependencies" ON)
# option(FORCE_FETCH_ALL "Ignore system libraries and fetch all deps" OFF)

# ============================================================================
# Safety Checks - to be handled in main project file
# ============================================================================
# if(FORCE_FETCH_ALL)
#     set(USE_SYSTEM_LIBS OFF CACHE BOOL "" FORCE)
#     message(STATUS "FORCE_FETCH_ALL enabled: ignoring system libraries")
# endif()

# ============================================================================
# Dependency Checks - verify that required modules are present
# ============================================================================
message( NOTICE "debug manage_dependency.cmake: CMAKE_CURRENT_LIST_DIR is ${CMAKE_CURRENT_LIST_DIR}" )
message( NOTICE "debug manage_dependency.cmake: CMAKE_SOURCE_DIR is ${CMAKE_SOURCE_DIR}" )
include(${CMAKE_CURRENT_LIST_DIR}/recursive_library_type_test.cmake) # not actually invoked from here, but from project created by write_find_test_config

# ============================================================================
# Enhanced Helper Functions
# ============================================================================

# Function:   FetchContent_Declare_custom
# Purpose:    Provide an alternative to FetchContent_Declare with the option to EXCLUDE_FROM_ALL
# Options:
#   EXCLUDE_FROM_ALL  if supported, pass EXCLUDE_FROM_ALL on to FetchContent_Declare
# Parameters:
#   LIB_NAME          name of the library to make available, as to be passed to FetchContent_Declare
#   URL               if provided together with a hash, declare content for fetching from URL
#   URL_HASH          must be provided with URL or URL won't be used
#   GIT_REPOSITORY    if URL or URL_HASH are missing: declare content for fetching from this repository
#   GIT_TAG           must be provided with GIT_REPOSITORY
# Author:     Lars Uffmann (coding23@uffmann.name)
function(FetchContent_Declare_custom)
    set(options "EXCLUDE_FROM_ALL")
    set(oneValueArgs
        LIB_NAME
        URL
        URL_HASH
        GIT_REPOSITORY
        GIT_TAG
    )
    set(multiValueArgs EXTRA_ARGS)
    cmake_parse_arguments(ARG "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})

    # empty initializers for all parameters to FetchContent_Declare
    set( URL_PARAM )
    set( URL_HASH_PARAM )
    set( URL_EXTRA_PARAMS )
    set( GIT_REPO_PARAM )
    set( GIT_TAG_PARAM )
    set( GIT_EXTRA_PARAMS )
    set( SHARED_PARAMS )

    # initialize source parameters
    if( ( "${ARG_URL}" STREQUAL "" ) OR ( "${ARG_URL_HASH}" STREQUAL "" ) )
        message( NOTICE "Fetching ${ARG_LIB_NAME} from repository ${ARG_GIT_REPOSITORY} with GIT_TAG ${ARG_GIT_TAG}" )

        # set( GIT_REPO_PARAM "GIT_REPOSITORY;${ARG_GIT_REPOSITORY}" )
        # NOTE: separating param name and value with a semicolon (above) is equivalent to list( APPEND [..] ) as seen below
        list( APPEND GIT_REPO_PARAM      GIT_REPOSITORY             ${ARG_GIT_REPOSITORY} )
        list( APPEND GIT_TAG_PARAM       GIT_TAG                    ${ARG_GIT_TAG} )
        list( APPEND GIT_EXTRA_PARAMS    GIT_SHALLOW                TRUE )
    else()
        message( NOTICE "Fetching ${ARG_LIB_NAME} via URL ${ARG_URL}" )

        list( APPEND URL_PARAM           URL                        ${ARG_URL} )
        list( APPEND URL_HASH_PARAM      URL_HASH                   ${ARG_URL_HASH} )
        list( APPEND URL_EXTRA_PARAMS    DOWNLOAD_EXTRACT_TIMESTAMP OFF ) # set extracted file timestamps to current time to ensure extraction triggers compilation
    endif()
    list( APPEND     SHARED_PARAMS       OVERRIDE_FIND_PACKAGE )    # Important: override system package
    if( CMAKE_VERSION VERSION_GREATER_EQUAL "3.30" )
        if( ARG_EXCLUDE_FROM_ALL )
            message( NOTICE "FetchContent_Declare in CMake >= 3.30 will be invoked with EXCLUDE_FROM_ALL" )
            list( APPEND SHARED_PARAMS   EXCLUDE_FROM_ALL )
        endif()
    endif()

    FetchContent_Declare(
        ${ARG_LIB_NAME}
        ${URL_PARAM}
        ${URL_HASH_PARAM}
        ${URL_EXTRA_PARAMS}
        ${GIT_REPO_PARAM}
        ${GIT_TAG_PARAM}
        ${GIT_EXTRA_PARAMS}
        ${SHARED_PARAMS}
    )
endfunction()
# End of Function:   FetchContent_Declare_custom


# Function:   FetchContent_MakeAvailable_custom
# Purpose:    Provide an alternative to FetchContent_MakeAvailable with the option to EXCLUDE_FROM_ALL in case FetchContent_Declare
#              does not support it
# Options:
#   EXCLUDE_FROM_ALL  FetchContent_MakeAvailable (CMake >= 3.30) or add_subdirectory will be executed with EXCLUDE_FROM_ALL
# Parameters:
#   LIB_NAME          name of the library to make available, as previously prepared with FetchContent_Declare
# Author:     Lars Uffmann (coding23@uffmann.name)
function(FetchContent_MakeAvailable_custom)
    set(options "EXCLUDE_FROM_ALL")
    set(oneValueArgs
        LIB_NAME
    )
    set(multiValueArgs EXTRA_ARGS)
    cmake_parse_arguments(ARG "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})

    # empty initializers for all parameters to FetchContent_MakeAvailable or add_subdirectory
    set( EXCLUDE_FROM_ALL_PARAM ) # ignored for CMake >= 3.30

    if( CMAKE_VERSION VERSION_GREATER_EQUAL "3.30" )
        # if CMake version >= 3.30: FetchContent_Declare was invoked with EXCLUDE_FROM_ALL, now use FetchContent_MakeAvailable nominally
        FetchContent_MakeAvailable(
            ${ARG_LIB_NAME}
        )
    else()
        # older versions of CMake: fall back to a combination of FetchContent_Populate and add_subdirectory

        # Logic from comment at  https://stackoverflow.com/questions/65527126/disable-install-for-fetchcontent
        FetchContent_GetProperties(${ARG_LIB_NAME})
        if(NOT ${ARG_LIB_NAME}_POPULATED)
            FetchContent_Populate(${ARG_LIB_NAME})
            string( TOLOWER ${ARG_LIB_NAME} LIB_NAME_LOWER ) # FetchContent_Populate populates variables with ARG_LIB_NAME in all lowercase
            if( ARG_EXCLUDE_FROM_ALL )
                message( NOTICE "-- FetchContent_MakeAvailable_custom (${ARG_LIB_NAME}): add_subdirectory will be used with EXCLUDE_FROM_ALL" )
                list( APPEND EXCLUDE_FROM_ALL_PARAM EXCLUDE_FROM_ALL )  # Ensure that the dependency is not installed if EXCLUDE_FROM_ALL is active
            endif()
            add_subdirectory(
                ${${LIB_NAME_LOWER}_SOURCE_DIR}
                ${${LIB_NAME_LOWER}_BINARY_DIR}
                ${EXCLUDE_FROM_ALL_PARAM}
            )
        endif()
    endif()
endfunction()
# End of Function:   FetchContent_MakeAvailable_custom


# Function: write_find_test_config
# Purpose: in ARG_TEMP_SRC_DIR, create a CMakeLists.txt with the sole purpose of attempting to find_package for a required library type
# Parameters:
#   ARG_TEMP_SRC_DIR    the (temporary) directory in which to create CMakeLists.txt
# Returns: N/A
# Author:     Lars Uffmann (coding23@uffmann.name)
function(write_find_test_config ARG_TEMP_SRC_DIR)
    file(WRITE "${ARG_TEMP_SRC_DIR}/CMakeLists.txt"
"cmake_minimum_required(VERSION 3.15)
project(TestPackage VERSION 0.1 LANGUAGES CXX)
# Script: CMakeLists.txt
# Purpose: demonstrate finding a required library type (static vs. shared) with Boost nowide as example
# Example Usage: for Boost_USE_STATIC_LIBS=OFF test, call like so (for ON setting, TEST_REQUIRED_TYPE=STATIC_LIBRARY and TEST_FIND_LIBRARY_SUFFIXES=\".a .lib\" would be needed:
#
# 1) for nowide_standalone:
# cmake --no-warn-unused-cli -L -S . \
#     -D TEST_PACKAGE_NAME=nowide \
#     -D TEST_VERSION= \
#     -D TEST_EXTRA_ARGS=QUIET \
#     -D TEST_COMPONENTS_PARAM= \
#     -D TEST_TARGET_NAME_SYSTEM=nowide::nowide \
#     -D TEST_REQUIRED_TYPE=SHARED_LIBRARY \
#     -D Boost_USE_STATIC_LIBS=OFF \
#     -D TEST_FIND_LIBRARY_SUFFIXES=\".so .dylib .dll .dll.a\" \
#     -D CMAKE_PREFIX_PATH= \
#     -D CMAKE_FIND_DEBUG_MODE=OFF
#
# 2) for Boost::nowide:
# cmake --no-warn-unused-cli -L -S . \
#     -D TEST_PACKAGE_NAME=Boost \
#     -D TEST_VERSION= \
#     -D TEST_EXTRA_ARGS=QUIET \
#     -D TEST_COMPONENTS_PARAM=\"COMPONENTS;nowide\" \
#     -D TEST_TARGET_NAME_SYSTEM=Boost::nowide \
#     -D TEST_REQUIRED_TYPE=SHARED_LIBRARY \
#     -D Boost_USE_STATIC_LIBS=OFF \
#     -D TEST_FIND_LIBRARY_SUFFIXES=\".so .dylib .dll .dll.a\" \
#     -D CMAKE_PREFIX_PATH= \
#     -D CMAKE_FIND_DEBUG_MODE=OFF

include(\"${CMAKE_CURRENT_LIST_DIR}/recursive_library_type_test.cmake\")

if(NOT DEFINED TEST_PACKAGE_NAME)
    message( FATAL_ERROR \"TEST_PACKAGE_NAME undefined\" )
endif()
if(NOT DEFINED TEST_VERSION)
    message( FATAL_ERROR \"TEST_VERSION undefined\" )
endif()
if(NOT DEFINED TEST_EXTRA_ARGS)
    message( FATAL_ERROR \"TEST_EXTRA_ARGS undefined\" )
endif()
if(NOT DEFINED TEST_COMPONENTS_PARAM)
    message( FATAL_ERROR \"TEST_COMPONENTS_PARAM undefined\" )
endif()
if(NOT DEFINED TEST_TARGET_NAME_SYSTEM)
    message( FATAL_ERROR \"TEST_TARGET_NAME_SYSTEM undefined\" )
endif()
if(NOT DEFINED TEST_REQUIRED_TYPE)
    message( FATAL_ERROR \"TEST_REQUIRED_TYPE undefined\" )
endif()
if(NOT DEFINED TEST_FIND_LIBRARY_SUFFIXES)
    message( FATAL_ERROR \"TEST_FIND_LIBRARY_SUFFIXES undefined\" )
endif()

# configure FIND_LIBRARY_SUFFIXES - setting will be forgotten when this script completes
set( CMAKE_FIND_LIBRARY_SUFFIXES \${TEST_FIND_LIBRARY_SUFFIXES} )

message( STATUS \">>>>>>>>>>> TestPackage: TEST_FIND_LIBRARY_SUFFIXES is \${TEST_FIND_LIBRARY_SUFFIXES}\" )

message( STATUS \">>>>>>>>>>> TestPackage: attempting to find_package \${TEST_PACKAGE_NAME} version \${TEST_VERSION} with \${TEST_EXTRA_ARGS} \${TEST_COMPONENTS_PARAM}\" )
message( STATUS \">>>>>>>>>>> TestPackage: expecting to find system target \${TEST_TARGET_NAME_SYSTEM}, required library type is \\\"\${TEST_REQUIRED_TYPE}\\\", CMAKE_FIND_LIBRARY_SUFFIXES is \${CMAKE_FIND_LIBRARY_SUFFIXES}\" )

# allow parent to force debug
set(CMAKE_FIND_DEBUG_MODE ${CMAKE_FIND_DEBUG_MODE} CACHE BOOL \"\" FORCE)

# perform the find
find_package(
    \${TEST_PACKAGE_NAME} \${TEST_VERSION}
    \${TEST_EXTRA_ARGS}          # QUIET can be passed here
    \${TEST_COMPONENTS_PARAM}
)
message( STATUS \">>>>>>>>>>> TestPackage: \${TEST_PACKAGE_NAME}_FOUND: \${\${TEST_PACKAGE_NAME}_FOUND}\" )

# evaluate result
set(\${TEST_PACKAGE_NAME}_SUITABLE_FOUND FALSE CACHE STRING \"\" FORCE)
if(\${TEST_PACKAGE_NAME}_FOUND)
    message( STATUS \">>>>>>>>>>> TestPackage: Performing recursive_library_type_test on \${TEST_TARGET_NAME_SYSTEM}\" )
    set(LIBRARY_TYPE \"INTERFACE_LIBRARY\")   # default: can't resolve library type
    recursive_library_type_test(\${TEST_TARGET_NAME_SYSTEM} \"\")
    if(\${TEST_TARGET_NAME_SYSTEM}_HAS_STATIC)
        if(NOT \${TEST_TARGET_NAME_SYSTEM}_HAS_SHARED)
            set(LIBRARY_TYPE \"STATIC_LIBRARY\")      # library is only static
        endif()
    elseif(\${TEST_TARGET_NAME_SYSTEM}_HAS_SHARED)
        set(LIBRARY_TYPE \"SHARED_LIBRARY\")          # library is only shared
    endif()
    message( STATUS \">>>>>>>>>>> TestPackage: recursive_library_type_test on \${TEST_TARGET_NAME_SYSTEM} determined type \${LIBRARY_TYPE}\" )

    # Final check: LIBRARY_TYPE must match expected type, if any
    if(\"\${TEST_REQUIRED_TYPE}\" STREQUAL \"\" OR \"\${LIBRARY_TYPE}\" STREQUAL \"\${TEST_REQUIRED_TYPE}\")
        message( STATUS \">>>>>>>>>>> TestPackage: \${TEST_TARGET_NAME_SYSTEM} type matches required type \${TEST_REQUIRED_TYPE}\" )
        if(TARGET \${TEST_TARGET_NAME_SYSTEM})
            set(\${TEST_PACKAGE_NAME}_SUITABLE_FOUND TRUE CACHE STRING \"\" FORCE)
        else()
            message( WARNING \"TestPackage: expected TARGET \${TEST_TARGET_NAME_SYSTEM} is not available!\" )
        endif()
    else()
        message( WARNING \"TestPackage: \${TEST_TARGET_NAME_SYSTEM} type does not match required type \${TEST_REQUIRED_TYPE}\" )
    endif()
endif()
"
)
endfunction()
# End of Function: write_find_test_config


# Function: write_parent_cache_init
# Author: unknown, stolen without attribution by a slop generator :(, then heavily modified
# Usage:
#   write_parent_cache_init(
#       OUTPUT_FILE "${DEST_DIR}/${OUTFILE}.cmake"
#       INCLUDE_FILTER_REGEX "^(INCLUDE_THESE_).*"
#       EXCLUDE_FILTER_REGEX "^(EXCLUDE_THESE).*"
#   )
# Parameters:
#   1) output file path
#   2) optional regex to select which cache variable names to include (default: include everything)
#   3) optional regex to select which cache variable names to exclude (default: exclude nothing)
function( write_parent_cache_init )
    set(options)
    set(oneValueArgs
        OUTPUT_FILE
        INCLUDE_FILTER_REGEX
        EXCLUDE_FILTER_REGEX
    )
    cmake_parse_arguments(_args "${options}" "${oneValueArgs}" "" ${ARGN})

    if(NOT _args_INCLUDE_FILTER_REGEX)
        set( _args_INCLUDE_FILTER_REGEX ".*" )  # default: include all
    endif()

    if(NOT _args_EXCLUDE_FILTER_REGEX)
        set( _args_EXCLUDE_FILTER_REGEX "" )    # default: exclude nothing
    endif()

    # Create cache init file
    file(WRITE "${_args_OUTPUT_FILE}" "# Auto-generated cache init file\n")

    # Get all cache variable names
    get_cmake_property(_cache_vars CACHE_VARIABLES)

# message( NOTICE "regex exclude filter is ${_args_EXCLUDE_FILTER_REGEX}" )
    foreach(_var ${_cache_vars})
# if( _var MATCHES "${_args_EXCLUDE_FILTER_REGEX}" )
#     message( NOTICE "_var ${_var} MATCHES regex exclude filter!" )
# else()
#     message( NOTICE "_var ${_var} does NOT match regex exclude filter!" )
# endif()

        # Skip variables that don't match the INCLUDE filter
        if( NOT _var MATCHES "${_args_INCLUDE_FILTER_REGEX}" )
    continue()
        endif()

        # Skip variables that do match the EXCLUDE filter
        if(     _var MATCHES "${_args_EXCLUDE_FILTER_REGEX}" )
    continue()
        endif()

        # Get type and value
        get_property( _type CACHE ${_var} PROPERTY TYPE )
        get_property( _value CACHE ${_var} PROPERTY VALUE )
# message( STATUS "var is ${_var}, type is ${_type}, value is ${_value}" )

        # Skip internal or undocumented entries
        if( _type STREQUAL "INTERNAL" )
    continue()
        endif()

        # Default type if missing
        if( NOT _type )
            set( _type "STRING" )
        endif()

        # Escape double quotes and backslashes in the value
        if( DEFINED _value )
            string( REPLACE "\\" "\\\\" _value_esc "${_value}" )
            string( REPLACE "\"" "\\\"" _value_esc "${_value_esc}" )
        else()
            set( _value_esc "" )
        endif()

        # Write a set(... CACHE ...) line
        file( APPEND "${_args_OUTPUT_FILE}"
            "set(${_var} \"${_value_esc}\" CACHE ${_type} \"copied from parent\")\n")
    endforeach()
endfunction()
# End of Function: write_parent_cache_init


# Function: run_test_package
# Purpose: determine if ARG_PACKAGE_NAME can be found on the system as the REQUIRED_LIBRARY_TYPE if any
# Parameters:
#     pass-through to find_package:
#   ARG_PACKAGE_NAME
#   ARG_VERSION
#   ARG_EXTRA_ARGS
#   COMPONENTS_PARAM
#     evaluation of success:
#   ARG_TARGET_NAME_SYSTEM  used to get_target_property TYPE after successful find_package
#   REQUIRED_LIBRARY_TYPE   evaluate against TYPE of found package
#   FIND_PACKAGE_VARIABLES  a NAME VALUE NAME VALUE list of variables to configure for the test_package call
#   FIND_LIBRARY_SUFFIXES   configure CMAKE_FIND_LIBRARY_SUFFIXES to this before call to find_package - has no effect when cmake package config files are found
# Returns:
#   ${PACKAGE_NAME}_SUITABLE_FOUND        (PARENT_SCOPE): TRUE if dependency is available as the REQUIRED_LIBRARY_TYPE, FALSE otherwise
# Author:     Lars Uffmann (coding23@uffmann.name)
function(run_test_package ARG_PACKAGE_NAME ARG_VERSION ARG_EXTRA_ARGS COMPONENTS_PARAM ARG_TARGET_NAME_SYSTEM REQUIRED_LIBRARY_TYPE FIND_PACKAGE_VARIABLES FIND_LIBRARY_SUFFIXES)
    # configure a build directory
    string(REPLACE ";" "_" pkg_safe "${ARG_PACKAGE_NAME}")

    # escape variables that may contain a list
    string(REPLACE ";" "\\;" PREFIX_ESCAPED "${CMAKE_PREFIX_PATH}")
    string(REPLACE ";" "\\;" ARG_EXTRA_ARGS_ESCAPED "${ARG_EXTRA_ARGS}")
    string(REPLACE ";" "\\;" COMPONENTS_PARAM_ESCAPED "${COMPONENTS_PARAM}")

    set( PARAM_FIND_PACKAGE_VARIABLES "" )  # for CMake execute_process parameters
    set(  ECHO_FIND_PACKAGE_VARIABLES "" )  # for command echo / bash executable
    unset( FPC_NAME )
    unset( FPC_VALUE )
    foreach( val IN LISTS ARG_FIND_PACKAGE_VARIABLES )
        if(NOT FPC_NAME)
            set(FPC_NAME ${val})
        else()
            set(FPC_VALUE ${val})
            message( STATUS ">>>>>>>>>>> run_test_package: FIND_PACKAGE_VARIABLES ${FPC_NAME}=${FPC_VALUE}" )
            if(NOT "${PARAM_FIND_PACKAGE_VARIABLES}" STREQUAL "")                                               # if this is not the first parameter to be configured
                set( PARAM_FIND_PACKAGE_VARIABLES "${PARAM_FIND_PACKAGE_VARIABLES};" )                              # append a semicolon as separator       -> CMake list
                set(  ECHO_FIND_PACKAGE_VARIABLES "${ECHO_FIND_PACKAGE_VARIABLES} " )                               # append a space for the command echo   -> bash compatible
            endif()
            set( PARAM_FIND_PACKAGE_VARIABLES "${PARAM_FIND_PACKAGE_VARIABLES}-D ${FPC_NAME}=${FPC_VALUE}" )    # append a parameter for the script execution environment   -> CMake list
            set(  ECHO_FIND_PACKAGE_VARIABLES  "${ECHO_FIND_PACKAGE_VARIABLES}-D ${FPC_NAME}=${FPC_VALUE}" )    # append a parameter for the command echo                   -> bash compatible
            unset( FPC_NAME )
            unset( FPC_VALUE )
        endif()
    endforeach()

    # create a temp file path
    set(_test_package_init "${TEMP_BUILD_DIR}/test_package_init.cmake")

    # write all variables that hopefully won't make the subscript interfere with the OpenXLSX build
    write_parent_cache_init(
        OUTPUT_FILE "${_test_package_init}"
        INCLUDE_FILTER_REGEX "^(CMAKE_|GIT_EXECUTABLE).*"
        EXCLUDE_FILTER_REGEX "^(CMAKE_PROJECT_|CMAKE_FIND_PACKAGE_).*" )

    # Echo the command to be executed for TestPackage to the log so that the user can manually invoke "cmake-find-test-<package>-src/CMakeLists.txt" and experiment with the results
    message( STATUS "run_test_package: testing package ${ARG_PACKAGE_NAME} with command:
                ${CMAKE_COMMAND} --no-warn-unused-cli -L -S \"${TEMP_SRC_DIR}\" \\
                -B \"${TEMP_BUILD_DIR}\" \\
                -C \"${_test_package_init}\" \\
                -D TEST_PACKAGE_NAME=${ARG_PACKAGE_NAME} \\
                -D TEST_VERSION=${ARG_VERSION} \\
                -D TEST_EXTRA_ARGS=${ARG_EXTRA_ARGS_ESCAPED} \\
                -D TEST_COMPONENTS_PARAM=${COMPONENTS_PARAM_ESCAPED} \\
                -D TEST_TARGET_NAME_SYSTEM=${ARG_TARGET_NAME_SYSTEM} \\
                -D TEST_REQUIRED_TYPE=${REQUIRED_LIBRARY_TYPE} \\
                -D TEST_FIND_LIBRARY_SUFFIXES=\"${FIND_LIBRARY_SUFFIXES}\" \\
                -D CMAKE_PREFIX_PATH=${PREFIX_ESCAPED} \\
                -D CMAKE_FIND_DEBUG_MODE=OFF \\
                ${ECHO_FIND_PACKAGE_VARIABLES} \\
    ")

    execute_process(
        COMMAND ${CMAKE_COMMAND} --no-warn-unused-cli -L -S "${TEMP_SRC_DIR}"
                -B "${TEMP_BUILD_DIR}"
                -C "${_test_package_init}"
                -D TEST_PACKAGE_NAME=${ARG_PACKAGE_NAME}
                -D TEST_VERSION=${ARG_VERSION}
                -D TEST_EXTRA_ARGS=${ARG_EXTRA_ARGS_ESCAPED}
                -D TEST_COMPONENTS_PARAM=${COMPONENTS_PARAM_ESCAPED}
                -D TEST_TARGET_NAME_SYSTEM=${ARG_TARGET_NAME_SYSTEM}
                -D TEST_REQUIRED_TYPE=${REQUIRED_LIBRARY_TYPE}
                -D TEST_FIND_LIBRARY_SUFFIXES="${FIND_LIBRARY_SUFFIXES}"
                -D CMAKE_PREFIX_PATH=${PREFIX_ESCAPED}
                -D CMAKE_FIND_DEBUG_MODE=OFF
                ${PARAM_FIND_PACKAGE_VARIABLES}
        RESULT_VARIABLE test_package_success
        OUTPUT_VARIABLE test_package_output
        ERROR_VARIABLE test_package_err
    )

    message(STATUS ">>>>>>>>>>> run_test_package output: ${test_package_output}")
    message(STATUS ">>>>>>>>>>> run_test_package errors: ${test_package_err}")
    message(STATUS ">>>>>>>>>>> run_test_package: result for ${ARG_PACKAGE_NAME} (test_package_success=${test_package_success})")

    # define function return value: ${ARG_PACKAGE_NAME}_SUITABLE_FOUND
    set(${ARG_PACKAGE_NAME}_SUITABLE_FOUND FALSE PARENT_SCOPE)
    if(test_package_output MATCHES "${ARG_PACKAGE_NAME}_SUITABLE_FOUND:STRING=([^\r\n]+)")     # test if test_package_output contains lines like: ${ARG_PACKAGE_NAME}_SUITABLE_FOUND:STRING=<bool>
        set(test_result "${CMAKE_MATCH_1}")
        # string(STRIP "${CMAKE_MATCH_1}" test_result)
        if("${test_result}" STREQUAL "TRUE")
            set(${ARG_PACKAGE_NAME}_SUITABLE_FOUND TRUE PARENT_SCOPE)
        endif()
    endif()
endfunction()
# End of Function: run_test_package


# Function: manage_dependency
# Purpose:  Safely finds or fetches a dependency
# Options:  -
# Parameters: (see below)
# Returns:
#   ${LIB_NAME}_FETCHED         (PARENT_SCOPE): TRUE if dependency was downloaded from a repository, FALSE if dependency was provided by find_package
#   ${LIB_NAME}_TARGET_STATIC   (PARENT_SCOPE): static library target file, if any
#   ${LIB_NAME}_PROVIDED_TARGET (PARENT_SCOPE): (potentially aliased) target provided by the dependency
#   ${LIB_NAME}_INSTALL_TARGET  (PARENT_SCOPE): (alias resolved) install target to be used for the dependency, if any
# Author:   This function is a friendly contribution by NLATP (openxlsx.genetics016@passinbox.com)
function(manage_dependency)
    set(options "")
    set(oneValueArgs
        LIB_NAME        # an arbitrary library name, for which ${LIB_NAME}_FETCHED will be set to TRUE for the caller in case of success
        PACKAGE_NAME    # the system library package name
        VERSION         # local required version for find_package with PACKAGE_NAME
        #
        GITHUB_REPO     # use for github repositories only
        GIT_REPOSITORY  # use to provide a full repository URL
        GIT_TAG         # use this to supply the github version tag (can be preceeded by a 'v' whereas VERSION is typically(?) numerical only
        HEADER_FILE     # TBD: what is the point of HEADER_FILE?
        URL             # URL & URL_HASH can be used together instead of GITHUB_REPO & GIT_TAG
        URL_HASH        #
        #
        TARGET_NAME  # expected target (can be an alias) that should be provided by the dependency
        TARGET_NAME_SYSTEM  # this is the target that shold be provided when the library is provided by the system
        # NOTE: depending on how the dependency is made available, manage_dependency will set the variable ${LIB_NAME}_PROVIDED_TARGET to
        #        either TARGET_NAME or TARGET_NAME_SYSTEM, so that ${LIB_NAME}_PROVIDED_TARGET can be used as DEPENDENCY_TARGET in target_link_interface
    )
    set(multiValueArgs
        COMPONENTS      # system library components name(s)
        TYPICAL_NAMES   # typical names for the dependency to be used with find_library in determining whether a static library is installed
        EXTRA_ARGS
        FIND_PACKAGE_VARIABLES       # dependency specific constants to be set for find_package, must be of pattern NAME VALUE NAME VALUE ...
    )

    cmake_parse_arguments(ARG "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})

    unset( FPC_NAME )
    unset( FPC_VALUE )
    foreach( val IN LISTS ARG_FIND_PACKAGE_VARIABLES )
        if(NOT FPC_NAME)
            set(FPC_NAME ${val})
        else()
            set(FPC_VALUE ${val})
            message( STATUS ">>>>>>>>>>> manage_dependency: configuring FIND_PACKAGE_VARIABLES ${FPC_NAME}=${FPC_VALUE}" )
            set(${FPC_NAME}_OLD_VALUE ${${FPC_NAME}})   # save old variable value
            set(${FPC_NAME} ${FPC_VALUE})               # store desired setting
            unset( FPC_NAME )
            unset( FPC_VALUE )
        endif()
    endforeach()
    if(FPC_NAME)
        message( FATAL_ERROR "manage_dependency: argument list for FIND_PACKAGE_VARIABLES requires a value for key ${FPC_NAME}" )
    endif()

    if(USE_SYSTEM_LIBS AND FORCE_FETCH_ALL)
        message( FATAL_ERROR "manage_dependency: USE_SYSTEM_LIBS and FORCE_FETCH_ALL are mutually exclusive - choose one!" )
    endif()

    # Initialize return values to defaults
    set(${ARG_LIB_NAME}_FETCHED FALSE PARENT_SCOPE)     # assume dependency will be found already installed
    set(${ARG_LIB_NAME}_TARGET_STATIC "" PARENT_SCOPE)  # assume no static library found on system
    set(${ARG_LIB_NAME}_PROVIDED_TARGET                 # this shall should be set to the name provided by a fetched dependency
        ${ARG_TARGET_NAME} PARENT_SCOPE)                #   (to be overridden by TARGET_NAME_SYSTEM if a system library is used)
    set(${ARG_LIB_NAME}_INSTALL_TARGET "" PARENT_SCOPE) # no install target for an already installed dependency

    set(should_fetch FALSE) # 2026-01-25: default initialization

    # DEBUG: enable to verify configuration
    # message( NOTICE     "   ++ Prefer static linking over shared libraries: ${PREFER_STATIC}" )
    # message( NOTICE     "   ++ Use system-installed libraries when available: ${USE_SYSTEM_LIBS}" )
    # message( NOTICE     "   ++ Automatically fetch missing dependencies: ${FETCH_DEPS_AUTO}" )
    # message( NOTICE     "   ++ Ignore system libraries and fetch all deps: ${FORCE_FETCH_ALL}" )

    # Determine search strategy
    if(NOT USE_SYSTEM_LIBS)
        set(should_fetch TRUE)
    else()
        # Try system first

        # 2026-01-25: test whether components are specified & call find_package accordingly
        set( COMPONENTS_PARAM )
        if( "${ARG_COMPONENTS}" STREQUAL "" )
            message( NOTICE ">>>>>>>>>>> manage_dependency: attempting to find_package ${ARG_PACKAGE_NAME} ${ARG_VERSION} with ${ARG_EXTRA_ARGS} QUIET" )
        else()
            list( APPEND COMPONENTS_PARAM COMPONENTS ${ARG_COMPONENTS} )
            message( NOTICE ">>>>>>>>>>> manage_dependency: attempting to find_package ${ARG_PACKAGE_NAME} ${ARG_VERSION} with ${ARG_EXTRA_ARGS} ${COMPONENTS_PARAM} QUIET" )
        endif()

        set(DO_TEST_PACKAGE FALSE)      # initialize: setting this to TRUE will trigger run_test_package
        set(TRY_FIND_PACKAGE TRUE)      # initialize: setting this to TRUE will trigger find_package

message( STATUS ">>>>>>>>>>> manage_dependency: PREFER_STATIC is ${PREFER_STATIC}, BUILD_SHARED_LIBS is ${BUILD_SHARED_LIBS}" )
        if(PREFER_STATIC OR BUILD_SHARED_LIBS)
            set(DO_TEST_PACKAGE TRUE)      # trigger run_test_package
            set(SAVED_CMAKE_FIND_LIBRARY_SUFFIXES ${CMAKE_FIND_LIBRARY_SUFFIXES})   # save CMAKE_FIND_LIBRARY_SUFFIXES
        endif()
message( STATUS ">>>>>>>>>>> manage_dependency: package ${ARG_PACKAGE_NAME} DO_TEST_PACKAGE is ${DO_TEST_PACKAGE}" )

        # Only attempt to find_package if desired library has not been determined unavailable before
        if(DO_TEST_PACKAGE)

            if(PREFER_STATIC)
                set(REQUIRED_LIBRARY_TYPE "STATIC_LIBRARY")
                set(CMAKE_FIND_LIBRARY_SUFFIXES ".a .lib")
            elseif(BUILD_SHARED_LIBS)
                set(REQUIRED_LIBRARY_TYPE "SHARED_LIBRARY")
                set(CMAKE_FIND_LIBRARY_SUFFIXES ".so .dylib .dll .dll.a")
            else()  # shouldn't execute
                set(REQUIRED_LIBRARY_TYPE "")
            endif()

            set(TEMP_SRC_DIR "${CMAKE_BINARY_DIR}/cmake-find-test-${ARG_PACKAGE_NAME}-src")
            file(MAKE_DIRECTORY "${TEMP_SRC_DIR}")
            write_find_test_config(${TEMP_SRC_DIR})

            set(TEMP_BUILD_DIR "${CMAKE_BINARY_DIR}/cmake-find-test-${ARG_PACKAGE_NAME}-build")
            set(TRY_EXTRA_ARGS ${ARG_EXTRA_ARGS} QUIET)

set(RUN_TEST_PACKAGE_ECHO "\"${ARG_PACKAGE_NAME}\" \"${ARG_VERSION}\"")
set(RUN_TEST_PACKAGE_ECHO "${RUN_TEST_PACKAGE_ECHO} \"${TRY_EXTRA_ARGS}\"")
set(RUN_TEST_PACKAGE_ECHO "${RUN_TEST_PACKAGE_ECHO} \"${COMPONENTS_PARAM}\"")
set(RUN_TEST_PACKAGE_ECHO "${RUN_TEST_PACKAGE_ECHO} \"${ARG_TARGET_NAME_SYSTEM}\"")
set(RUN_TEST_PACKAGE_ECHO "${RUN_TEST_PACKAGE_ECHO} \"${REQUIRED_LIBRARY_TYPE}\"")
set(RUN_TEST_PACKAGE_ECHO "${RUN_TEST_PACKAGE_ECHO} \"${ARG_FIND_PACKAGE_VARIABLES}\"")
set(RUN_TEST_PACKAGE_ECHO "${RUN_TEST_PACKAGE_ECHO} \"${CMAKE_FIND_LIBRARY_SUFFIXES}\"")
message( STATUS ">>>>>>>>>>> manage_dependency: run_test_package( ${RUN_TEST_PACKAGE_ECHO} )" )
            run_test_package(
                "${ARG_PACKAGE_NAME}" "${ARG_VERSION}"
                "${TRY_EXTRA_ARGS}"
                "${COMPONENTS_PARAM}"
                "${ARG_TARGET_NAME_SYSTEM}"
                "${REQUIRED_LIBRARY_TYPE}"
                "${ARG_FIND_PACKAGE_VARIABLES}"
                "${CMAKE_FIND_LIBRARY_SUFFIXES}"
            )
message( STATUS ">>>>>>>>>>> manage_dependency: ${ARG_PACKAGE_NAME}_SUITABLE_FOUND is ${${ARG_PACKAGE_NAME}_SUITABLE_FOUND}" )
            if(NOT ${ARG_PACKAGE_NAME}_SUITABLE_FOUND)
                message( WARNING "manage_dependency: could not find a suitable installed version of ${ARG_PACKAGE_NAME}" )
                set(TRY_FIND_PACKAGE FALSE)  # inhibit find_package
            else()
            endif()

if(FALSE) # temporarily disable temporary file deletion
            # Delete temporary folders used to analyze find_package results
            file(REMOVE_RECURSE "${TEMP_SRC_DIR}" "${TEMP_BUILD_DIR}")
endif()
        endif()
message( STATUS ">>>>>>>>>>> manage_dependency: package ${ARG_PACKAGE_NAME} TRY_FIND_PACKAGE is ${TRY_FIND_PACKAGE}" )

        if(TRY_FIND_PACKAGE)
            set(DO_EXTRA_ARGS ${ARG_EXTRA_ARGS} QUIET)
message( STATUS ">>>>>>>>>>> manage_dependency: find_package( ${ARG_PACKAGE_NAME} ${ARG_VERSION} ${DO_EXTRA_ARGS} ${COMPONENTS_PARAM} )" )
            find_package(${ARG_PACKAGE_NAME} ${ARG_VERSION}
                ${DO_EXTRA_ARGS}
                ${COMPONENTS_PARAM}
                # QUIET
            )
message( STATUS ">>>>>>>>>>> manage_dependency: ${ARG_PACKAGE_NAME}_FOUND is ${${ARG_PACKAGE_NAME}_FOUND}" )

            # TODO: find a better strategy to deal with a missing component (prevent polluting build configuration)
            if(${ARG_PACKAGE_NAME}_FOUND)
                foreach(component IN LISTS ARG_COMPONENTS)
                    string(TOUPPER "${component}" component_upper)
                    string(REGEX REPLACE "[^A-Z0-9]" "_" component_upper "${component_upper}")
                    message( NOTICE ">>>>>>>>>>> manage_dependency:  COMPONENT is ${component}, upper case is ${component_upper}" )
                    set(foundVar "${ARG_PACKAGE_NAME}_${component_upper}_FOUND")
                    if(NOT DEFINED ${foundVar} OR NOT ${${foundVar}})
                        message( WARNING "Found package ${ARG_PACKAGE_NAME}, but component ${component} is not locally available" )
                        unset(${ARG_PACKAGE_NAME}_FOUND) # force attempt to fetch
                    endif()
                endforeach()
                if(NOT ${ARG_PACKAGE_NAME}_FOUND)
                    message( WARNING ">>>>>>>>>>> manage_dependency: after call to find_package, the cmake configuration is polluted with settings for ${ARG_PACKAGE_NAME}."
                                     " An automatic fetch will be attempted but might fail due to that. Install necessary dependencies yourself, if needed" )
                endif()
            endif()
        endif()

        if(PREFER_STATIC OR BUILD_SHARED_LIBS)
            set(CMAKE_FIND_LIBRARY_SUFFIXES ${SAVED_CMAKE_FIND_LIBRARY_SUFFIXES})   # restore CMAKE_FIND_LIBRARY_SUFFIXES
        endif()

        if(${ARG_PACKAGE_NAME}_FOUND AND TARGET ${ARG_TARGET_NAME_SYSTEM}) # 2026-04-04 check for both package name found and expected target
            if(PREFER_STATIC)
                set(${ARG_LIB_NAME}_TARGET_STATIC "$<TARGET_FILE:${ARG_TARGET_NAME}>" PARENT_SCOPE)
            endif()

            set(should_fetch FALSE)
            set(${ARG_LIB_NAME}_PROVIDED_TARGET             # ${ARG_LIB_NAME}_PROVIDED_TARGET should be the name as provided by an installed library
                ${ARG_TARGET_NAME_SYSTEM} PARENT_SCOPE)     #
            set(ARG_TARGET_NAME ${ARG_TARGET_NAME_SYSTEM})  # also, the variable ${ARG_TARGET_NAME_SYSTEM}_FOUND will be set to TRUE, so adjust ARG_TARGET_NAME
            message(STATUS ">>>>>>>>>>> manage_dependency: manage_dependency: Found system ${ARG_LIB_NAME}: ${${ARG_PACKAGE_NAME}_VERSION}")
        else()
            set(should_fetch ${FETCH_DEPS_AUTO})       # BUGFIX 2026-01-25: was previously not evaluating variable FETCH_DEPS_AUTO
        endif()
    endif()

    # Fetch if needed
    if(should_fetch)
        if( NOT "${ARG_GITHUB_REPO}" STREQUAL "" )
            set( ARG_GIT_REPOSITORY "https://github.com/${ARG_GITHUB_REPO}.git" )
        endif()

        include(FetchContent)

# message( NOTICE "FetchContent_Declare_custom(" )
# message( NOTICE "    LIB_NAME       ${ARG_LIB_NAME}_fetch" )
# message( NOTICE "    URL            ${ARG_URL}" )
# message( NOTICE "    URL_HASH       ${ARG_URL_HASH}" )
# message( NOTICE "    GIT_REPOSITORY ${ARG_GIT_REPOSITORY}" )
# message( NOTICE "    GIT_TAG        ${ARG_GIT_TAG}" )
# message( NOTICE "    EXCLUDE_FROM_ALL" )
# message( NOTICE ")" )
        FetchContent_Declare_custom(
                LIB_NAME       ${ARG_LIB_NAME}_fetch
                URL            ${ARG_URL}
                URL_HASH       ${ARG_URL_HASH}
                GIT_REPOSITORY ${ARG_GIT_REPOSITORY}
                GIT_TAG        ${ARG_GIT_TAG}
                EXCLUDE_FROM_ALL
        )

        # Set appropriate build type
        if(PREFER_STATIC)
            set(${ARG_LIB_NAME}_BUILD_SHARED_LIBS OFF CACHE BOOL "" FORCE)
        else()
            set(${ARG_LIB_NAME}_BUILD_SHARED_LIBS ON CACHE BOOL "" FORCE)
        endif()

        # Make available but don't build yet
        FetchContent_MakeAvailable_custom(
            LIB_NAME         ${ARG_LIB_NAME}_fetch
            EXCLUDE_FROM_ALL
        )

        if(PREFER_STATIC)
            # set return value for path to static library, if any
            set(${ARG_LIB_NAME}_TARGET_STATIC "$<TARGET_FILE:${ARG_TARGET_NAME}>" PARENT_SCOPE)
        endif()
        set(${ARG_LIB_NAME}_FETCHED TRUE PARENT_SCOPE)  # set return value indicating the dependency was fetched from source repository
    endif()

    # Verify target exists
    if(NOT TARGET "${ARG_TARGET_NAME}")
        message(FATAL_ERROR
            "manage_dependency: Dependency ${ARG_LIB_NAME} (target ${ARG_TARGET_NAME}) not available. "
            "Check your system installation or enable FETCH_DEPS_AUTO."
        )
    else() # target exists
        get_target_property(TARGET_IS_IMPORTED "${ARG_TARGET_NAME}" IMPORTED)
        if(TARGET_IS_IMPORTED)
            message( DEBUG ">>>>>>>>>>> manage_dependency: target ${ARG_TARGET_NAME} is imported" )
            set(${ARG_LIB_NAME}_INSTALL_TARGET "" PARENT_SCOPE)  # imported targets do not need to be installed
        else()
            message( DEBUG ">>>>>>>>>>> manage_dependency: target ${ARG_TARGET_NAME} is not imported and must be installed" )
            # unfortunately, ARG_TARGET_NAME can be an alias, which does not behave well with cmake install
            # so: get the aliased target, if any
            get_target_property(ALIASED_INSTALL_TARGET "${ARG_TARGET_NAME}" ALIASED_TARGET)
            if(ALIASED_INSTALL_TARGET)
                message( STATUS ">>>>>>>>>>> manage_dependency: target ${ARG_TARGET_NAME} is an alias, resolved to ${ALIASED_INSTALL_TARGET}" )
                set(${ARG_LIB_NAME}_INSTALL_TARGET "${ALIASED_INSTALL_TARGET}" PARENT_SCOPE)
            else()
                message( DEBUG ">>>>>>>>>>> manage_dependency: target ${ARG_TARGET_NAME} is already the final target" )
                set(${ARG_LIB_NAME}_INSTALL_TARGET "${ARG_TARGET_NAME}" PARENT_SCOPE)
            endif()

        endif()
    endif()

    # restore saved values of FIND_PACKAGE_VARIABLES - FPC_NAME and FPC_VALUE are still expected to be unset when execution gets here
    foreach( val IN LISTS ARG_FIND_PACKAGE_VARIABLES )
        if(NOT FPC_NAME)
            set(FPC_NAME ${val})
        else()
            set(FPC_VALUE ${val})
            set(${FPC_NAME} ${${FPC_NAME}_OLD_VALUE})   # restore original value
            message( STATUS ">>>>>>>>>>> manage_dependency: restoring FIND_PACKAGE_VARIABLES ${FPC_NAME}=${${FPC_NAME}}" )
            unset( FPC_NAME )
            unset( FPC_VALUE )
        endif()
    endforeach()
endfunction()
# End of Function:   manage_dependency


# Function: target_link_interface
# Purpose:  link an interface library that does not inherit an install target from dependency
# Author:   Lars Uffmann (coding23@uffmann.name)
# NOTE: INTERFACE_NAME has to be included in install TARGETS by the caller
function(target_link_interface)
    set(options PUBLIC PRIVATE) # one option must be explicitly indicated:
    #                           #   PRIVATE: users of ${ARG_INTERFACE_NAME} shall not use features of ${ARG_DEPENDENCY_ALIAS}
    #                           #   PUBLIC:  users of ${ARG_INTERFACE_NAME} may use features of ${ARG_DEPENDENCY_ALIAS}
    set(oneValueArgs
        LIB_NAME           # the library that shall be linked versus an interface to dependency
        INTERFACE_NAME     # name of the interface library to be created (will be needed by caller in install targets)
        DEPENDENCY_ALIAS   # the name (can be alias) of the actual underlying dependency
        DEPENDENCY_TARGET  # the name (not alias) of the underlying dependency
    )
    set(multiValueArgs EXTRA_ARGS)
    cmake_parse_arguments(ARG "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})

    if(ARG_PUBLIC AND ARG_PRIVATE)
        message( FATAL_ERROR "target_link_interface: flags PUBLIC and PRIVATE are mutually exclusive")
    endif()
    if(NOT (ARG_PUBLIC OR ARG_PRIVATE))
        message( FATAL_ERROR "target_link_interface: PUBLIC or PRIVATE must be specified")
    endif()

    # Ridiculous workaround part 1: Define an IMPORTED library to properly inherit INTERFACE_INCLUDE_DIRECTORIES from DEPENDENCY_NAME
    set(IMPORTED_NAME "${ARG_INTERFACE_NAME}_imported")  # NOTE: this name is assumed to be unique within project
    add_library(${IMPORTED_NAME} STATIC IMPORTED)
    set_target_properties(${IMPORTED_NAME} PROPERTIES
        # What we do not care about: imported location
        # IMPORTED_LOCATION "" # using a FAKE target that exists, the imported library can be used to fetch nowide INTERFACE_INCLUDE_DIRECTORIES
        # to omit the requirement for this fake target, a second abstraction level via Ridiculous workaround part 2 is needed

        # What we are actually interested in: interface include directories and compile definitions
        INTERFACE_INCLUDE_DIRECTORIES "$<TARGET_PROPERTY:${ARG_DEPENDENCY_ALIAS},INTERFACE_INCLUDE_DIRECTORIES>"
        INTERFACE_COMPILE_DEFINITIONS "$<TARGET_PROPERTY:${ARG_DEPENDENCY_ALIAS},INTERFACE_COMPILE_DEFINITIONS>"
    )

    # Ridiculous workaround part 2: Define an INTERFACE library to the IMPORTED ${IMPORTED_NAME}, while eliminating the need for an IMPORTED_LOCATION
    add_library(${ARG_INTERFACE_NAME} INTERFACE)
    # NOTE: INTERFACE_NAME will expose INTERFACE_INCLUDE_DIRECTORIES and INTERFACE_COMPILE_DEFINITIONS of DEPENDENCY_ALIAS via the IMPORTED library
    target_include_directories(${ARG_INTERFACE_NAME} INTERFACE $<TARGET_PROPERTY:${IMPORTED_NAME},INTERFACE_INCLUDE_DIRECTORIES> )
    target_compile_definitions(${ARG_INTERFACE_NAME} INTERFACE $<TARGET_PROPERTY:${IMPORTED_NAME},INTERFACE_COMPILE_DEFINITIONS> )
    # CAUTION: do *NOT* use target_link_libraries to link to the IMPORTED library, this would re-activate the need for an IMPORTED_LOCATION

    # Ridiculous workaround part 3: Link LIB_NAME against the INTERFACE library
    if(ARG_PUBLIC)
        set( ARG_PRIVACY "PUBLIC" )
    else()
        set( ARG_PRIVACY "PRIVATE" )   # begin of function has asserted that either ARG_PUBLIC or ARG_PRIVATE must be true
    endif()
    target_link_libraries(${ARG_LIB_NAME}
        ${ARG_PRIVACY}        # PRIVATE means users of ${ARG_INTERFACE_NAME} shall not use features of ${ARG_DEPENDENCY_ALIAS}
        ${ARG_INTERFACE_NAME} # this links INTERFACE_INCLUDE_DIRECTORIES
    )

    # Ridiculous workaround part 4: link to the built archive file so that ARG_LIB_NAME does not require DEPENDENCY_TARGET to be in the export set
    target_link_libraries(${ARG_LIB_NAME}
        ${ARG_PRIVACY}  # TBD: would always PRIVATE mangle symbols so that users can not access them from the static LIB_NAME?
        $<TARGET_FILE:${ARG_DEPENDENCY_TARGET}>
    )

    # Ridiculous workaround part 5: ensure DEPENDENCY_ALIAS is built before LIB_NAME (so that archive exists at link time)
    add_dependencies(${ARG_LIB_NAME} ${ARG_DEPENDENCY_ALIAS})
endfunction()
# End of Function:   target_link_interface
