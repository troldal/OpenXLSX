#function(add_dependency)
#    set(options)
#    set(oneValueArgs
#            NAME TARGET VERSION GIT_REPOSITORY FIND_PACKAGE_NAME GIT_TAG
#            PACKAGE_LIST LIBRARY_LIST OBJECT_LIST HEADER_LIST
#    )
#    set(multiValueArgs FIND_PACKAGE_ARGS FIND_PACKAGE_OPTIONS OPTIONS)
#    cmake_parse_arguments(DEP "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})
#
#    if (NOT DEP_NAME OR NOT DEP_TARGET OR NOT DEP_GIT_REPOSITORY)
#        message(FATAL_ERROR "add_dependency requires NAME, TARGET, and GIT_REPOSITORY")
#    endif ()
#
#    if (NOT DEP_FIND_PACKAGE_NAME)
#        set(DEP_FIND_PACKAGE_NAME "${DEP_NAME}")
#    endif ()
#
#    # Mode resolution (global, then per-dep override)
#    set(_mode "${DEP_MODE}")
#    if (DEFINED ${DEP_NAME}_DEP_MODE)
#        set(_mode "${${DEP_NAME}_DEP_MODE}")
#    endif ()
#    string(TOUPPER "${_mode}" _mode)
#
#    # Whether to bundle this dep as objects
#    set(_bundle "${BUNDLE_REMOTE_DEPS}")
#    if (DEFINED ${DEP_NAME}_BUNDLE_AS_OBJECT)
#        set(_bundle "${${DEP_NAME}_BUNDLE_AS_OBJECT}")
#    endif ()
#
#    # Build CPMAddPackage argument list
#    set(_cpm_args
#            NAME ${DEP_NAME}
#            VERSION ${DEP_VERSION}
#            GIT_REPOSITORY ${DEP_GIT_REPOSITORY}
#    )
#    if (DEP_GIT_TAG)
#        list(APPEND _cpm_args GIT_TAG ${DEP_GIT_TAG})
#    endif ()
#    if (DEP_OPTIONS)
#        list(APPEND _cpm_args OPTIONS ${DEP_OPTIONS})
#    endif ()
#
#
#
#    # We’ll fill these and then append to the caller lists (if requested)
#    set(_chosen_lib_target "")
#    set(_created_obj_target "")
#    set(_dep_includes "")
#
#    if (_mode STREQUAL "REMOTE")
#        CPMAddPackage(${_cpm_args})
#        if (NOT TARGET ${DEP_TARGET})
#            message(FATAL_ERROR "CPM fetched '${DEP_NAME}', but target '${DEP_TARGET}' is missing.")
#        endif ()
#        set(_chosen_lib_target "${DEP_TARGET}")
#        get_target_property(_chosen_lib_target_type "${_chosen_lib_target}" TYPE)
#        if (_bundle AND NOT _chosen_lib_target_type STREQUAL "INTERFACE_LIBRARY")
#            set(_created_obj_target "${DEP_NAME}__obj")
#            make_object_library(${DEP_TARGET} ${_created_obj_target})
#
#            set(_objs "${${DEP_OBJECT_LIST}}")
#            list(APPEND _objs "$<TARGET_OBJECTS:${_created_obj_target}>")
#            set(${DEP_OBJECT_LIST} "${_objs}" PARENT_SCOPE)
#
#            set(_incs "${${DEP_HEADER_LIST}}")
#            list(APPEND _incs "$<TARGET_PROPERTY:${DEP_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
#            set(${DEP_HEADER_LIST} "${_incs}" PARENT_SCOPE)
#
#        else()
#
#            set(_tgt "${${DEP_PACKAGE_LIST}}")
#            list(APPEND _tgt "${DEP_FIND_PACKAGE_NAME}")
#            set(${DEP_PACKAGE_LIST} "${_tgt}" PARENT_SCOPE)
#
#            # If not bundling, just add the library target to the caller list (if requested)
#            if (DEP_LIBRARY_LIST)
#                set(_libs "${${DEP_LIBRARY_LIST}}")
#                list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
#                set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
#            endif ()
#
#
#        endif ()
#
#    elseif (_mode STREQUAL "LOCAL")
#        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} ${DEP_FIND_PACKAGE_OPTIONS} QUIET)
#        if (NOT TARGET ${DEP_TARGET})
#            message(FATAL_ERROR "Local package '${DEP_FIND_PACKAGE_NAME}' found but target '${DEP_TARGET}' not provided.")
#        endif ()
#        set(_chosen_lib_target "${DEP_TARGET}")
#
#        if (DEP_LIBRARY_LIST)
#            set(_libs "${${DEP_LIBRARY_LIST}}")
#            list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
#            set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
#        endif ()
#
#    else () # AUTO: try local, else fetch
#        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} ${DEP_FIND_PACKAGE_OPTIONS} QUIET)
#        if(${DEP_FIND_PACKAGE_NAME}_FOUND)
#            message_color(STATUS
#                    "Found local package '${DEP_FIND_PACKAGE_NAME}' (target '${DEP_TARGET}') in AUTO mode."
#                    COLOR magenta)
#
#            set(_chosen_lib_target "${DEP_TARGET}")
#
#            if (DEP_LIBRARY_LIST)
#                set(_libs "${${DEP_LIBRARY_LIST}}")
#                list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
#                set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
#            endif ()
#
#        else ()
#            # Scrub (example: Boost) to avoid NOTFOUND bleed-through; adjust if needed
#            get_cmake_property(_allvars VARIABLES)
#            foreach(v IN LISTS _allvars)
#                if(v MATCHES "^Boost(_.*)?$")
#                    unset(${v} CACHE)
#                    unset(${v})
#                endif()
#            endforeach()
#
#
#
#            message_color(
#                    STATUS
#                        "Local package '${DEP_FIND_PACKAGE_NAME}' not found. Fetching '${DEP_NAME}' via CPM in AUTO mode."
#                    COLOR
#                        magenta)
#            CPMAddPackage(${_cpm_args})
#            if (NOT TARGET ${DEP_TARGET})
#                message(FATAL_ERROR
#                        "Failed to obtain '${DEP_NAME}' in AUTO mode; target '${DEP_TARGET}' not available after CPM fetch.")
#            endif ()
#            set(_chosen_lib_target "${DEP_TARGET}")
#            get_target_property(_chosen_lib_target_type "${_chosen_lib_target}" TYPE)
#            if (_bundle AND NOT _chosen_lib_target_type STREQUAL "INTERFACE_LIBRARY")
#                set(_created_obj_target "${DEP_NAME}__obj")
#                make_object_library(${DEP_TARGET} ${_created_obj_target})
#
#                set(_objs "${${DEP_OBJECT_LIST}}")
#                list(APPEND _objs "$<TARGET_OBJECTS:${_created_obj_target}>")
#                set(${DEP_OBJECT_LIST} "${_objs}" PARENT_SCOPE)
#
#                set(_incs "${${DEP_HEADER_LIST}}")
#                list(APPEND _incs "$<TARGET_PROPERTY:${DEP_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
#                set(${DEP_HEADER_LIST} "${_incs}" PARENT_SCOPE)
#            else()
#
#                set(_tgt "${${DEP_PACKAGE_LIST}}")
#                list(APPEND _tgt "${DEP_FIND_PACKAGE_NAME}")
#                set(${DEP_PACKAGE_LIST} "${_tgt}" PARENT_SCOPE)
#
#                # If not bundling, just add the library target to the caller list (if requested)
#                if (DEP_LIBRARY_LIST)
#                    set(_libs "${${DEP_LIBRARY_LIST}}")
#                    list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
#                    set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
#                endif ()
#            endif ()
#        endif ()
#    endif ()
#
#    # Gather header directories from the chosen target (if it has any)
#    if(TARGET "${_chosen_lib_target}")
#        get_target_property(_dep_includes "${_chosen_lib_target}" INTERFACE_INCLUDE_DIRECTORIES)
#        if(_dep_includes)
#            # Keep as-is (may include generator expressions)
#            list(REMOVE_DUPLICATES _dep_includes)
#        else()
#            set(_dep_includes "")
#        endif()
#    endif()
#endfunction()


# add_dependency(...)
# Purpose:
#   Resolve a dependency in one of three modes: LOCAL, REMOTE, or AUTO (local-first, then fetch).
#   On success, it appends targets or artifacts to caller-provided lists (in the parent scope)
#   depending on bundling strategy.
#
# Inputs (parsed via cmake_parse_arguments as DEP_*):
#   - NAME                Logical dependency name (also default for FIND_PACKAGE_NAME).
#   - TARGET              CMake target expected/provided by the dependency.
#   - VERSION             Version constraint used for find_package/CPM.
#   - GIT_REPOSITORY      Repository URL for CPMAddPackage.
#   - GIT_TAG             Git tag/commit for CPMAddPackage (optional).
#   - FIND_PACKAGE_NAME   Package name used with find_package (defaults to NAME).
#   - FIND_PACKAGE_ARGS   Extra args forwarded to find_package(...) (e.g., REQUIRED is not used here).
#   - FIND_PACKAGE_OPTIONS Options forwarded to find_package(...) (e.g., CONFIG).
#   - OPTIONS             Extra options forwarded to CPMAddPackage(...).
#
# Behavior control (variables read from cache/scope):
#   - DEP_MODE / <NAME>_DEP_MODE
#       Resolution mode. Per-dependency variable <NAME>_DEP_MODE overrides DEP_MODE.
#       Accepted values (case-insensitive): LOCAL, REMOTE, AUTO.
#   - BUNDLE_REMOTE_DEPS / <NAME>_BUNDLE_AS_OBJECT
#       Whether to bundle fetched (non-interface) libraries into an OBJECT library that
#       exposes:
#         * ${DEP_OBJECT_LIST} += $<TARGET_OBJECTS:...>
#         * ${DEP_HEADER_LIST} += $<TARGET_PROPERTY:${DEP_TARGET},INTERFACE_INCLUDE_DIRECTORIES>
#       The per-dependency flag <NAME>_BUNDLE_AS_OBJECT overrides the global BUNDLE_REMOTE_DEPS.
#
# Outputs (appended in PARENT_SCOPE if the corresponding variable names are provided):
#   - ${DEP_PACKAGE_LIST}  += FIND_PACKAGE_NAME (when not bundling)
#   - ${DEP_LIBRARY_LIST}  += $<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}> (when not bundling)
#   - ${DEP_OBJECT_LIST}   += $<TARGET_OBJECTS:${created_obj_target}> (when bundling)
#   - ${DEP_HEADER_LIST}   += $<TARGET_PROPERTY:${DEP_TARGET},INTERFACE_INCLUDE_DIRECTORIES> (when bundling)
#
# Notes:
#   - Requires CPM.cmake's CPMAddPackage to be available in scope.
#   - Requires a helper function make_object_library(original_target, out_obj_target) in scope when bundling.
#   - Uses message_color(STATUS ...) if available for colored status output.
#   - Generator expressions are intentionally preserved in list entries for deferred evaluation.
function(add_dependency)
    # Parse arguments to DEP_*
    set(options)
    set(oneValueArgs NAME TARGET VERSION GIT_REPOSITORY FIND_PACKAGE_NAME GIT_TAG PACKAGE_LIST LIBRARY_LIST OBJECT_LIST HEADER_LIST)
    set(multiValueArgs FIND_PACKAGE_ARGS FIND_PACKAGE_OPTIONS OPTIONS)
    cmake_parse_arguments(DEP "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})

    # Minimal validation of required inputs
    if (NOT DEP_NAME OR NOT DEP_TARGET OR NOT DEP_GIT_REPOSITORY)
        message(FATAL_ERROR "add_dependency requires NAME, TARGET, and GIT_REPOSITORY")
    endif ()

    # Default the find_package name to NAME unless overridden
    if (NOT DEP_FIND_PACKAGE_NAME)
        set(DEP_FIND_PACKAGE_NAME "${DEP_NAME}")
    endif ()

    # Resolve mode: global DEP_MODE, then per-dep override <NAME>_DEP_MODE; normalize to upper-case
    set(_mode "${DEP_MODE}")
    if (DEFINED ${DEP_NAME}_DEP_MODE)
        set(_mode "${${DEP_NAME}_DEP_MODE}")
    endif ()
    string(TOUPPER "${_mode}" _mode)

    # Bundling decision: global BUNDLE_REMOTE_DEPS, then per-dep override <NAME>_BUNDLE_AS_OBJECT
    set(_bundle "${BUNDLE_REMOTE_DEPS}")
    if (DEFINED ${DEP_NAME}_BUNDLE_AS_OBJECT)
        set(_bundle "${${DEP_NAME}_BUNDLE_AS_OBJECT}")
    endif ()

    # Build CPMAddPackage(...) argument list
    set(_cpm_args NAME ${DEP_NAME} VERSION ${DEP_VERSION} GIT_REPOSITORY ${DEP_GIT_REPOSITORY})
    if (DEP_GIT_TAG)
        list(APPEND _cpm_args GIT_TAG ${DEP_GIT_TAG})
    endif ()
    if (DEP_OPTIONS)
        list(APPEND _cpm_args OPTIONS ${DEP_OPTIONS})
    endif ()

    # Internal state (local to this function)
    set(_chosen_lib_target "")   # The final target to link or bundle
    set(_created_obj_target "")  # When bundling, the name of the generated OBJECT library
    set(_dep_includes "")        # Collected include dirs (kept for compatibility)

    # Unified messaging helpers (ensure consistent and greppable messages across modes)
    # Note: message_color is assumed to exist in the project; falls back to message(...) otherwise.
    macro(_dep_status _msg)
        message_color(STATUS "Dependency '${DEP_NAME}': ${_msg} [mode: ${_mode}]" COLOR magenta)
    endmacro()
    macro(_dep_fatal _msg)
        message(FATAL_ERROR "Dependency '${DEP_NAME}': ${_msg} [mode: ${_mode}]")
    endmacro()

    # Append a single item to a parent-scope list variable name provided by the caller
    macro(_append_to_parent_list _var _item)
        set(_tmp "${${_var}}")
        list(APPEND _tmp "${_item}")
        set(${_var} "${_tmp}" PARENT_SCOPE)
    endmacro()

    # Convenience: record a library target for later link usage (non-bundled case)
    macro(_add_library_to_parent_list _target)
        _append_to_parent_list(${DEP_LIBRARY_LIST} "$<BUILD_LOCAL_INTERFACE:${_target}>")
    endmacro()

    # Convenience: record the package name that provided the target (non-bundled case)
    macro(_add_package_name)
        _append_to_parent_list(${DEP_PACKAGE_LIST} "${DEP_FIND_PACKAGE_NAME}")
    endmacro()

    # Bundle fetched library into an OBJECT library and expose objects and headers upward
    # Requires make_object_library(<src_target> <out_obj_target>) defined by the project.
    macro(_bundle_dep)
        set(_created_obj_target "${DEP_NAME}__obj")
        make_object_library(${DEP_TARGET} ${_created_obj_target})
        _append_to_parent_list(${DEP_OBJECT_LIST} "$<TARGET_OBJECTS:${_created_obj_target}>")
        _append_to_parent_list(${DEP_HEADER_LIST} "$<TARGET_PROPERTY:${DEP_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
    endmacro()

    # Finalize how the obtained target is surfaced:
    # - If bundling and target is not an interface library, create objects and expose headers.
    # - Otherwise, record the package name and add the target for linking.
    macro(_finalize_obtained_target)
        get_target_property(_chosen_lib_target_type "${_chosen_lib_target}" TYPE)
        if (_bundle AND NOT _chosen_lib_target_type STREQUAL "INTERFACE_LIBRARY")
            _bundle_dep()
        else ()
            _add_package_name()
            if (DEP_LIBRARY_LIST)
                _add_library_to_parent_list("${_chosen_lib_target}")
            endif ()
        endif ()
    endmacro()

    # Explicit REMOTE: always fetch via CPM
    if (_mode STREQUAL "REMOTE")
        CPMAddPackage(${_cpm_args})
        if (NOT TARGET ${DEP_TARGET})
            _dep_fatal("fetched package, but required target '${DEP_TARGET}' is not defined")
        endif ()
        set(_chosen_lib_target "${DEP_TARGET}")
        _dep_status("using REMOTE target '${_chosen_lib_target}'")
        _finalize_obtained_target()

    # Explicit LOCAL: must be available via find_package and provide the requested target
    elseif (_mode STREQUAL "LOCAL")
        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} ${DEP_FIND_PACKAGE_OPTIONS} QUIET)
        if (NOT ${DEP_FIND_PACKAGE_NAME}_FOUND)
            _dep_fatal("local package '${DEP_FIND_PACKAGE_NAME}' was not found")
        endif ()
        if (NOT TARGET ${DEP_TARGET})
            _dep_fatal("local package '${DEP_FIND_PACKAGE_NAME}' found, but required target '${DEP_TARGET}' is not defined")
        endif ()
        set(_chosen_lib_target "${DEP_TARGET}")
        _dep_status("using LOCAL target '${_chosen_lib_target}'")
        if (DEP_LIBRARY_LIST)
            _add_library_to_parent_list("${_chosen_lib_target}")
        endif ()

    # AUTO: prefer local; if not found, fetch remotely and then finalize
    else () # AUTO
        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} ${DEP_FIND_PACKAGE_OPTIONS} QUIET)
        if (${DEP_FIND_PACKAGE_NAME}_FOUND)
            if (NOT TARGET ${DEP_TARGET})
                _dep_fatal("local package '${DEP_FIND_PACKAGE_NAME}' found, but required target '${DEP_TARGET}' is not defined")
            endif ()
            set(_chosen_lib_target "${DEP_TARGET}")
            _dep_status("using LOCAL target '${_chosen_lib_target}'")
            if (DEP_LIBRARY_LIST)
                _add_library_to_parent_list("${_chosen_lib_target}")
            endif ()
        else ()
            # Clean specific cache entries that are known to bleed NOTFOUND between attempts (e.g., Boost)
            get_cmake_property(_allvars VARIABLES)
            foreach(v IN LISTS _allvars)
                if (v MATCHES "^Boost(_.*)?$")
                    unset(${v} CACHE)
                    unset(${v})
                endif ()
            endforeach ()

            _dep_status("local package '${DEP_FIND_PACKAGE_NAME}' not found; fetching from '${DEP_GIT_REPOSITORY}'")
            CPMAddPackage(${_cpm_args})
            if (NOT TARGET ${DEP_TARGET})
                _dep_fatal("fetched package, but required target '${DEP_TARGET}' is not defined")
            endif ()
            set(_chosen_lib_target "${DEP_TARGET}")
            _dep_status("using REMOTE target '${_chosen_lib_target}'")
            _finalize_obtained_target()
        endif ()
    endif ()

    # Collect include directories from the chosen target (compatibility; may be used by callers)
    if (TARGET "${_chosen_lib_target}")
        get_target_property(_dep_includes "${_chosen_lib_target}" INTERFACE_INCLUDE_DIRECTORIES)
        if (_dep_includes)
            list(REMOVE_DUPLICATES _dep_includes)
        else ()
            set(_dep_includes "")
        endif ()
    endif ()
endfunction()