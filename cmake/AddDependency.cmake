#function(add_dependency)
#    # add_dependency(
#    #   NAME <name>
#    #   TARGET <imported_target>
#    #   VERSION <version>
#    #   GIT_REPOSITORY <git_or_tarball_url>
#    #   GIT_TAG <tag_or_commit>        # optional for git GIT_REPOSITORY
#    #   FIND_PACKAGE_NAME <name>       # optional, defaults to NAME
#    #   FIND_PACKAGE_ARGS <args...>    # optional
#    #   OPTIONS <opt1;opt2;...>        # forwarded to CPMAddPackage(OPTIONS ...)
#    # )
#    set(options)
#    set(oneValueArgs NAME TARGET VERSION GIT_REPOSITORY FIND_PACKAGE_NAME GIT_TAG)
#    set(multiValueArgs FIND_PACKAGE_ARGS OPTIONS)
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
#    if (_mode STREQUAL "REMOTE")
#        CPMAddPackage(${_cpm_args})
#        if (NOT TARGET ${DEP_TARGET})
#            message(FATAL_ERROR "CPM fetched '${DEP_NAME}', but target '${DEP_TARGET}' is missing.")
#        endif ()
#
#        if (_bundle)
#            # Create a unique object target name
#            set(_obj_tgt "${DEP_NAME}__obj")
#            make_object_library(${DEP_TARGET} ${_obj_tgt})
#        endif ()
#
#    elseif (_mode STREQUAL "LOCAL")
#        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} REQUIRED)
#        if (NOT TARGET ${DEP_TARGET})
#            message(FATAL_ERROR "Local package '${DEP_FIND_PACKAGE_NAME}' found but target '${DEP_TARGET}' not provided.")
#        endif ()
#
#        # Intentionally do NOT bundle local packages as objects here.
#        # If you want this, allow it explicitly with <name>_BUNDLE_AS_OBJECT and call make_object_library.
#
#        if (_bundle)
#            # Only makes sense if local package is actually present and has sources
#            set(_obj_tgt "${DEP_NAME}__obj")
#            make_object_library(${DEP_TARGET} ${_obj_tgt})
#        endif ()
#
#    else () # AUTO: try local, else fetch
#        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} QUIET)
#        if(${DEP_FIND_PACKAGE_NAME}_FOUND)
#            message_color(STATUS "Found local package '${DEP_FIND_PACKAGE_NAME}' (target '${DEP_TARGET}') in AUTO mode." COLOR magenta)
#        else ()
#            # Scrub FindBoost cache variables to avoid NOTFOUND bleed-through
#            get_cmake_property(_allvars VARIABLES)
#            foreach(v IN LISTS _allvars)
#                if(v MATCHES "^Boost(_.*)?$")  # adjust prefix for your dep
#                    unset(${v} CACHE)
#                    unset(${v})
#                endif()
#            endforeach()
#
#            message_color(STATUS "Local package '${DEP_FIND_PACKAGE_NAME}' not found. Fetching '${DEP_NAME}' via CPM in AUTO mode." COLOR magenta)
#            CPMAddPackage(${_cpm_args})
#            if (NOT TARGET ${DEP_TARGET})
#                message(FATAL_ERROR
#                        "Failed to obtain '${DEP_NAME}' in AUTO mode; target '${DEP_TARGET}' not available after CPM fetch.")
#            endif ()
#            if (_bundle)
#                set(_obj_tgt "${DEP_NAME}__obj")
#                make_object_library(${DEP_TARGET} ${_obj_tgt})
#            endif ()
#        endif ()
#    endif ()
#endfunction()


function(add_dependency)
    # add_dependency(
    #   NAME <name>
    #   TARGET <imported_target>
    #   VERSION <version>
    #   GIT_REPOSITORY <git_or_tarball_url>
    #   GIT_TAG <tag_or_commit>        # optional for git GIT_REPOSITORY
    #   FIND_PACKAGE_NAME <name>       # optional, defaults to NAME
    #   FIND_PACKAGE_ARGS <args...>    # optional
    #   OPTIONS <opt1;opt2;...>        # forwarded to CPMAddPackage(OPTIONS ...)
    #   DEP_LIBRARY_LIST <listvar>     # optional: list var name in caller to append the lib target
    #   DEP_OBJECT_LIST  <listvar>     # optional: list var name in caller to append the obj target (if bundled)
    #   DEP_HEADER_LIST  <listvar>     # optional: list var name in caller to append header dirs
    # )

    set(options)
    set(oneValueArgs
            NAME TARGET VERSION GIT_REPOSITORY FIND_PACKAGE_NAME GIT_TAG
            PACKAGE_LIST LIBRARY_LIST OBJECT_LIST HEADER_LIST
    )
    set(multiValueArgs FIND_PACKAGE_ARGS FIND_PACKAGE_OPTIONS OPTIONS)
    cmake_parse_arguments(DEP "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})

    if (NOT DEP_NAME OR NOT DEP_TARGET OR NOT DEP_GIT_REPOSITORY)
        message(FATAL_ERROR "add_dependency requires NAME, TARGET, and GIT_REPOSITORY")
    endif ()

    if (NOT DEP_FIND_PACKAGE_NAME)
        set(DEP_FIND_PACKAGE_NAME "${DEP_NAME}")
    endif ()

    # Mode resolution (global, then per-dep override)
    set(_mode "${DEP_MODE}")
    if (DEFINED ${DEP_NAME}_DEP_MODE)
        set(_mode "${${DEP_NAME}_DEP_MODE}")
    endif ()
    string(TOUPPER "${_mode}" _mode)

    # Whether to bundle this dep as objects
    set(_bundle "${BUNDLE_REMOTE_DEPS}")
    if (DEFINED ${DEP_NAME}_BUNDLE_AS_OBJECT)
        set(_bundle "${${DEP_NAME}_BUNDLE_AS_OBJECT}")
    endif ()

    # Build CPMAddPackage argument list
    set(_cpm_args
            NAME ${DEP_NAME}
            VERSION ${DEP_VERSION}
            GIT_REPOSITORY ${DEP_GIT_REPOSITORY}
    )
    if (DEP_GIT_TAG)
        list(APPEND _cpm_args GIT_TAG ${DEP_GIT_TAG})
    endif ()
    if (DEP_OPTIONS)
        list(APPEND _cpm_args OPTIONS ${DEP_OPTIONS})
    endif ()



    # We’ll fill these and then append to the caller lists (if requested)
    set(_chosen_lib_target "")
    set(_created_obj_target "")
    set(_dep_includes "")

    if (_mode STREQUAL "REMOTE")
        CPMAddPackage(${_cpm_args})
        if (NOT TARGET ${DEP_TARGET})
            message(FATAL_ERROR "CPM fetched '${DEP_NAME}', but target '${DEP_TARGET}' is missing.")
        endif ()
        set(_chosen_lib_target "${DEP_TARGET}")
        get_target_property(_chosen_lib_target_type "${_chosen_lib_target}" TYPE)
        if (_bundle AND NOT _chosen_lib_target_type STREQUAL "INTERFACE_LIBRARY")
            set(_created_obj_target "${DEP_NAME}__obj")
            make_object_library(${DEP_TARGET} ${_created_obj_target})

            set(_objs "${${DEP_OBJECT_LIST}}")
            list(APPEND _objs "$<TARGET_OBJECTS:${_created_obj_target}>")
            set(${DEP_OBJECT_LIST} "${_objs}" PARENT_SCOPE)

            set(_incs "${${DEP_HEADER_LIST}}")
            list(APPEND _incs "$<TARGET_PROPERTY:${DEP_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
            set(${DEP_HEADER_LIST} "${_incs}" PARENT_SCOPE)

        else()

            set(_tgt "${${DEP_PACKAGE_LIST}}")
            list(APPEND _tgt "${DEP_FIND_PACKAGE_NAME}")
            set(${DEP_PACKAGE_LIST} "${_tgt}" PARENT_SCOPE)

            # If not bundling, just add the library target to the caller list (if requested)
            if (DEP_LIBRARY_LIST)
                set(_libs "${${DEP_LIBRARY_LIST}}")
                list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
                set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
            endif ()


        endif ()

    elseif (_mode STREQUAL "LOCAL")
        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} ${DEP_FIND_PACKAGE_OPTIONS} QUIET)
        if (NOT TARGET ${DEP_TARGET})
            message(FATAL_ERROR "Local package '${DEP_FIND_PACKAGE_NAME}' found but target '${DEP_TARGET}' not provided.")
        endif ()
        set(_chosen_lib_target "${DEP_TARGET}")

#        if (_bundle)
#            set(_created_obj_target "${DEP_NAME}__obj")
#            make_object_library(${DEP_TARGET} ${_created_obj_target})
#        endif ()
        if (DEP_LIBRARY_LIST)
            set(_libs "${${DEP_LIBRARY_LIST}}")
            list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
            set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
        endif ()

    else () # AUTO: try local, else fetch
        find_package(${DEP_FIND_PACKAGE_NAME} ${DEP_VERSION} ${DEP_FIND_PACKAGE_ARGS} ${DEP_FIND_PACKAGE_OPTIONS} QUIET)
        if(${DEP_FIND_PACKAGE_NAME}_FOUND)
            message_color(STATUS
                    "Found local package '${DEP_FIND_PACKAGE_NAME}' (target '${DEP_TARGET}') in AUTO mode."
                    COLOR magenta)

            set(_chosen_lib_target "${DEP_TARGET}")
#            if (_bundle)
#                set(_created_obj_target "${DEP_NAME}__obj")
#                make_object_library(${DEP_TARGET} ${_created_obj_target})
#            endif ()
            if (DEP_LIBRARY_LIST)
                set(_libs "${${DEP_LIBRARY_LIST}}")
                list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
                set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
            endif ()

        else ()
            # Scrub (example: Boost) to avoid NOTFOUND bleed-through; adjust if needed
            get_cmake_property(_allvars VARIABLES)
            foreach(v IN LISTS _allvars)
                if(v MATCHES "^Boost(_.*)?$")
                    unset(${v} CACHE)
                    unset(${v})
                endif()
            endforeach()



            message_color(
                    STATUS
                        "Local package '${DEP_FIND_PACKAGE_NAME}' not found. Fetching '${DEP_NAME}' via CPM in AUTO mode."
                    COLOR
                        magenta)
            CPMAddPackage(${_cpm_args})
            if (NOT TARGET ${DEP_TARGET})
                message(FATAL_ERROR
                        "Failed to obtain '${DEP_NAME}' in AUTO mode; target '${DEP_TARGET}' not available after CPM fetch.")
            endif ()
            set(_chosen_lib_target "${DEP_TARGET}")
            get_target_property(_chosen_lib_target_type "${_chosen_lib_target}" TYPE)
            if (_bundle AND NOT _chosen_lib_target_type STREQUAL "INTERFACE_LIBRARY")
                set(_created_obj_target "${DEP_NAME}__obj")
                make_object_library(${DEP_TARGET} ${_created_obj_target})

                set(_objs "${${DEP_OBJECT_LIST}}")
                list(APPEND _objs "$<TARGET_OBJECTS:${_created_obj_target}>")
                set(${DEP_OBJECT_LIST} "${_objs}" PARENT_SCOPE)

                set(_incs "${${DEP_HEADER_LIST}}")
                list(APPEND _incs "$<TARGET_PROPERTY:${DEP_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
                set(${DEP_HEADER_LIST} "${_incs}" PARENT_SCOPE)
            else()

                set(_tgt "${${DEP_PACKAGE_LIST}}")
                list(APPEND _tgt "${DEP_FIND_PACKAGE_NAME}")
                set(${DEP_PACKAGE_LIST} "${_tgt}" PARENT_SCOPE)

                # If not bundling, just add the library target to the caller list (if requested)
                if (DEP_LIBRARY_LIST)
                    set(_libs "${${DEP_LIBRARY_LIST}}")
                    list(APPEND _libs "$<BUILD_LOCAL_INTERFACE:${_chosen_lib_target}>")
                    set(${DEP_LIBRARY_LIST} "${_libs}" PARENT_SCOPE)
                endif ()
            endif ()
        endif ()
    endif ()

    # Gather header directories from the chosen target (if it has any)
    if(TARGET "${_chosen_lib_target}")
        get_target_property(_dep_includes "${_chosen_lib_target}" INTERFACE_INCLUDE_DIRECTORIES)
        if(_dep_includes)
            # Keep as-is (may include generator expressions)
            list(REMOVE_DUPLICATES _dep_includes)
        else()
            set(_dep_includes "")
        endif()
    endif()
endfunction()

