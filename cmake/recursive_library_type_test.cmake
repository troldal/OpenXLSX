# Function: recursive_library_type_test
# Purpose: test what library types a target is linking to, recurse into subtargets for INTERFACE_LIBRARY types
# Parameters:
#   TARGET_NAME     test this library target
# Returns:
#   ${TARGET_NAME}_HAS_STATIC   (PARENT_SCOPE): TRUE if at least one static library is linked
#   ${TARGET_NAME}_HAS_SHARED   (PARENT_SCOPE): TRUE if at least one shared library is linked
# Author:     Lars Uffmann (coding23@uffmann.name)
function(recursive_library_type_test TARGET_NAME PREFIX)
    set(_has_static FALSE)
    set(_has_shared FALSE)

    message( VERBOSE "      recursive_library_type_test:${PREFIX} TARGET ${TARGET_NAME}" )

    get_target_property(_type ${TARGET_NAME} TYPE)
    if("${_type}" STREQUAL "STATIC_LIBRARY")
        set(_has_static TRUE)
        message( DEBUG "      recursive_library_type_test:${PREFIX}   ${TARGET_NAME} is STATIC_LIBRARY" )
    elseif("${_type}" STREQUAL "SHARED_LIBRARY")
        set(_has_shared TRUE)
        message( DEBUG "      recursive_library_type_test:${PREFIX}   ${TARGET_NAME} is SHARED_LIBRARY" )
    elseif("${_type}" STREQUAL "INTERFACE_LIBRARY")
        get_target_property(_links ${TARGET_NAME} INTERFACE_LINK_LIBRARIES)
        message( DEBUG "      recursive_library_type_test:${PREFIX}   ${TARGET_NAME} is INTERFACE_LIBRARY, INTERFACE_LINK_LIBRARIES is ${_links}" )
        foreach(_link IN LISTS _links)
            if(TARGET ${_link})
                recursive_library_type_test( ${_link} "  ${PREFIX}" ) # recurse
                if(${_link}_HAS_STATIC)
                    set(_has_static TRUE)
                    message( DEBUG "      recursive_library_type_test:${PREFIX}   recursion found a static lib in ${_link}" )
                endif()
                if(${_link}_HAS_SHARED)
                    set(_has_shared TRUE)
                    message( DEBUG "      recursive_library_type_test:${PREFIX}   recursion found a shared lib in ${_link}" )
                endif()
                # message( DEBUG "   ${_link}_HAS_STATIC is ${${_link}_HAS_STATIC}, ${_link}_HAS_SHARED is ${${_link}_HAS_SHARED}" )
            elseif(IS_ABSOLUTE "${_link}")
                # _link is a path: check extension (.a/.lib => static; .so/.dylib/.dll => shared)
                message( FATAL_ERROR "recursive_library_type_test:${PREFIX}   absolute: ${_link} - still need to address this" )
            else()
                # _link might be a linker name (-lfoo) or generator expression; handle below
                message( FATAL_ERROR "recursive_library_type_test:${PREFIX}   linker name? ${_link} - still need to address this" )
            endif()
        endforeach()
    else()
    endif()

    message( VERBOSE "      recursive_library_type_test:${PREFIX} TARGET ${TARGET_NAME} _has_static is ${_has_static}, _has_shared is ${_has_shared}" )
    set(${TARGET_NAME}_HAS_STATIC ${_has_static} PARENT_SCOPE)
    set(${TARGET_NAME}_HAS_SHARED ${_has_shared} PARENT_SCOPE)
endfunction()
# End of Function: recursive_library_type_test
