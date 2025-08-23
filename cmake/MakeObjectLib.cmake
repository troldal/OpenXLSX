function(make_object_library ORIG_TARGET OUT_OBJ_TARGET)
    if (NOT TARGET ${ORIG_TARGET})
        message(FATAL_ERROR "make_object_library: target '${ORIG_TARGET}' not found")
    endif()

    # Grab the target's sources and its own source dir
    get_target_property(_srcs     ${ORIG_TARGET} SOURCES)
    get_target_property(_src_dir  ${ORIG_TARGET} SOURCE_DIR)

    # Convert any plain relative paths to absolute, leave generator expressions untouched
    set(_abs_srcs "")
    foreach(s IN LISTS _srcs)
        if (s MATCHES "\\$<")  # contains a generator expression -> keep as-is
            list(APPEND _abs_srcs "${s}")
        elseif(NOT IS_ABSOLUTE "${s}")
            list(APPEND _abs_srcs "${_src_dir}/${s}")
        else()
            list(APPEND _abs_srcs "${s}")
        endif()
    endforeach()

    # Build the object wrapper
    add_library(${OUT_OBJ_TARGET} OBJECT)
    target_sources(${OUT_OBJ_TARGET} PRIVATE ${_abs_srcs})

    # Mirror compile settings so the objects are built like the original
    target_include_directories(${OUT_OBJ_TARGET} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},INCLUDE_DIRECTORIES>"
            "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
    target_compile_definitions(${OUT_OBJ_TARGET} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},COMPILE_DEFINITIONS>"
            "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_COMPILE_DEFINITIONS>")
    target_compile_options(${OUT_OBJ_TARGET} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},COMPILE_OPTIONS>"
            "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_COMPILE_OPTIONS>")
    target_compile_features(${OUT_OBJ_TARGET} PRIVATE
            "$<TARGET_PROPERTY:${ORIG_TARGET},COMPILE_FEATURES>")

    # Respect PIC setting (needed if you also build a shared OpenXLSX)
    get_target_property(_pic ${ORIG_TARGET} POSITION_INDEPENDENT_CODE)
    if (DEFINED _pic)
        set_property(TARGET ${OUT_OBJ_TARGET} PROPERTY POSITION_INDEPENDENT_CODE ${_pic})
    endif()
endfunction()

## Internal accumulators (directory-scope). Use CACHE INTERNAL so repeated calls persist.
#set(BUNDLED_DEP_OBJECTS "" CACHE INTERNAL "Bundled dependency objects")
#set(BUNDLED_DEP_INCLUDE_DIRS "" CACHE INTERNAL "Bundled dependency includes")
#
## Resolve ALIAS to real target name
#function(_resolve_real_target IN OUT_VAR)
#    get_target_property(_type ${IN} TYPE)
#    if(_type STREQUAL "ALIAS")
#        get_target_property(_real ${IN} ALIASED_TARGET)
#        set(${OUT_VAR} "${_real}" PARENT_SCOPE)
#    else()
#        set(${OUT_VAR} "${IN}" PARENT_SCOPE)
#    endif()
#endfunction()
#
#function(make_object_library ORIG_TARGET OUT_OBJ_TARGET)
#    if (NOT TARGET ${ORIG_TARGET})
#        message(FATAL_ERROR "make_object_library: target '${ORIG_TARGET}' not found")
#    endif()
#
#    _resolve_real_target(${ORIG_TARGET} _REAL)
#    get_target_property(_type ${_REAL} TYPE)
#    if(_type STREQUAL "INTERFACE_LIBRARY")
#        # Nothing to compile; let caller just harvest includes.
#        return()
#    endif()
#
#    # Grab sources and source dir
#    get_target_property(_srcs     ${_REAL} SOURCES)
#    get_target_property(_src_dir  ${_REAL} SOURCE_DIR)
#
#    # Convert relative src -> absolute, leave genex as-is
#    set(_abs_srcs "")
#    foreach(s IN LISTS _srcs)
#        if(s MATCHES "\\$<") # genex
#            list(APPEND _abs_srcs "${s}")
#        elseif(NOT IS_ABSOLUTE "${s}")
#            list(APPEND _abs_srcs "${_src_dir}/${s}")
#        else()
#            list(APPEND _abs_srcs "${s}")
#        endif()
#    endforeach()
#
#    # Build the object wrapper
#    add_library(${OUT_OBJ_TARGET} OBJECT)
#    target_sources(${OUT_OBJ_TARGET} PRIVATE ${_abs_srcs})
#
#    # Mirror compile settings so the objects are built like the original
#    target_include_directories(${OUT_OBJ_TARGET} PRIVATE
#            "$<TARGET_PROPERTY:${_REAL},INCLUDE_DIRECTORIES>"
#            "$<TARGET_PROPERTY:${_REAL},INTERFACE_INCLUDE_DIRECTORIES>")
#    target_compile_definitions(${OUT_OBJ_TARGET} PRIVATE
#            "$<TARGET_PROPERTY:${_REAL},COMPILE_DEFINITIONS>"
#            "$<TARGET_PROPERTY:${_REAL},INTERFACE_COMPILE_DEFINITIONS>")
#    target_compile_options(${OUT_OBJ_TARGET} PRIVATE
#            "$<TARGET_PROPERTY:${_REAL},COMPILE_OPTIONS>"
#            "$<TARGET_PROPERTY:${_REAL},INTERFACE_COMPILE_OPTIONS>")
#    target_compile_features(${OUT_OBJ_TARGET} PRIVATE
#            "$<TARGET_PROPERTY:${_REAL},COMPILE_FEATURES>")
#
#    # Respect PIC from original
#    get_target_property(_pic ${_REAL} POSITION_INDEPENDENT_CODE)
#    if(DEFINED _pic)
#        set_property(TARGET ${OUT_OBJ_TARGET} PROPERTY POSITION_INDEPENDENT_CODE ${_pic})
#    endif()
#endfunction()
#
#function(include_object_library ORIG_TARGET OUT_OBJ_TARGET)
#    if (NOT TARGET ${ORIG_TARGET})
#        message(FATAL_ERROR "include_object_library: target '${ORIG_TARGET}' not found")
#    endif()
#
#    _resolve_real_target(${ORIG_TARGET} _REAL)
#    get_target_property(_type ${_REAL} TYPE)
#
#    # Always collect include dirs
#    set(_incs "${BUNDLED_DEP_INCLUDE_DIRS}")
#    list(APPEND _incs "$<TARGET_PROPERTY:${_REAL},INTERFACE_INCLUDE_DIRECTORIES>")
#    set(BUNDLED_DEP_INCLUDE_DIRS "${_incs}" CACHE INTERNAL "Bundled dependency includes")
#
#    # If it's INTERFACE library, there are no sources to wrap
#    if(_type STREQUAL "INTERFACE_LIBRARY")
#        return()
#    endif()
#
#    # Wrap into object library and collect its objects
#    make_object_library(${_REAL} ${OUT_OBJ_TARGET})
#
#    set(_objs "${BUNDLED_DEP_OBJECTS}")
#    list(APPEND _objs "$<TARGET_OBJECTS:${OUT_OBJ_TARGET}>")
#    set(BUNDLED_DEP_OBJECTS "${_objs}" CACHE INTERNAL "Bundled dependency objects")
#endfunction()
