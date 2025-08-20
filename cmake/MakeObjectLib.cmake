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
