function(make_object_library ORIG_TARGET OUT_OBJ_TARGET)
    if (NOT TARGET ${ORIG_TARGET})
        message(FATAL_ERROR "make_object_library: target '${ORIG_TARGET}' not found")
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
