#function(include_object_library ORIG_TARGET OUT_OBJ_TARGET)
#    make_object_library(${ORIG_TARGET} ${OUT_OBJ_TARGET})
#    list(APPEND OPENXLSX_DEPENDENCY_OBJECTS "$<TARGET_OBJECTS:${OUT_OBJ_TARGET}>")
#    list(APPEND OPENXLSX_DEPENDENCY_HEADERS "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
#endfunction()

function(include_object_library ORIG_TARGET OUT_OBJ_TARGET)
    make_object_library(${ORIG_TARGET} ${OUT_OBJ_TARGET})

    # Objects
    set(_objs "${OPENXLSX_DEPENDENCY_OBJECTS}")
    list(APPEND _objs "$<TARGET_OBJECTS:${OUT_OBJ_TARGET}>")
    set(OPENXLSX_DEPENDENCY_OBJECTS "${_objs}" PARENT_SCOPE)

    # Include dirs (keep as genex so it resolves at generate time)
    set(_incs "${OPENXLSX_DEPENDENCY_HEADERS}")
    list(APPEND _incs "$<TARGET_PROPERTY:${ORIG_TARGET},INTERFACE_INCLUDE_DIRECTORIES>")
    set(OPENXLSX_DEPENDENCY_HEADERS "${_incs}" PARENT_SCOPE)
endfunction()
