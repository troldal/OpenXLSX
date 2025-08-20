    CPMFindPackage(
            NAME miniz
            GITHUB_REPOSITORY richgel999/miniz
            GIT_TAG 3.0.2
            OPTIONS
            "AMALGAMATE_SOURCES ON
            BUILD_HEADER_ONLY OFF
            BUILD_SHARED_LIBS OFF"
            # use of reserved BUILD_SHARED_LIBS automatically passes this option through
    )

    if (TARGET miniz)
        make_object_library(miniz miniz-obj)
        list(APPEND OPENXLSX_DEPENDENCY_OBJECTS "$<TARGET_OBJECTS:miniz-obj>")
        list(APPEND OPENXLSX_DEPENDENCY_HEADERS "$<TARGET_PROPERTY:miniz,INTERFACE_INCLUDE_DIRECTORIES>")
    else()
        message(FATAL_ERROR "Couldn't find a miniz target to wrap.")
    endif()