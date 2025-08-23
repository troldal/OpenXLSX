CPMFindPackage(
        NAME miniz
        URI https://github.com/richgel999/miniz.git
        GIT_TAG 3.0.2
        OPTIONS
        "AMALGAMATE_SOURCES ON
         BUILD_HEADER_ONLY OFF
         BUILD_SHARED_LIBS OFF"
        # use of reserved BUILD_SHARED_LIBS automatically passes this option through
    )

include_object_library(miniz miniz-obj)
