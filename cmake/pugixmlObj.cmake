CPMFindPackage(
        NAME pugixml
        GITHUB_REPOSITORY zeux/pugixml
        GIT_TAG v1.15
        OPTIONS
        "PUGIXML_COMPACT ${OPENXLSX_COMPACT_MODE}"
        # use of reserved BUILD_SHARED_LIBS automatically passes this option through
)

# Suppose the zlib target is 'zlibstatic' (common on Windows builds) or 'zlib'
if (TARGET pugixml-static)
    make_object_library(pugixml-static pugixml-obj)
else()
    message(FATAL_ERROR "Couldn't find a pugixml target to wrap.")
endif()