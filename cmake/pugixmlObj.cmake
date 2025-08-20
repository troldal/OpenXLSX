CPMFindPackage(
        NAME pugixml
        GITHUB_REPOSITORY zeux/pugixml
        GIT_TAG v1.15
        OPTIONS
        "PUGIXML_COMPACT ${OPENXLSX_COMPACT_MODE}"
        # use of reserved BUILD_SHARED_LIBS automatically passes this option through
)

include_object_library(pugixml-static pugixml-obj)
