CPMFindPackage(
        NAME Boost
        VERSION 1.73.0
#        GITHUB_REPOSITORY boostorg/nowide
        GIT_REPOSITORY https://github.com/boostorg/nowide.git
        # NAME nowide
        GIT_TAG standalone
#       OPTIONS
#       "NOWIDE_INSTALL ON"
#       FIND_PACKAGE_ARGUMENTS "COMPONENTS nowide"
    )

include_object_library(nowide nowide-obj)
