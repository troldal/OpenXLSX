    CPMFindPackage(
            NAME Boost
            VERSION 1.73.0
            GITHUB_REPOSITORY boostorg/nowide
            # NAME nowide
            GIT_TAG standalone
#            OPTIONS
#            "NOWIDE_INSTALL ON"
#            FIND_PACKAGE_ARGUMENTS "COMPONENTS nowide"
    )

    if (TARGET nowide)
        make_object_library(nowide nowide-obj)
        list(APPEND OPENXLSX_DEPENDENCY_OBJECTS "$<TARGET_OBJECTS:nowide-obj>")
        list(APPEND OPENXLSX_DEPENDENCY_HEADERS "$<TARGET_PROPERTY:nowide,INTERFACE_INCLUDE_DIRECTORIES>")
    else()
        message(FATAL_ERROR "Couldn't find a nowide target to wrap.")
    endif()