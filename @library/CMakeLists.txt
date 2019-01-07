include(ExternalProject)

#=======================================================================================================================
# Download and build the PugiXML library
#=======================================================================================================================
ExternalProject_Add(pugi
        GIT_REPOSITORY https://github.com/zeux/pugixml.git
        GIT_TAG 8436f2a
        INSTALL_DIR ${OPENXLSX_DEPENDENCIES}
        UPDATE_COMMAND ""
        CMAKE_ARGS
            -DCMAKE_BUILD_TYPE=${CMAKE_BUILD_TYPE}
            -DCMAKE_C_COMPILER=${CMAKE_C_COMPILER}
            -DCMAKE_CXX_COMPILER=${CMAKE_CXX_COMPILER}
            -DCMAKE_POSITION_INDEPENDENT_CODE=ON
            -DCMAKE_INSTALL_PREFIX:PATH=<INSTALL_DIR>
            -DCMAKE_PREFIX_PATH:PATH=<INSTALL_DIR>
            -Wno-dev)

#=======================================================================================================================
# Download the header-only library 'Flat'
#=======================================================================================================================
ExternalProject_Add(flat
        GIT_REPOSITORY https://github.com/pubby/flat.git
        GIT_TAG bfa3b77
        INSTALL_DIR ${OPENXLSX_DEPENDENCIES}
        UPDATE_COMMAND ""
        CONFIGURE_COMMAND ""
        BUILD_COMMAND ""
        INSTALL_COMMAND
            ${CMAKE_COMMAND} -E copy_directory <SOURCE_DIR>/include ${OPENXLSX_DEPENDENCIES}/include/flat
        CMAKE_ARGS
            -DCMAKE_BUILD_TYPE=${CMAKE_BUILD_TYPE}
            -DCMAKE_C_COMPILER=${CMAKE_C_COMPILER}
            -DCMAKE_CXX_COMPILER=${CMAKE_CXX_COMPILER}
            -DCMAKE_INSTALL_PREFIX:PATH=<INSTALL_DIR>
            -DCMAKE_PREFIX_PATH:PATH=<INSTALL_DIR>
            -Wno-dev)

#=======================================================================================================================
# Download and build the zlib library
#=======================================================================================================================
ExternalProject_Add(zlib
        GIT_REPOSITORY https://github.com/madler/zlib.git
        GIT_TAG cacf7f1
        INSTALL_DIR ${OPENXLSX_DEPENDENCIES}
        UPDATE_COMMAND ""
        CMAKE_ARGS
            -DCMAKE_BUILD_TYPE=${CMAKE_BUILD_TYPE}
            -DCMAKE_C_COMPILER=${CMAKE_C_COMPILER}
            -DCMAKE_CXX_COMPILER=${CMAKE_CXX_COMPILER}
            -DCMAKE_POSITION_INDEPENDENT_CODE=ON
            -DCMAKE_INSTALL_PREFIX:PATH=<INSTALL_DIR>
            -DCMAKE_PREFIX_PATH:PATH=<INSTALL_DIR>
            -Wno-dev)

#=======================================================================================================================
# Download and build the bzip2 library
#=======================================================================================================================
ExternalProject_Add(bzip
        GIT_REPOSITORY https://github.com/osrf/bzip2_cmake.git
        GIT_TAG 0496eaa
        INSTALL_DIR ${OPENXLSX_DEPENDENCIES}
        UPDATE_COMMAND ""
        CMAKE_ARGS
            -DCMAKE_BUILD_TYPE=${CMAKE_BUILD_TYPE}
            -DCMAKE_C_COMPILER=${CMAKE_C_COMPILER}
            -DCMAKE_CXX_COMPILER=${CMAKE_CXX_COMPILER}
            -DCMAKE_POSITION_INDEPENDENT_CODE=ON
            -DCMAKE_INSTALL_PREFIX:PATH=<INSTALL_DIR>
            -DCMAKE_PREFIX_PATH:PATH=<INSTALL_DIR>
            -Wno-dev)

#=======================================================================================================================
# Download and build the libzip library
#=======================================================================================================================
ExternalProject_Add(libzip
        GIT_REPOSITORY https://github.com/nih-at/libzip.git
        GIT_TAG ff55682
        INSTALL_DIR ${OPENXLSX_DEPENDENCIES}
        UPDATE_COMMAND ""
        CMAKE_ARGS
            -DCMAKE_BUILD_TYPE=${CMAKE_BUILD_TYPE}
            -DCMAKE_C_COMPILER=${CMAKE_C_COMPILER}
            -DCMAKE_CXX_COMPILER=${CMAKE_CXX_COMPILER}
            -DENABLE_WINDOWS_CRYPTO=OFF
            -DCMAKE_POSITION_INDEPENDENT_CODE=ON
            -DBUILD_SHARED_LIBS=OFF
            -DCMAKE_INSTALL_PREFIX:PATH=<INSTALL_DIR>
            -DCMAKE_PREFIX_PATH:PATH=<INSTALL_DIR>
            -Wno-dev)

#=======================================================================================================================
# Build the libzip++ library
#=======================================================================================================================
ExternalProject_Add(libzip++
        SOURCE_DIR ${CMAKE_CURRENT_LIST_DIR}/@libzip++
        INSTALL_DIR ${OPENXLSX_DEPENDENCIES}
        UPDATE_COMMAND ""
        CMAKE_ARGS
            -DCMAKE_BUILD_TYPE=${CMAKE_BUILD_TYPE}
            -DCMAKE_C_COMPILER=${CMAKE_C_COMPILER}
            -DCMAKE_CXX_COMPILER=${CMAKE_CXX_COMPILER}
            -DCMAKE_CXX_STANDARD=${CMAKE_CXX_STANDARD}
            -DCMAKE_CXX_STANDARD_REQUIRED=${CMAKE_CXX_STANDARD_REQUIRED}
            -DCMAKE_CXX_EXTENSIONS=${CMAKE_CXX_EXTENSIONS}
            -DCMAKE_POSITION_INDEPENDENT_CODE=ON
            -DCMAKE_INSTALL_PREFIX:PATH=<INSTALL_DIR>
            -DCMAKE_PREFIX_PATH:PATH=<INSTALL_DIR>
            -Wno-dev)

#=======================================================================================================================
# Load the OpenXLSX library folder
#=======================================================================================================================
add_subdirectory(@openxlsx)

#=======================================================================================================================
# Define Dependencies
#=======================================================================================================================
ExternalProject_Add_StepDependencies(libzip install zlib bzip)
ExternalProject_Add_StepDependencies(libzip++ install libzip)
add_dependencies(OpenXLSX flat pugi libzip++)
