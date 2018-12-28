include(FetchContent)

#=======================================================================================================================
# Fetch pugixml
#=======================================================================================================================
FetchContent_Declare(
        pugi
        GIT_REPOSITORY https://github.com/zeux/pugixml.git
        GIT_TAG 8436f2a
        GIT_PROGRESS TRUE)

FetchContent_GetProperties(pugi)
if(NOT pugi_POPULATED)
    FetchContent_Populate(pugi)
    add_subdirectory(
            ${pugi_SOURCE_DIR}
            ${pugi_BINARY_DIR})
endif()

#=======================================================================================================================
# Fetch zlib
#=======================================================================================================================
FetchContent_Declare(
        zlib
        GIT_REPOSITORY https://github.com/madler/zlib.git
        GIT_TAG cacf7f1
        GIT_PROGRESS TRUE)

FetchContent_GetProperties(zlib)
if(NOT zlib_POPULATED)
    FetchContent_Populate(zlib)
    add_subdirectory(
            ${zlib_SOURCE_DIR}
            ${zlib_BINARY_DIR})
endif()

#=======================================================================================================================
# Fetch libzip
#=======================================================================================================================
FetchContent_Declare(
        libzip
        GIT_REPOSITORY https://github.com/nih-at/libzip.git
        GIT_TAG ff55682
        GIT_PROGRESS TRUE
        CMAKE_ARGS -DBUILD_SHARED_LIBS=OFF
                   -DZLIB_INCLUDE_DIR=${zlib_SOURCE_DIR})

FetchContent_GetProperties(libzip)
if(NOT libzip_POPULATED)
    FetchContent_Populate(libzip)
    add_subdirectory(
            ${libzip_SOURCE_DIR}
            ${libzip_BINARY_DIR})
endif()

#=======================================================================================================================
# Create libzip++ target
#=======================================================================================================================
add_library(zip++ STATIC libzip++.h libzip++.cpp)
target_link_libraries(zip++ PUBLIC zip)
target_include_directories(zip++ PUBLIC ${libzip_SOURCE_DIR}/lib ${libzip_BINARY_DIR})

#=======================================================================================================================
# Append include paths to target
#=======================================================================================================================
target_include_directories(OpenXLSX PUBLIC ${pugi_SOURCE_DIR} ${CMAKE_CURRENT_LIST_DIR} ${CMAKE_CURRENT_LIST_DIR}/@flat)