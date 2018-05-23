cmake_minimum_required(VERSION 3.5)
project (OpenXLSX.All)

set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_SOURCE_DIR}/CMake/Modules/")

#set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -flto")
IF(CMAKE_BUILD_TYPE STREQUAL "Debug")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wextra")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pedantic")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unused-value")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unused-parameter")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Weffc++")
ENDIF(CMAKE_BUILD_TYPE STREQUAL "Debug")
#set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pg")
#SET(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -pg")
#SET(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_SHARED_LINKER_FLAGS} -pg")

find_package(libzip REQUIRED)

add_subdirectory(@source)
add_subdirectory(@benchmark)
add_subdirectory(@demo)