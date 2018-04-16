cmake_minimum_required(VERSION 3.5)
project (OpenXLSX)

set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_SOURCE_DIR}/CMake/Modules/")

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED TRUE)
IF(CMAKE_BUILD_TYPE STREQUAL "Debug")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -Weffc++ -pedantic -Wno-unused-value -Wno-unused-parameter")
ENDIF(CMAKE_BUILD_TYPE STREQUAL "Debug")
#set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pg")
#SET(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -pg")
#SET(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_SHARED_LINKER_FLAGS} -pg")

find_package(libzip REQUIRED)
find_package(Boost COMPONENTS system filesystem REQUIRED)
set(Boost_USE_STATIC_LIBS ON)

add_subdirectory(src)
add_subdirectory(demo)