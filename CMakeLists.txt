cmake_minimum_required(VERSION 3.5)
project (OpenXLSX.All)



#set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pg")
#SET(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -pg")
#SET(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_SHARED_LINKER_FLAGS} -pg")

#find_package(libzip REQUIRED)

add_subdirectory(@library)
add_subdirectory(@benchmark)
add_subdirectory(@demo)