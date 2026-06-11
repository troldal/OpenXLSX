#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Examples)
list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")


#=======================================================================================================================
# Define Demo1 target
#=======================================================================================================================
add_executable(Demo1 Demo1.cpp)
# NOTE: Linking against OpenXLSX-bundled does not yet work
# if(${OPENXLSX_BUNDLED_STATIC_LIB})
#     # target_include_directories(Demo1
#     #             PUBLIC
#     #             $<BUILD_INTERFACE:${CMAKE_SOURCE_DIR}/OpenXLSX>
#     #             $<BUILD_INTERFACE:${CMAKE_SOURCE_DIR}/OpenXLSX/headers>
#     #             # ${Boost_INCLUDE_DIR}    # just in case
#     #             # ${NOWIDE_INCLUDE_DIR}   # includes {Boost_INCLUDE_DIR}/boost
#     #             $<BUILD_INTERFACE:${CMAKE_BINARY_DIR}/OpenXLSX>)     # For export header
#     target_link_libraries(Demo1 PRIVATE OpenXLSX-bundled)
# else()
    target_link_libraries(Demo1 PRIVATE OpenXLSX::OpenXLSX)
# endif()


#=======================================================================================================================
# Define Demo1A target
#=======================================================================================================================
if(${OPENXLSX_ENABLE_LIBZIP_EXAMPLE})
    find_package(LIBZIP MODULE REQUIRED)
    add_executable(Demo1A Demo1A.cpp)
    target_link_libraries(Demo1A PRIVATE OpenXLSX::OpenXLSX libzip::zip)
    if (${OPENXLSX_ENABLE_NOWIDE})
        target_compile_definitions(Demo1A PRIVATE ENABLE_NOWIDE)
        target_include_directories(Demo1A PRIVATE ${EXAMPLES_NOWIDE_INCLUDE})
        target_link_libraries(Demo1A PRIVATE ${NOWIDE_ALIAS})
    endif()
endif()

#=======================================================================================================================
# Define Demo2 target
#=======================================================================================================================
add_executable(Demo2 Demo2.cpp)
target_link_libraries(Demo2 PRIVATE OpenXLSX::OpenXLSX)

#=======================================================================================================================
# Define Demo3 target
#=======================================================================================================================
add_executable(Demo3 Demo3.cpp)
target_link_libraries(Demo3 PRIVATE OpenXLSX::OpenXLSX)

#=======================================================================================================================
# Define Demo4 target
#=======================================================================================================================
add_executable(Demo4 Demo4.cpp)
target_link_libraries(Demo4 PRIVATE OpenXLSX::OpenXLSX)
if (${OPENXLSX_ENABLE_NOWIDE})
    target_compile_definitions(Demo4 PRIVATE ENABLE_NOWIDE)
    target_include_directories(Demo4 PRIVATE ${EXAMPLES_NOWIDE_INCLUDE})
    target_link_libraries(Demo4 PRIVATE ${NOWIDE_ALIAS})
endif()

#=======================================================================================================================
# Define Demo9 target
#=======================================================================================================================
add_executable(Demo5 Demo5.cpp)
target_link_libraries(Demo5 PRIVATE OpenXLSX::OpenXLSX)

#=======================================================================================================================
# Define Demo6 target
#=======================================================================================================================
add_executable(Demo6 Demo6.cpp)
target_link_libraries(Demo6 PRIVATE OpenXLSX::OpenXLSX)

#=======================================================================================================================
# Define Demo7 target
#=======================================================================================================================
add_executable(Demo7 Demo7.cpp)
target_link_libraries(Demo7 PRIVATE OpenXLSX::OpenXLSX)

#=======================================================================================================================
# Define Demo8 target
#=======================================================================================================================
add_executable(Demo8 Demo8.cpp)
target_link_libraries(Demo8 PRIVATE OpenXLSX::OpenXLSX)

#=======================================================================================================================
# Define Demo9 target
#=======================================================================================================================
add_executable(Demo9 Demo9.cpp)
target_link_libraries(Demo9 PRIVATE OpenXLSX::OpenXLSX)
if (${OPENXLSX_ENABLE_NOWIDE})
    target_compile_definitions(Demo9 PRIVATE ENABLE_NOWIDE)
    target_link_libraries(Demo9 PRIVATE ${NOWIDE_ALIAS})
endif()

#=======================================================================================================================
# Define Demo10 target
#=======================================================================================================================
add_executable(Demo10 Demo10.cpp)
target_link_libraries(Demo10 PRIVATE OpenXLSX::OpenXLSX)
