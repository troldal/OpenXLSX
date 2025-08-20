#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.15 FATAL_ERROR)
project(OpenXLSX.Examples)
list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")

#=======================================================================================================================
# Include boost::nowide library (required for Demo1A and Demo4)
#=======================================================================================================================
#if (${OPENXLSX_ENABLE_NOWIDE})
#    if (NOT ${OPENXLSX_ENABLE_LIBBOOST_NOWIDE})
#        # boost nowide is header-only, a "library" only makes sense in the context of the nowide header files shipped with OpenXLSX
#        add_subdirectory(external/nowide EXCLUDE_FROM_ALL) # TBD if this line is needed at all
#        add_library(NoWide INTERFACE IMPORTED)
#        target_include_directories(NoWide SYSTEM INTERFACE $<BUILD_INTERFACE:${CMAKE_SOURCE_DIR}/OpenXLSX/external/nowide/>)
#    endif()
#endif()

#=======================================================================================================================
# Define Demo1 target
#=======================================================================================================================
add_executable(Demo1 Demo1.cpp)
target_link_libraries(Demo1 PRIVATE OpenXLSX::OpenXLSX)

#=======================================================================================================================
# Define Demo1A target
#=======================================================================================================================
if(${OPENXLSX_ENABLE_LIBZIP_EXAMPLE})
    find_package(LIBZIP MODULE REQUIRED)
    add_executable(Demo1A Demo1A.cpp)
    target_link_libraries(Demo1A PRIVATE OpenXLSX::OpenXLSX libzip::zip)
    if (${OPENXLSX_ENABLE_NOWIDE})
#        target_compile_definitions(Demo1A PRIVATE ENABLE_NOWIDE)
        target_link_libraries(Demo1A PRIVATE ${NOWIDE_TARGET})
        # target_link_libraries(Demo4 PRIVATE nowide::nowide)
#        if (${OPENXLSX_ENABLE_LIBBOOST_NOWIDE})
#            # when libboost-nowide is used, a preprocessor directive is needed to change include path and namespace
#            target_compile_definitions(Demo1A PRIVATE ENABLE_LIBBOOST_NOWIDE)
#        else()
#            target_link_libraries(Demo1A PRIVATE $<BUILD_INTERFACE:NoWide>)
#        endif()
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
    target_link_libraries(Demo4 PRIVATE ${NOWIDE_TARGET})
    # target_link_libraries(Demo4 PRIVATE nowide::nowide)
#    if (${OPENXLSX_ENABLE_LIBBOOST_NOWIDE})
#        # when libboost-nowide is used, a preprocessor directive is needed to change include path and namespace
#        target_compile_definitions(Demo4 PRIVATE ENABLE_LIBBOOST_NOWIDE)
#    else()
#        target_link_libraries(Demo1A PRIVATE $<BUILD_INTERFACE:NoWide>)
#
#    endif()
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

#=======================================================================================================================
# Define Demo10 target
#=======================================================================================================================
add_executable(Demo10 Demo10.cpp)
target_link_libraries(Demo10 PRIVATE OpenXLSX::OpenXLSX)
