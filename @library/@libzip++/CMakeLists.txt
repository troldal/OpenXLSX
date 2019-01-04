#=======================================================================================================================
# Preamble
#=======================================================================================================================
cmake_minimum_required(VERSION 3.11 FATAL_ERROR)
project(libzip++ VERSION 0.5.0)

#=======================================================================================================================
# Create libzip++ target
#=======================================================================================================================
add_library(zip++ STATIC
        ${CMAKE_CURRENT_LIST_DIR}/libzip++.h
        ${CMAKE_CURRENT_LIST_DIR}/libzip++.cpp)

target_link_libraries(zip++ PUBLIC zip)
target_link_directories(zip++ PUBLIC ${CMAKE_INSTALL_PREFIX}/lib)
target_include_directories(zip++ PUBLIC ${CMAKE_INSTALL_PREFIX}/include)
#target_include_directories(OpenXLSX PUBLIC ${CMAKE_CURRENT_LIST_DIR}/@libzip++)

include(GNUInstallDirs)
install(TARGETS zip++
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
        LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
        RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR})
install(FILES ${CMAKE_CURRENT_LIST_DIR}/libzip++.h DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})
