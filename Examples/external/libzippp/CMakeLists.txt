cmake_minimum_required(VERSION 3.12.0)

project(libzippp VERSION 3.0.0)

# Use C++11
set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# Do not install or build tests by default when this file is included
# as a sub-project (via add_subdirectory)
if(CMAKE_SOURCE_DIR STREQUAL CMAKE_CURRENT_SOURCE_DIR)
  set(is_root_project ON)
else()
  set(is_root_project OFF)
endif()

option(LIBZIPPP_INSTALL "Install library" ${is_root_project})
option(LIBZIPPP_INSTALL_HEADERS "Install the headers" ${is_root_project})
option(LIBZIPPP_BUILD_TESTS "Build unit tests" ${is_root_project})
option(LIBZIPPP_ENABLE_ENCRYPTION "Build with encryption enabled" OFF)

list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")

find_package(LIBZIP MODULE REQUIRED)

add_library(libzippp "src/libzippp.cpp")
add_library(libzippp::libzippp ALIAS libzippp) # Convenience alias
target_include_directories(libzippp 
  PUBLIC  
    $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/src>
    $<INSTALL_INTERFACE:include/libzippp>
)
set_target_properties(libzippp PROPERTIES PREFIX "") # Avoid duplicate "lib" prefix
target_link_libraries(libzippp PRIVATE libzip::zip)

if(LIBZIPPP_ENABLE_ENCRYPTION)
  target_compile_definitions(libzippp PUBLIC LIBZIPPP_WITH_ENCRYPTION)
endif()

if (BUILD_SHARED_LIBS)
  target_compile_definitions(libzippp PRIVATE LIBZIPPP_EXPORTS)
else()
  set_target_properties(libzippp PROPERTIES OUTPUT_NAME "libzippp_static")
endif()

export(TARGETS libzippp NAMESPACE libzippp:: FILE ${PROJECT_BINARY_DIR}/libzippp-target.cmake)

if(LIBZIPPP_BUILD_TESTS)
  enable_testing()
  add_executable(libzippp_test "tests/tests.cpp")
  target_link_libraries(libzippp_test PRIVATE libzippp)
  add_test(NAME libzippp_tests COMMAND libzippp_test)
  
  if (BUILD_SHARED_LIBS)
    set_target_properties(libzippp_test PROPERTIES OUTPUT_NAME "libzippp_shared_test")
  else()
    set_target_properties(libzippp_test PROPERTIES OUTPUT_NAME "libzippp_static_test")
  endif()
endif()

if(LIBZIPPP_INSTALL)
  install(
    TARGETS libzippp 
    EXPORT libzipppTargets
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib 
    RUNTIME DESTINATION bin 
  )

  if(LIBZIPPP_INSTALL_HEADERS)
    install(FILES src/libzippp.h DESTINATION include/libzippp)
  endif()

  include(CMakePackageConfigHelpers)

  set(PROJECT_CONFIG_FILE "${CMAKE_CURRENT_BINARY_DIR}/generated/${PROJECT_NAME}Config.cmake")
  set(PROJECT_VERSION_FILE "${CMAKE_CURRENT_BINARY_DIR}/generated/${PROJECT_NAME}ConfigVersion.cmake")
  if(WIN32)
    set(configInstallDestination cmake/${PROJECT_NAME})
  else()
    set(configInstallDestination share/${PROJECT_NAME})
  endif()
  
  configure_package_config_file(
    "Config.cmake.in" 
    ${PROJECT_CONFIG_FILE} 
    INSTALL_DESTINATION ${configInstallDestination}
  )
  write_basic_package_version_file(
    ${PROJECT_VERSION_FILE}
    COMPATIBILITY SameMajorVersion
  )

  install(
    FILES ${PROJECT_CONFIG_FILE} ${PROJECT_VERSION_FILE} cmake/FindLIBZIP.cmake
    DESTINATION ${configInstallDestination}
  )

  install(
    EXPORT libzipppTargets
    NAMESPACE "${PROJECT_NAME}::"
    DESTINATION ${configInstallDestination}
  )
endif()

