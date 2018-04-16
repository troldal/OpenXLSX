set(SOURCE_FILES main.cpp)

add_executable(Demo ${SOURCE_FILES})
target_include_directories(Demo PUBLIC "${LIBZIP_INCLUDE_DIR}")
target_link_libraries(Demo PUBLIC OpenXLSX ${LIBZIP_LIBRARY} ${Boost_LIBRARIES})