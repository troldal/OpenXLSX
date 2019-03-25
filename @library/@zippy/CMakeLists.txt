#=======================================================================================================================
# Define library targets
#=======================================================================================================================
add_library(Zippy INTERFACE)
add_library(Zippy::Zippy ALIAS Zippy)
target_include_directories(Zippy SYSTEM INTERFACE ${CMAKE_CURRENT_LIST_DIR})