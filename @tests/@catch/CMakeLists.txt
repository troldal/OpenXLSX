#=======================================================================================================================
# Define library targets
#=======================================================================================================================
add_library(Catch INTERFACE)
add_library(Catch::Catch ALIAS Catch)
target_include_directories(Catch SYSTEM INTERFACE ${CMAKE_CURRENT_LIST_DIR})