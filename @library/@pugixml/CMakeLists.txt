#=======================================================================================================================
# Define library targets
#=======================================================================================================================
add_library(PugiXML INTERFACE)
add_library(PugiXML::PugiXML ALIAS PugiXML)
target_include_directories(PugiXML SYSTEM INTERFACE ${CMAKE_CURRENT_LIST_DIR})