cmake_minimum_required(VERSION 3.5)
project(OpenXLSX.Xml)

add_library(xml OBJECT xml.h pugiconfig.h pugixml.h pugixml.cpp)