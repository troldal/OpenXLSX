cmake_minimum_required(VERSION 3.5)
project(OpenXLSX.Zip VERSION 0.5)

add_library(xml STATIC xml.h pugiconfig.h pugixml.h pugixml.cpp)