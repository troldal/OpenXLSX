#!/bin/bash

# save original internal field separator & set it to newline
DEFAULT_IFS=$IFS
IFS=$'\n'

if [ "$1" = "doit" ]; then
	# delete individual known files, suppressing error output (in case they do not exist)
	file="CMakeCache.txt"
	recurse=""
	if [ -e "$file" ]; then
		rm $recurse "$file"
	fi
	file="cmake-log"
	recurse=""
	if [ -e "$file" ]; then
		rm $recurse "$file"
	fi
	file="OpenXLSX/OpenXLSX/"
	recurse="-r"
	if [ -e "$file" ]; then
		rm $recurse "$file"
	fi

	# use find command to locate files and folders that occur multiple times & then loop to delete them
	CMakeFiles=`find . -name "CMakeFiles" -exec echo {} \;`
	for file in $CMakeFiles; do
		rm -r "$file"
	done
   cmake_install_cmake=`find . -name "cmake_install.cmake" -exec echo {} \;`
	for file in $cmake_install_cmake; do
		rm -r "$file"
	done
	Makefiles=`find . -name "Makefile" -exec echo {} \;`
	for file in $Makefiles; do
		rm -r "$file"
	done
else
	# echo commands for deleting individual known files
	file="CMakeCache.txt"
	recurse=""
	if [ -e "$file" ]; then
		echo "rm $recurse \"$file\""
	fi
	file="cmake-log"
	recurse=""
	if [ -e "$file" ]; then
		echo "rm $recurse \"$file\""
	fi
	file="OpenXLSX/OpenXLSX/"
	recurse="-r"
	if [ -e "$file" ]; then
		echo "rm $recurse \"$file\""
	fi

	# use find command to locate files and folders that occur multiple times & then echo commands for deleting them
	CMakeFiles=`find . -name "CMakeFiles" -exec echo {} \;`
	for file in $CMakeFiles; do
		echo "rm -r \"$file\""
	done
   cmake_install_cmake=`find . -name "cmake_install.cmake" -exec echo {} \;`
	for file in $cmake_install_cmake; do
		echo "rm -r \"$file\""
	done
	Makefiles=`find . -name "Makefile" -exec echo {} \;`
	for file in $Makefiles; do
		echo "rm -r \"$file\""
	done
fi

# restore original internal field separator
IFS=$DEFAULT_IFS
