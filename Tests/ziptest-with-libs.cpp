/*
 * Program:  ziptest-with-libs
 * Compile instructions: g++ -DUSE_LIBZIP -IOpenXLSX -IOpenXLSX/headers ziptest-with-libs.cpp OpenXLSX/sources/XLZipArchive.cpp -lzip -o ziptest-with-libs
 * Part of: OpenXLSX
 *
 * OpenXLSX Copyright (c) 2018, Kenneth Troldal Balslev
 * ziptest Copyright 2025 Lars Uffmann <coding23@uffmann.name>
 *
 * --- NOTE: The below license is not applicable to namespace OpenXLSX! ---
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
/*!
 * @program ziptest-with-libs
 * @depends libzip
 * @version 2025-04-27 19:00 CEST
 * @brief   provide a wrapper API for libzip to interface with OpenXLSX XLZipArchive
 * @disclaimer This module is in a draft stage, to be reviewed / tested
 */
#include <cstdio>       // fprintf
#include <filesystem>   // std::filesystem::remove, std::filesystem::rename
#include <iostream>     // std::cout / std::cerr
#include <memory>       // std::shared_ptr
#include <random>       // std::random_device, std::mt19937, std::uniform_int_distribution
#include <string>       // std::string
#include <string.h>     // stderror
#include <sys/stat.h>   // stat, struct stat

#include <XLZipArchive.hpp>

int main(int argc, char *argv[])
{
	std::string archiveName = "Demo01";
	std::string archiveExt = ".xlsx";

	OpenXLSX::XLZipArchive myArc{};

	std::cout << "myArc.usesZippy() is " << ( myArc.usesZippy() ? "true" : "false" ) << std::endl;
	std::cout << "myArc.isOpen() is " << ( myArc.isOpen() ? "true" : "false" ) << std::endl;
	myArc.open(archiveName + archiveExt);
	std::cout << "myArc.isOpen() is " << ( myArc.isOpen() ? "true" : "false" ) << std::endl;
	std::string entryName = "docProps/core.xml";
	std::cout << "myArc.hasEntry(\"" << entryName << "\") is " << ( myArc.hasEntry(entryName) ? "true" : "false" ) << std::endl;

	int i = 0;
	ssize_t entryCount = myArc.entryCount();
	std::cout << "myArc.entryCount() is " << entryCount << std::endl;
	while (i < entryCount) {
		std::cout << "myArc.entryName(" << i << ") is " << myArc.entryName(i) << std::endl;
		++i;
	}

	std::string testEntry = "xl/worksheets/sheet2.xml";
	std::cout << "adding a file to myArc: " << testEntry << std::endl;
	myArc.addEntry(testEntry, "Hello World!");

	myArc.commitChanges();

	// testEntry = "xl/worksheets/sheet1.xml";
	std::string test = myArc.getEntry(testEntry);
	std::cout << "contents of " << testEntry << ":" << std::endl;
	std::cout << "------------------------------------------------" << std::endl;
	std::cout << test << std::endl;
	std::cout << "------------------------------------------------" << std::endl;

// TBD: for some reason, replacing any file except the ones that this program added itself, fixes the central general_purpose_bit_flags error that unzip reports
// entryName = "_rels/.rels";
// entryName = testEntry;
// std::string eData = myArc.getEntry(entryName);
// myArc.addEntry(entryName, eData); // replace archive contents
// myArc.commitChanges();

	myArc.save(archiveName + "_1" + archiveExt); // saves and re-opens from source
	myArc.close();
	std::cout << "myArc.isOpen() is " << ( myArc.isOpen() ? "true" : "false" ) << std::endl;

	return 0;
}
