#include <OpenXLSX.hpp>

#ifdef ENABLE_NOWIDE
    #include <nowide/iostream.hpp>
    using nowide::cout;
#else
    #include <iostream>
    using std::cout;
#endif

using namespace OpenXLSX;

int main()
{
    cout << "********************************************************************************\n";
    cout << "DEMO PROGRAM #04: Unicode\n";
    cout << "********************************************************************************\n";

    // Unicode can be a real pain in the neck. While UTF-8 encoding has become the de facto standard
    // encoding on Linux, macOS and the internet, some systems use other encodings, most notably
    // Windows which use UTF-16.
    // OpenXLSX is based on UTF-8. That means that all text input and output MUST be in UTF-8 format.
    // On Linux and MacOS, this will work out of the box because UTF-8 is baked into those systems.
    // On Windows, on the other hand, input and output must be converted to UTF-8 first. This includes
    // input strings that are hard-coded into your program. To stay sane, it is recommended that
    // all your source code files are in UTF-8 encoding. All major compilers and IDEs on windows
    // support UTF-8 encoded source files.
    // To convert input/output manually, you can use the Windows API, or 3rd party libraries, such
    // as Boost.NoWide.
    // In this example Boost.NoWide is used, as it provides some handy functionality that enables
    // console input/output with implicit conversion to/from UTF-8 on Windows.

    // First, create a new document and access the sheet named 'Sheet1'.
    // Then rename the worksheet to 'Простыня'.
    XLDocument doc1;
    doc1.create("./Demo04.xlsx", XLForceOverwrite);
    auto wks1 = doc1.workbook().worksheet("Sheet1");
    wks1.setName("Простыня");

    // Cell values can be set to any Unicode string using the normal value assignment methods.
    wks1.cell(XLCellReference("A1")).value() = "안녕하세요 세계!";
    wks1.cell(XLCellReference("A2")).value() = "你好，世界!";
    wks1.cell(XLCellReference("A3")).value() = "こんにちは 世界";
    wks1.cell(XLCellReference("A4")).value() = "नमस्ते दुनिया!";
    wks1.cell(XLCellReference("A5")).value() = "Привет, мир!";
    wks1.cell(XLCellReference("A6")).value() = "Γειά σου Κόσμε!";

    // Workbooks can also be saved and loaded with Unicode names
    doc1.save();
    doc1.saveAs("./スプレッドシート.xlsx", XLForceOverwrite);
    doc1.close();

    doc1.open("./スプレッドシート.xlsx");
    wks1 = doc1.workbook().worksheet("Простыня");

    // The nowide::cout object is a drop-in replacement of the std::cout that enables console output of UTF-8, even on Windows.
    cout << "Cell A1 (Korean)  : " << wks1.cell(XLCellReference("A1")).value().get<std::string>() << '\n';
    cout << "Cell A2 (Chinese) : " << wks1.cell(XLCellReference("A2")).value().get<std::string>() << '\n';
    cout << "Cell A3 (Japanese): " << wks1.cell(XLCellReference("A3")).value().get<std::string>() << '\n';
    cout << "Cell A4 (Hindi)   : " << wks1.cell(XLCellReference("A4")).value().get<std::string>() << '\n';
    cout << "Cell A5 (Russian) : " << wks1.cell(XLCellReference("A5")).value().get<std::string>() << '\n';
    cout << "Cell A6 (Greek)   : " << wks1.cell(XLCellReference("A6")).value().get<std::string>() << '\n';


    cout << "\nNOTE: If you are using a Windows terminal, the above output may look like gibberish,\n"
                    "because the Windows terminal does not support UTF-8 at the moment. To view to output,\n"
                    "you can use the overloaded 'cout' in the boost::nowide library (as in this sample program).\n"
                    "This will require a UTF-8 enabled font in the terminal. Lucinda Console supports some\n"
                    "non-ASCII scripts, such as Cyrillic and Greek. NSimSun supports some asian scripts.\n\n";

    doc1.close();

    return 0;
}
