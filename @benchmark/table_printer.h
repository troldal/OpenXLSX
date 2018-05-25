// Adapted from bprinter by (https://github.com/dattanchu/bprinter)


#ifndef BPRINTER_TABLE_PRINTER_H_
#define BPRINTER_TABLE_PRINTER_H_

#include <iostream>
#include <iomanip>
#include <vector>
#include <string>
#include <sstream>
#include <cmath>
#include <stdexcept>

namespace bprinter {
    class endl{};

    /**
     * @class TablePrinter
     *
     * Print a pretty table into your output of choice.
     *
     * Usage:
     *   TablePrinter tp(&std::cout);
     *   tp.AddColumn("Name", 25);
     *   tp.AddColumn("Age", 3);
     *   tp.AddColumn("Position", 30);
     *
     *   tp.PrintHeader();
     *   tp << "Dat Chu" << 25 << "Research Assistant";
     *   tp << "John Doe" << 26 << "Professional Anonymity";
     *   tp << "Jane Doe" << tp.SkipToNextLine();
     *   tp << "Tom Doe" << 7 << "Student";
     *   tp.PrintFooter();
     *
     * @todo Add support for padding in each table cell
     **/
    class TablePrinter{
    public:
        TablePrinter(std::ostream * output, const std::string & separator = "|");
        ~TablePrinter();

        int get_num_columns() const;
        int get_table_width() const;
        void set_separator(const std::string & separator);
        void set_flush_left();
        void set_flush_right();

        void AddColumn(const std::string & header_name, int column_width);
        void PrintTitle(const std::string &title);
        void PrintHeader();
        void PrintFooter();

        TablePrinter& operator<<(endl input){
            while (j_ != 0){
                *this << "";
            }
            return *this;
        }

        template<typename T>
        TablePrinter& operator<<(T input){
            if constexpr(std::is_floating_point<T>::value) {
                OutputDecimalNumber<T>(input);
            }
            else {

                if (j_ == 0)
                    *out_stream_ << "|";

                if (flush_left_)
                    *out_stream_ << std::left;
                else
                    *out_stream_ << std::right;

                // Leave 3 extra space: One for negative sign, one for zero, one for decimal
                *out_stream_ << std::setw(column_widths_.at(j_))
                             << input;

                if (j_ == get_num_columns() - 1) {
                    *out_stream_ << "|\n";
                    i_ = i_ + 1;
                    j_ = 0;
                }
                else {
                    *out_stream_ << separator_;
                    j_ = j_ + 1;
                }
            }
            return *this;
        }

    private:
        void PrintHorizontalLine(char character = '-');

        template<typename T>
        void OutputDecimalNumber(T input);

        std::ostream * out_stream_;
        std::vector<std::string> column_headers_;
        std::vector<int> column_widths_;
        std::string separator_;

        int i_; // index of current row
        int j_; // index of current column

        int table_width_;
        bool flush_left_;
    };

    template<typename T>
    void TablePrinter::OutputDecimalNumber(T input){
        // If we cannot handle this number, indicate so
        if (input < 10*(column_widths_.at(j_)-1) || input > 10*column_widths_.at(j_)){
            std::stringstream string_out;
            string_out << std::setiosflags(std::ios::fixed)
                       << std::setprecision(column_widths_.at(j_))
                       << std::setw(column_widths_.at(j_))
                       << input;

            std::string string_rep_of_number = string_out.str();

            string_rep_of_number[column_widths_.at(j_)-1] = '*';
            std::string string_to_print = string_rep_of_number.substr(0, column_widths_.at(j_));
            *out_stream_ << string_to_print;
        } else {

            // determine what precision we need
            int precision = column_widths_.at(j_) - 1; // leave room for the decimal point
            if (input < 0)
                --precision; // leave room for the minus sign

            // leave room for digits before the decimal?
            if (input < -1 || input > 1){
                int num_digits_before_decimal = 1 + (int)log10(std::abs(input));
                precision -= num_digits_before_decimal;
            }
            else
                precision --; // e.g. 0.12345 or -0.1234

            if (precision < 0)
                precision = 0; // don't go negative with precision

            *out_stream_ << std::setiosflags(std::ios::fixed)
                         << std::setprecision(precision)
                         << std::setw(column_widths_.at(j_))
                         << input;
        }

        if (j_ == get_num_columns()-1){
            *out_stream_ << "|\n";
            i_ = i_ + 1;
            j_ = 0;
        } else {
            *out_stream_ << separator_;
            j_ = j_ + 1;
        }
    }

    TablePrinter::TablePrinter(std::ostream * output, const std::string & separator){
        out_stream_ = output;
        i_ = 0;
        j_ = 0;
        separator_ = separator;
        table_width_ = 0;
        flush_left_ = false;
    }

    TablePrinter::~TablePrinter(){

    }

    int TablePrinter::get_num_columns() const {
        return column_headers_.size();
    }

    int TablePrinter::get_table_width() const {
        return table_width_;
    }

    void TablePrinter::set_separator(const std::string &separator){
        separator_ = separator;
    }

    void TablePrinter::set_flush_left(){
        flush_left_ = true;
    }

    void TablePrinter::set_flush_right(){
        flush_left_ = false;
    }

    /**
     * @brief Add a column to our table
     * @param header_name Name to be print for the header
     * @param column_width the width of the column (has to be >=5)
     **/
    void TablePrinter::AddColumn(const std::string & header_name, int column_width){
        if (column_width < 4){
            throw std::invalid_argument("Column size has to be >= 4");
        }

        column_headers_.push_back(header_name);
        column_widths_.push_back(column_width);
        table_width_ += column_width + separator_.size(); // for the separator
    }

    void TablePrinter::PrintHorizontalLine(char character)
    {
        *out_stream_ << "+"; // the left bar

        for (int i=0; i<table_width_-1; ++i)
            *out_stream_ << character;

        *out_stream_ << "+"; // the right bar
        *out_stream_ << "\n";
    }

    void TablePrinter::PrintTitle(const std::string &title)
    {
        auto totalWidth = 0;
        for (auto &it : column_widths_) totalWidth += it;
        totalWidth += column_widths_.size() - 1;

        auto tit = title;
        auto titleWidth = tit.length();
        if (titleWidth>totalWidth) tit = tit.substr(0, totalWidth);

        auto pre = (totalWidth - tit.length()) / 2;
        auto post = (totalWidth - tit.length() - pre);

        PrintHorizontalLine('=');
        *out_stream_ << "|";
        for (int i = 0; i < pre; ++i) *out_stream_ << " ";
        *out_stream_ << tit;
        for (int i = 0; i < post; ++i) *out_stream_ << " ";
        *out_stream_ << "|\n";

    }

    void TablePrinter::PrintHeader(){
        PrintHorizontalLine('=');
        *out_stream_ << "|";

        for (int i=0; i<get_num_columns(); ++i){

            if(flush_left_)
                *out_stream_ << std::left;
            else
                *out_stream_ << std::right;

            *out_stream_ << std::setw(column_widths_.at(i)) << column_headers_.at(i).substr(0, column_widths_.at(i));
            if (i != get_num_columns()-1){
                *out_stream_ << separator_;
            }
        }

        *out_stream_ << "|\n";
        PrintHorizontalLine('=');
    }

    void TablePrinter::PrintFooter(){
        PrintHorizontalLine();
    }
}

#endif //BPRINTER_TABLE_PRINTER_H_
