#pragma once
#include <string>
#include <vector>
#include <optional>
#include "nitro_sheet.hpp"

void fill_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t header_row,        // 1-based Excel row
    const std::uint32_t first_data_row,    // 1-based first row of data
    const std::size_t col_index,           // 0-based column index in sheet.cols
    const std::string &fill_with,
    const std::string &new_header
);

void add_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t header_row,
    const std::uint32_t first_data_row,
    const std::string &at,          // "end", "beginning"/"start", or column letters
    const std::string &fill_with,
    const std::string &new_header
);

void remove_column_nitro(
    NitroSheet &sheet,
    const size_t col_index         // 0-based column index
);

void split_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t header_row,              // 1-based Excel header row
    const std::uint32_t first_data_row,          // 1-based first data row
    const std::size_t source_col_index,          // 0-based column index to split
    const char delimiter,                              // delimiter character
    const std::vector<size_t> &target_col_indices,  // 0-based target column indices
    const std::vector<std::string> &new_headers // optional headers for target columns
);

void uppercase_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t first_data_row,  // 1-based row index
    const std::size_t col_index          // 0-based column index
);

void replace_in_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t first_data_row,  // 1-based row index
    const std::size_t col_index,         // 0-based column index
    const std::string &find,
    const std::string &repl
);

void transform_row_nitro(
    NitroSheet &sheet,
    const std::size_t row_index,                // 0-based row index
    const std::string &to,           // "camelCase", "PascalCase", "snake_case", "upper", "lower"
    std::optional<char> delim = std::nullopt // optional delimiter for camel/pascal
);

void transform_header_nitro(
    NitroSheet &sheet,
    const std::string &to,           // "camelCase", "PascalCase", "snake_case", "upper", "lower"
    std::optional<char> delim = std::nullopt // optional delimiter for camel/pascal
);

void rename_header_nitro(
    NitroSheet &sheet,
    const std::size_t col_index,
    const std::string new_name
);

void group_collect_to_array_nitro(
    NitroSheet &sheet,
    std::size_t group_col,
    std::size_t collect_col,
    std::size_t output_col
);

void sort_rows_by_column_nitro(
    NitroSheet &sheet,
    const std::size_t col_index,
    const bool ascending
);

void reassign_numbering_nitro(
    NitroSheet &sheet,
    const std::size_t col_index,
    const std::string &prefix,
    const std::string &suffix,
    const size_t start_number,
    const size_t step
);