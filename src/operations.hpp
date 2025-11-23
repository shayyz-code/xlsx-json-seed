#pragma once
#include <string>
#include <vector>
#include <optional>
#include <xlnt/xlnt.hpp>

void fill_column_xlsx(
    xlnt::worksheet &ws,
    const std::uint32_t header_row,
    const std::uint32_t first_data_row,
    const std::string &column,
    const std::string &fill_with,
    const std::string &new_header
);

void split_column_xlsx(
    xlnt::worksheet &ws,
    const std::uint32_t header_row,
    const std::uint32_t first_data_row,
    const std::string &source,
    const std::string &delimiter,
    const std::vector<std::string> &targets,
    const std::vector<std::string> &new_headers
);

void uppercase_column_xlsx(xlnt::worksheet &ws, const std::uint32_t first_data_row, const std::string &column);

void replace_in_column_xlsx(xlnt::worksheet &ws, const std::uint32_t first_data_row, const std::string &column, const std::string &find, const std::string &repl);

void transform_row_xlsx(
    xlnt::worksheet &ws,
    const std::uint32_t row,
    const std::string &to,
    std::optional<char> delim = std::nullopt
);

// helpers

int col_to_index(const std::string &col);
std::string to_upper(const std::string &str);
std::string to_lower(const std::string &str);

std::string to_camel(const std::string &str, char &delim);
std::string to_pascal(const std::string &str, char &delim);
std::string to_snake(const std::string &str);