#pragma once
#include <string>
#include <vector>
#include <optional>
#include <xlnt/xlnt.hpp>

void split_column_xlsx(xlnt::worksheet &ws, const std::string &source, const std::string &delimiter, const std::vector<std::string> &targets);

void uppercase_column_xlsx(xlnt::worksheet &ws, const std::string &column);

void replace_in_column_xlsx(xlnt::worksheet &ws, const std::string &column, const std::string &find, const std::string &repl);

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