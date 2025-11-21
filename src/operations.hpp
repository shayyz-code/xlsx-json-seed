#pragma once
#include <string>
#include <vector>
#include <xlnt/xlnt.hpp>

void split_column_xlsx(xlnt::worksheet &ws, const std::string &source, const std::string &delimiter, const std::vector<std::string> &targets);

void uppercase_column_xlsx(xlnt::worksheet &ws, const std::string &column);

void replace_in_column_xlsx(xlnt::worksheet &ws, const std::string &column, const std::string &find, const std::string &repl);

// helpers

int col_to_index(const std::string &col);
std::string to_upper(const std::string &str);