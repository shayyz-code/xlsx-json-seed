#include <string>
#include <optional>

// helpers
std::string str_slice_from(const std::string &s, std::size_t start);

bool str_starts_with(const std::string &s, const std::string &prefix);

bool str_contains(const std::string &s, const std::string &subs);

bool str_contains_at_least_one_placeholder(const std::string &s);

std::size_t col_to_index(const std::string &letters);

std::string index_to_col(size_t index);

std::string to_clean_number(const std::string &s);

std::string to_upper(const std::string &str);

std::string to_lower(const std::string &str);

std::string to_camel(const std::string &str, char &delim);

std::string to_pascal(const std::string &str, char &delim);

std::string to_snake(const std::string &str);

// json-firestore-seed funcs
std::string random_past_utc_date_within_n_years(
    std::optional<uint32_t> n_years = 1
);

// os terminal helpers
int get_terminal_width();
void print_full_line_utf8(const std::string &color, const std::string &glyph);