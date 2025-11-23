#include <string>
#include <optional>

// helpers
std::string str_slice_from(const std::string &s, size_t start);

bool str_starts_with(const std::string &s, const std::string &prefix);

int col_to_index(const std::string &col);

std::string to_upper(const std::string &str);

std::string to_lower(const std::string &str);

std::string to_camel(const std::string &str, char &delim);

std::string to_pascal(const std::string &str, char &delim);

std::string to_snake(const std::string &str);

// json-firestore-seed funcs

std::string random_past_utc_date_within_n_years(
    std::optional<uint32_t> n_years = 1
);