#include <algorithm>
#include <sstream>
#include "operations.hpp"
#include "utils/dynamic_placeholder.hpp"

// ops

// ----------------------
// Fill a NitroSheet column
// ----------------------
void fill_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t header_row,        // 1-based Excel row
    const std::uint32_t first_data_row,    // 1-based first row of data
    const size_t col_index,           // 0-based column index in sheet.cols
    const std::string &fill_with,
    const std::string &new_header = ""
)
{
    if (col_index >= sheet.cols.size())
        sheet.cols.resize(col_index + 1);

    const size_t total_rows = sheet.num_rows;
    if (total_rows == 0) return;

    size_t hdr = header_row - 1;
    size_t data_start = (first_data_row > 0 ? first_data_row - 1 : 1);

    if (data_start >= total_rows) return;

    Column &col = sheet.cols[col_index];
    // Ensure column has enough rows
    if (col.vals.size() < sheet.num_rows)
        col.vals.resize(sheet.num_rows);

    const std::string prefix = "firestore-random-past-date-n-year-";

    for (size_t r = 0; r < total_rows; ++r)
    {
        if (fill_with == "firestore-now") // now
        {
            col.vals[r] = "__fire_ts_now__";
        }
        else if (fill_with.compare(0, prefix.size(), prefix) == 0) // random past date
        {
            std::string years_part = fill_with.substr(prefix.size());
            std::optional<uint32_t> n_years;

            try {
                n_years = std::stoul(years_part);
            }
            catch (...) {
                std::cerr << "WARNING: Could not parse N years: " << fill_with << "\n";
            }

            std::string ts = random_past_utc_date_within_n_years(n_years);
            col.vals[r] = "{ \"__fire_ts_from_date__\": \"" + ts + "\" }";
        }
        else if (str_contains_at_least_one_placeholder(fill_with))
        {
            auto placeholders = scan_placeholders(fill_with);

            // start with the base string
            col.vals[r] = fill_with;

            for (auto &p : placeholders)
            {
                std::string replacement;

                if (p.key.rfind("col ", 0) == 0) // key starts with "col "
                {
                    // simple column reference
                    std::string col_letters = p.key.substr(4);
                    size_t ref_col_index = col_to_index(col_letters);

                    if (ref_col_index < sheet.cols.size() && r < sheet.cols[ref_col_index].vals.size())
                        replacement = sheet.cols[ref_col_index].vals[r];
                }
                else if (p.key.rfind("ifcol ", 0) == 0)
                {
                    std::string expr = p.key.substr(6); // remove "ifcol "

                    auto trim = [](std::string s) -> std::string {
                        size_t start = s.find_first_not_of(" \t");
                        size_t end = s.find_last_not_of(" \t");
                        if (start == std::string::npos) return "";
                        return s.substr(start, end - start + 1);
                    };

                    auto extract_quoted = [&](const std::string &s) -> std::string {
                        std::string t = trim(s);
                        if ((t.front() == '\'' && t.back() == '\'') || (t.front() == '"' && t.back() == '"'))
                            return t.substr(1, t.size() - 2);
                        return t;
                    };

                    // find ? and :
                    size_t qmark_pos = expr.find("?");
                    size_t colon_pos = expr.find(":");
                    if (qmark_pos == std::string::npos || colon_pos == std::string::npos) continue;

                    std::string cond_part = trim(expr.substr(0, qmark_pos));
                    std::string true_result  = extract_quoted(expr.substr(qmark_pos + 1, colon_pos - (qmark_pos + 1)));
                    std::string false_result = extract_quoted(expr.substr(colon_pos + 1));

                    // parse condition: <col_letters> <op> <value>
                    std::string col_letters, op, value;
                    {
                        std::istringstream iss(cond_part);
                        if (!(iss >> col_letters >> op)) continue;
                        std::getline(iss, value);
                        value = trim(value);
                        value = extract_quoted(value);
                    }

                    size_t ref_col_index = col_to_index(col_letters);
                    if (ref_col_index >= sheet.cols.size() || r >= sheet.cols[ref_col_index].vals.size()) continue;
                    const std::string &cell_val = sheet.cols[ref_col_index].vals[r];

                    // comparison
                    bool condition = false;

                    auto is_number = [](const std::string &s) -> bool {
                        if (s.empty()) return false;
                        char* endptr = nullptr;
                        strtod(s.c_str(), &endptr);
                        return (*endptr == 0);
                    };

                    // numeric comparison if both sides are numbers
                    if (is_number(cell_val) && is_number(value))
                    {
                        double lhs = std::stod(cell_val);
                        double rhs = std::stod(value);

                        if (op == "==") condition = lhs == rhs;
                        else if (op == "!=") condition = lhs != rhs;
                        else if (op == ">")  condition = lhs > rhs;
                        else if (op == "<")  condition = lhs < rhs;
                        else if (op == ">=") condition = lhs >= rhs;
                        else if (op == "<=") condition = lhs <= rhs;
                    }
                    else
                    {
                        // string comparison only supports == and !=
                        if (op == "==") condition = cell_val == value;
                        else if (op == "!=") condition = cell_val != value;
                    }

                    // resolve true_result / false_result if they reference columns
                    auto resolve_result = [&](const std::string &res) -> std::string {
                        std::string s = trim(res);
                        if (s.rfind("col ", 0) == 0)
                        {
                            size_t tcol_index = col_to_index(s.substr(4));
                            if (tcol_index < sheet.cols.size() && r < sheet.cols[tcol_index].vals.size())
                                return sheet.cols[tcol_index].vals[r];
                            else
                                return "";
                        }
                        return s;
                    };

                    replacement = condition ? resolve_result(true_result) : resolve_result(false_result);
                }



                // replace placeholder in current cell
                std::string full = "${" + p.key + "}";
                size_t pos = 0;
                while ((pos = col.vals[r].find(full, pos)) != std::string::npos)
                {
                    col.vals[r].replace(pos, full.size(), replacement);
                    pos += replacement.size();
                }
            }

        }
        else
        {
            col.vals[r] = fill_with;
        }
    }

    
    // Update header
    if (!new_header.empty() && hdr < col.vals.size())
    {
        col.header = new_header;
    }
}

// ----------------------
// Add a new column to NitroSheet
// ----------------------
void add_column_nitro(
    NitroSheet &sheet,
    const uint32_t header_row,
    const uint32_t first_data_row,
    const std::string &at,          // "end", "beginning"/"start", or column letters
    const std::string &fill_with,
    const std::string &new_header
)
{
    const size_t total_rows = sheet.num_rows;
    if (total_rows == 0)
        throw std::runtime_error("Cannot add column: sheet has no rows.");

    size_t insert_at = 0;

    if (at == "end")
        insert_at = sheet.cols.size();
    else if (at == "beginning" || at == "start")
        insert_at = 0;
    else
        insert_at = col_to_index(at);

    // ----------------------
    // Insert a new column at insert_at
    // ----------------------
    Column new_col;
    new_col.vals.resize(total_rows, ""); // empty column

    if (insert_at >= sheet.cols.size())
    {
        // append at the end
        sheet.cols.push_back(std::move(new_col));
    }
    else
    {
        // shift columns right
        sheet.cols.insert(sheet.cols.begin() + insert_at, std::move(new_col));
    }


    // ----------------------
    // Fill the new column
    // ----------------------
    fill_column_nitro(sheet, header_row, first_data_row, insert_at, fill_with, new_header);
}

// ----------------------
// Remove a column from NitroSheet
// ----------------------
void remove_column_nitro(
    NitroSheet &sheet,
    const size_t col_index         // 0-based column index
)
{
    if (col_index >= sheet.cols.size())
        return; // nothing to do

    sheet.cols.erase(sheet.cols.begin() + col_index);
}

// fast split by single char (avoids stringstream)
void split_simple(const std::string &s, char delim, std::vector<std::string> &out_parts)
{
    out_parts.clear();
    size_t start = 0;
    size_t n = s.size();
    for (size_t i = 0; i < n; ++i)
    {
        if (s[i] == delim)
        {
            out_parts.emplace_back(s.substr(start, i - start));
            start = i + 1;
        }
    }
    // last part (may be empty)
    out_parts.emplace_back(s.substr(start));
}

// Robust Nitro splitter
// Universal Nitro splitter with per-row normalization
void split_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t /*header_row*/,
    const std::uint32_t /*first_data_row*/,
    const std::size_t col_index,
    const char delimiter,
    const std::vector<size_t> &target_col_indices,
    const std::vector<std::string> &new_headers = {},
    const std::vector<std::uint32_t> &proper_positions = {}
)
{
    if (sheet.num_rows == 0 || col_index >= sheet.cols.size()) return;

    // ensure all target columns exist
    size_t max_target = col_index;
    for (size_t idx : target_col_indices)
        if (idx > max_target) max_target = idx;
    if (max_target >= sheet.cols.size()) sheet.cols.resize(max_target + 1);
    for (size_t t = 0; t <= max_target; ++t)
        if (sheet.cols[t].vals.size() < sheet.num_rows)
            sheet.cols[t].vals.resize(sheet.num_rows);

    Column &src = sheet.cols[col_index];
    if (src.vals.size() < sheet.num_rows)
        src.vals.resize(sheet.num_rows);

    std::vector<std::string> parts;
    parts.reserve(8);

    const size_t T = target_col_indices.size();

    for (size_t r = 0; r < sheet.num_rows; ++r)
    {
        const std::string &cell_value = src.vals[r];

        if (cell_value.empty())
        {
            for (size_t tcol : target_col_indices)
                sheet.cols[tcol].vals[r] = "";
            continue;
        }

        // split the value
        split_simple(cell_value, delimiter, parts);
        size_t N = parts.size();

        // ---- universal per-row normalization ----
        std::vector<std::string> normalized(T, "");

        if (N >= 1) normalized[0] = parts[0];             // first column = first part
        if (N >= 2) normalized[T-1] = parts[N-1];         // last column = last part

        // fill middle columns (1..T-2)
        for (size_t i = 1; i < T-1; ++i)
        {
            if (i < N-1)
                normalized[i] = parts[i];
            else
                normalized[i] = ""; // pad missing middle parts
        }
        // ----------------------------------------

        // assign values to target columns using proper_positions if provided
        for (size_t i = 0; i < T; ++i)
        {
            size_t pos = (!proper_positions.empty()) ? proper_positions[i] : (i + 1);

            std::string out_value = "";

            if (pos > 0 && pos <= normalized.size())
                out_value = normalized[pos - 1];

            sheet.cols[target_col_indices[i]].vals[r] = out_value;
        }
    }

    // set headers if provided
    if (!new_headers.empty())
    {
        for (size_t i = 0; i < T; ++i)
        {
            size_t tcol = target_col_indices[i];
            if (tcol >= sheet.cols.size()) continue;
            sheet.cols[tcol].header = (i < new_headers.size()) ? new_headers[i] : "";
        }
    }
}


void uppercase_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t first_data_row,  // 1-based row index
    const std::size_t col_index          // 0-based column index
)
{
    const size_t total_rows = sheet.num_rows;
    if (total_rows == 0 || col_index >= sheet.cols.size())
        return;

    size_t data_start = (first_data_row > 0 ? first_data_row - 1 : 1);

    Column &col = sheet.cols[col_index];

    for (size_t r = 0; r < total_rows; ++r)
    {
        std::string &val = col.vals[r];
        if (!val.empty())
            std::transform(val.begin(), val.end(), val.begin(), ::toupper);
    }
}

// helper trim function
static inline void trim_inplace(std::string &s)
{
    auto not_space = [](unsigned char c){ return !std::isspace(c); };

    // trim left
    s.erase(s.begin(), std::find_if(s.begin(), s.end(), not_space));

    // trim right
    s.erase(std::find_if(s.rbegin(), s.rend(), not_space).base(), s.end());
}

// ----------------------
// Replace substring in a NitroSheet column
// ----------------------
void replace_in_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t first_data_row,  // 1-based row index
    const std::size_t col_index,         // 0-based column index
    const std::string &find,
    const std::string &repl
)
{
    const size_t total_rows = sheet.num_rows;
    if (total_rows == 0 || col_index >= sheet.cols.size() || find.empty())
        return;

    size_t data_start = (first_data_row > 0 ? first_data_row - 1 : 1);

    Column &col = sheet.cols[col_index];

    for (size_t r = data_start; r < total_rows; ++r)
    {
        std::string &val = col.vals[r];
        if (val.empty()) continue;

        // 🧹 trim before replace
        trim_inplace(val);
        if (val.empty()) continue;

        size_t pos = 0;
        while ((pos = val.find(find, pos)) != std::string::npos)
        {
            val.replace(pos, find.size(), repl);
            pos += repl.size();
        }
    }
}

// ----------------------
// Apply transformations to a NitroSheet row
// ----------------------
void transform_row_nitro(
    NitroSheet &sheet,
    const std::size_t row_index,
    const std::string &to,
    std::optional<char> delim
)
{
    if (sheet.cols.empty() || row_index >= sheet.num_rows)
        return;

    for (size_t col = 0; col < sheet.cols.size(); ++col)
    {
        Column &column = sheet.cols[col];

        // Ensure this column has storage for all rows
        if (column.vals.size() < sheet.num_rows)
            column.vals.resize(sheet.num_rows);

        // Now it's safe
        std::string &val = column.vals[row_index];
        if (val.empty()) continue;

        if (to == "camelCase")
        {
            if (!delim) continue;
            val = to_camel(val, *delim);
        }
        else if (to == "PascalCase")
        {
            if (!delim) continue;
            val = to_pascal(val, *delim);
        }
        else if (to == "snake_case")
        {
            val = to_snake(val);
        }
        else if (to == "upper")
        {
            std::transform(val.begin(), val.end(), val.begin(), ::toupper);
        }
        else if (to == "lower")
        {
            std::transform(val.begin(), val.end(), val.begin(), ::tolower);
        }
        else
        {
            throw std::runtime_error("Unknown transform: " + to);
        }
    }
}


void transform_header_nitro(
    NitroSheet &sheet,
    const std::string &to,           // "camelCase", "PascalCase", "snake_case", "upper", "lower"
    std::optional<char> delim
)
{
    if (sheet.cols.empty())
        return;

    for (size_t col = 0; col < sheet.cols.size(); ++col)
    {
        Column &column = sheet.cols[col];

        // Ensure this column has storage for all rows
        if (column.vals.size() < sheet.num_rows)
            column.vals.resize(sheet.num_rows);

        // Now it's safe
        std::string &val = column.header;
        if (val.empty()) continue;

        if (to == "camelCase")
        {
            if (!delim) continue;
            val = to_camel(val, *delim);
        }
        else if (to == "PascalCase")
        {
            if (!delim) continue;
            val = to_pascal(val, *delim);
        }
        else if (to == "snake_case")
        {
            val = to_snake(val);
        }
        else if (to == "upper")
        {
            std::transform(val.begin(), val.end(), val.begin(), ::toupper);
        }
        else if (to == "lower")
        {
            std::transform(val.begin(), val.end(), val.begin(), ::tolower);
        }
        else
        {
            throw std::runtime_error("Unknown transform header: " + to);
        }
    }
}

void rename_header_nitro(
    NitroSheet &sheet,
    const std::size_t col_index,
    const std::string new_name)
{
    if (sheet.cols.empty())
        return;

    Column &col = sheet.cols[col_index];

    col.header = new_name;
}

void group_collect_nitro(
    NitroSheet &sheet,
    std::size_t group_col,
    std::size_t collect_col,
    std::size_t output_col,
    bool marked_unique,
    std::size_t do_maths_col,
    const std::string &do_maths_operation)
{
    if (sheet.cols.empty())
        return;

    std::string current_key;
    std::vector<std::string> collected;

    // temporary storage of numeric values for math ops
    std::vector<double> math_values;

    std::vector<bool> keep_row(sheet.num_rows, true);

    std::size_t first_row_index = 0;

    auto flush_group = [&](size_t first_row)
    {
        // --- Build JSON array for collected column ---
        std::string json = "[";
        std::size_t collected_count_for_last_comma = collected.size();
        bool add_comma = true;
        for (size_t i = 0; i < collected.size(); ++i)
        {
            if (marked_unique)
            {
                // check if already exists in previous items
                bool exists = false;
                for (size_t j = 0; j < i; ++j)
                {
                    if (collected[j] == collected[i])
                    {
                        exists = true;
                        break;
                    }
                }
                if (exists) continue; // skip duplicate
                else if (i > 0 && json.size() > 1) json += ",";\
            }
            else if (i > 0 && json.size() > 1) json += ",";

            if (collected[i].find('{') != std::string::npos ||
                collected[i].find('[') != std::string::npos)
            {
                // assume already JSON formatted
                json += collected[i];
            }
            else
                json += "\"" + collected[i] +  "\"";
        }
        json += "]";
        sheet.cols[output_col].vals[first_row] = json;

        // --- Perform Maths Operation ---
        if (!do_maths_operation.empty() && !math_values.empty())
        {
            double result = 0.0;

            if (do_maths_operation == "sum")
            {
                for (double v : math_values) result += v;
            }
            else if (do_maths_operation == "avg")
            {
                for (double v : math_values) result += v;
                result /= math_values.size();
            }
            else if (do_maths_operation == "min")
            {
                result = *std::min_element(math_values.begin(), math_values.end());
            }
            else if (do_maths_operation == "max")
            {
                result = *std::max_element(math_values.begin(), math_values.end());
            }
            else if (do_maths_operation == "count")
            {
                result = static_cast<double>(math_values.size());
            }
            else
            {
                throw std::runtime_error("Unknown maths operation: " + do_maths_operation);
            }

            sheet.cols[do_maths_col].vals[first_row] = std::to_string(result);
        }
    };

    // Pass 1: collect values + mark duplicate rows for deletion
    for (std::size_t r = 0; r < sheet.num_rows; ++r)
    {
        const std::string &key = sheet.cols[group_col].vals[r];
        const std::string &val = sheet.cols[collect_col].vals[r];
        const std::string &math_val_str = sheet.cols[do_maths_col].vals[r];

        if (r == 0)
        {
            current_key = key;
            if (!val.empty()) collected.push_back(val);

            if (!math_val_str.empty())
            {
                try {
                    math_values.push_back(std::stod(math_val_str));
                } catch (...) {}
            }
            continue;
        }

        if (key != current_key)
        {
            // flush previous group
            flush_group(first_row_index);

            // start new group
            current_key = key;
            collected.clear();
            math_values.clear();

            if (!val.empty()) collected.push_back(val);
            if (!math_val_str.empty()) {
                try { math_values.push_back(std::stod(math_val_str)); } catch (...) {}
            }

            first_row_index = r;
        }
        else
        {
            // same group → keep only first row
            keep_row[r] = false;

            if (!val.empty()) collected.push_back(val);
            if (!math_val_str.empty())
            {
                try { math_values.push_back(std::stod(math_val_str)); } catch (...) {}
            }
        }
    }

    // Final flush
    if (!collected.empty())
        flush_group(first_row_index);

    // Pass 2: Physically remove duplicate rows
    size_t write_index = 0;
    for (size_t r = 0; r < sheet.num_rows; ++r)
    {
        if (keep_row[r])
        {
            if (write_index != r)
            {
                for (auto &col : sheet.cols)
                    col.vals[write_index] = std::move(col.vals[r]);
            }
            write_index++;
        }
    }

    // Resize all columns
    for (auto &col : sheet.cols)
        col.vals.resize(write_index);

    sheet.num_rows = write_index;
}

void sort_rows_by_column_nitro(
    NitroSheet &sheet,
    const std::size_t col_index,
    const bool ascending
)
{
    if (sheet.cols.empty() || col_index >= sheet.cols.size())
        return;

    const size_t total_rows = sheet.num_rows;
    if (total_rows == 0)
        return;

    // Create index vector
    std::vector<size_t> indices(total_rows);
    for (size_t i = 0; i < total_rows; ++i)
        indices[i] = i;

    // Sort indices based on the target column's values
    Column &col = sheet.cols[col_index];
    std::sort(indices.begin(), indices.end(),
        [&](size_t a, size_t b) {
            if (ascending)
                return col.vals[a] < col.vals[b];
            else
                return col.vals[a] > col.vals[b];
        });

    // Reorder all columns based on sorted indices
    for (auto &column : sheet.cols)
    {
        std::vector<std::string> sorted_vals(total_rows);
        for (size_t i = 0; i < total_rows; ++i)
        {
            sorted_vals[i] = std::move(column.vals[indices[i]]);
        }
        column.vals = std::move(sorted_vals);
    }
}

void reassign_numbering_nitro(
    NitroSheet &sheet,
    const std::size_t col_index,
    const std::string &prefix,
    const std::string &suffix,
    const size_t start_number,
    const size_t step
)
{
    if (sheet.cols.empty() || col_index >= sheet.cols.size())
        return;

    Column &col = sheet.cols[col_index];

    size_t number = start_number;

    for (size_t r = 0; r < sheet.num_rows; ++r)
    {
        col.vals[r] = prefix + std::to_string(number) + suffix;
        number += step;
    }
}