#include "operations.hpp"
#include <algorithm>
#include <sstream>

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
        if (fill_with == "firestore-now")
        {
            col.vals[r] = "__fire_ts_now__";
        }
        else if (fill_with.compare(0, prefix.size(), prefix) == 0)
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
void split_column_nitro(
    NitroSheet &sheet,
    const std::uint32_t /*header_row*/,          // kept for API symmetry (not used)
    const std::uint32_t /*first_data_row*/,      // kept for API symmetry (not used)
    const std::size_t col_index,                 // 0-based column index to split
    const char delimiter,                        // delimiter character
    const std::vector<size_t> &target_col_indices,  // 0-based target column indices
    const std::vector<std::string> &new_headers = {} // optional headers for target columns
)
{
    // validate
    if (sheet.num_rows == 0) return;
    if (col_index >= sheet.cols.size()) return;

    // ensure all target columns exist and have correct vals size
    size_t max_target = col_index;
    for (size_t idx : target_col_indices) if (idx > max_target) max_target = idx;

    if (max_target >= sheet.cols.size()) sheet.cols.resize(max_target + 1);

    for (size_t t = 0; t <= max_target; ++t)
    {
        if (sheet.cols[t].vals.size() < sheet.num_rows)
            sheet.cols[t].vals.resize(sheet.num_rows);
    }

    Column &src = sheet.cols[col_index];

    // safety: make sure source vals vector has expected length
    if (src.vals.size() < sheet.num_rows)
        src.vals.resize(sheet.num_rows);

    std::vector<std::string> parts;
    parts.reserve(4);

    // iterate over data rows (Nitro convention: vals[0] == first data row)
    for (size_t r = 0; r < sheet.num_rows; ++r)
    {
        const std::string &value = src.vals[r];

        // fast skip if empty
        if (value.empty())
        {
            // assign empty string to targets (consistent)
            for (size_t i = 0; i < target_col_indices.size(); ++i)
            {
                size_t tcol = target_col_indices[i];
                sheet.cols[tcol].vals[r] = "";
            }
            continue;
        }

        // split
        split_simple(value, delimiter, parts);

        // write into targets
        for (size_t i = 0; i < target_col_indices.size(); ++i)
        {
            size_t tcol = target_col_indices[i];
            // bounds should be guaranteed above; but check defensively
            if (tcol >= sheet.cols.size()) continue;
            sheet.cols[tcol].vals[r] = (i < parts.size()) ? parts[i] : std::string();
        }
    }

    // set headers for targets if provided (write to Column::header)
    if (!new_headers.empty())
    {
        for (size_t i = 0; i < target_col_indices.size(); ++i)
        {
            size_t tcol = target_col_indices[i];
            if (tcol >= sheet.cols.size()) continue;
            sheet.cols[tcol].header = (i < new_headers.size()) ? new_headers[i] : std::string();
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