#include "operations.hpp"
#include <algorithm>
#include <sstream>

// helpers
int col_to_index(const std::string &col)
{
    return col[0] - 'A';
}

std::string to_upper(const std::string &str)
{
    std::string r = str;
    std::transform(r.begin(), r.end(), r.begin(), [](unsigned char c) { return std::toupper(c);});

    return r;
}

std::string to_lower(const std::string &str)
{
    std::string r = str;
    std::transform(r.begin(), r.end(), r.begin(), [](unsigned char c) { return std::tolower(c);});

    return r;
}

std::string to_camel(const std::string &str, char &delim)
{
    std::string r = to_lower(str);
    bool cap_next = false;

    for (size_t i = 0; i < r.size(); i++)
    {
        if (r[i] == delim)
        {
            cap_next = true;
        }
        else if (cap_next)
        {
            r[i] = std::toupper(r[i]);
            cap_next = false;
        }
    }

    r.erase(std::remove(r.begin(), r.end(), delim), r.end());

    return r;
}


std::string to_pascal(const std::string &str, char &delim)
{
    std::string r = to_camel(str, delim);
    if (!r.empty())
    {
        r[0] = std::toupper(r[0]);
    }
    return r;
}

// helper for to_snake
static inline bool is_sep(char c)
{
    return std::isspace((unsigned char)c) || std::ispunct((unsigned char)c);
}

std::string to_snake(const std::string &str)
{
    std::string out;
    out.reserve(str.size() * 2);

    bool last_was_sep = false;

    for (size_t i = 0; i < str.size(); ++i)
    {
        unsigned char c = str[i];

        // ---- Separator / punctuation / space ----
        if (is_sep(c))
        {
            if (!out.empty() && !last_was_sep)
            {
                out += '_';
                last_was_sep = true;
            }
            continue;
        }

        // ---- Uppercase letter and not first char ----
        if (std::isupper(c))
        {
            if (!out.empty() && !last_was_sep)
            {
                out += '_';
            }
            out += (char)std::tolower(c);
            last_was_sep = false;
            continue;
        }

        // ---- Lowercase or digit ----
        out += (char)std::tolower(c);
        last_was_sep = false;
    }

    // Trim trailing underscore
    while (!out.empty() && out.back() == '_')
        out.pop_back();

    return out;
}

// ops

void split_column_xlsx(
    xlnt::worksheet &ws,
    const std::uint32_t header_row,
    const std::uint32_t first_data_row,
    const std::string &source,
    const std::string &delimiter,
    const std::vector<std::string> &targets,
    const std::vector<std::string> &new_headers)
{
    auto max_row = ws.highest_row();

    for (std::uint32_t row = first_data_row; row <= max_row; ++row)
    {
        std::string cell_ref = source + std::to_string(row);

        if (!ws.has_cell(cell_ref))
            continue;

        std::string value = ws.cell(cell_ref).to_string();
        
        if (value.empty())
            continue;

        std::vector<std::string> parts;
        std::stringstream ss(value);
        std::string segment;

        while (std::getline(ss, segment, delimiter[0]))
            parts.push_back(segment);

        for (std::size_t i = 0; i < targets.size(); i++)
        {
            std::string target = targets[i] + std::to_string(row);
            ws.cell(target).value(i < parts.size() ? parts[i] : "");
        }
    }

    // Set new headers if provided
    if (!new_headers.empty())
    {
        for (std::size_t i = 0; i < targets.size(); i++)
        {
            std::string target = targets[i] + std::to_string(header_row);
            ws.cell(target).value(i < new_headers.size() ? new_headers[i] : "");
        }
    }
}

void uppercase_column_xlsx(xlnt::worksheet &ws, const std::uint32_t first_data_row, const std::string &column)
{
    auto max_row = ws.highest_row();

    for (std::uint32_t row = first_data_row; row <= max_row; ++row)
    {
        std::string cell_ref = column + std::to_string(row);

        if (!ws.has_cell(cell_ref))
            continue;

        auto val = ws.cell(cell_ref).to_string();

        ws.cell(cell_ref).value(to_upper(val));
    }
}

void replace_in_column_xlsx(
    xlnt::worksheet &ws,
    const std::uint32_t first_data_row,
    const std::string &column,
    const std::string &find,
    const std::string &repl)
{
    auto max_row = ws.highest_row();

    for (std::uint32_t row = first_data_row; row <= max_row; ++row)
    {
        std::string cell_ref = column + std::to_string(row);

        if (!ws.has_cell(cell_ref))
            continue;

        auto val = ws.cell(cell_ref).to_string();

        size_t pos = 0;
        while ((pos = val.find(find, pos)) != std::string::npos)
        {
            val.replace(pos, find.length(), repl);
            pos += repl.length();
        }

        ws.cell(cell_ref).value(val);
    }
}

void transform_row_xlsx(
    xlnt::worksheet &ws,
    std::uint32_t row,
    const std::string &to,
    std::optional<char> delim
)
{
    auto max_col = ws.highest_column().index; // 1-based

    for (std::uint32_t col = 1; col <= max_col; ++col)
    {
        xlnt::column_t col_t(col);               // convert integer → xlnt column handle
        xlnt::cell_reference ref(col_t, row);    // (column, row)

        if (!ws.has_cell(ref))
            continue;

        xlnt::cell cell = ws.cell(ref);
        std::string val = cell.to_string();

        // ---- Apply transform ----
        if (to == "camelCase")
        {
            if (!delim) continue;
            cell.value(to_camel(val, *delim));
        }
        else if (to == "PascalCase")
        {
            if (!delim) continue;
            cell.value(to_pascal(val, *delim));
        }
        else if (to == "snake_case")
        {
            cell.value(to_snake(val));
        }
        else if (to == "upper")
        {
            cell.value(to_upper(val));
        }
        else if (to == "lower")
        {
            cell.value(to_lower(val));
        }
        else
        {
            throw std::runtime_error("Unknown transform: " + to);
        }
    }
}

