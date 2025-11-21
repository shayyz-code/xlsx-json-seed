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

void split_column_xlsx(
    xlnt::worksheet &ws,
    const std::string &source,
    const std::string &delimiter,
    const std::vector<std::string> &targets)
{
    auto max_row = ws.highest_row();

    for (std::size_t row = 1; row <= max_row; ++row)
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
}

void uppercase_column_xlsx(xlnt::worksheet &ws, const std::string &column)
{
    auto max_row = ws.highest_row();

    for (std::size_t row = 1; row <= max_row; ++row)
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
    const std::string &column,
    const std::string &find,
    const std::string &repl)
{
    auto max_row = ws.highest_row();

    for (std::size_t row = 1; row <= max_row; ++row)
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