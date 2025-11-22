#pragma once
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <string>
#include <vector>
#include <cstdint>

// Escape JSON characters
inline std::string json_escape(const std::string &s)
{
    std::string out;
    out.reserve(s.size() + 8);

    for (char c : s)
    {
        switch (c)
        {
        case '\"': out += "\\\""; break;
        case '\\': out += "\\\\"; break;
        case '\b': out += "\\b"; break;
        case '\f': out += "\\f"; break;
        case '\n': out += "\\n"; break;
        case '\r': out += "\\r"; break;
        case '\t': out += "\\t"; break;
        default:
            if ((unsigned char)c < 0x20)
            {
                char buf[7];
                std::snprintf(buf, sizeof(buf), "\\u%04x", c);
                out += buf;
            }
            else out += c;
        }
    }
    return out;
}

// ------------------------------------------------------
// JSON Exporter with configurable header row
// ------------------------------------------------------
inline void save_json(
    const xlnt::worksheet &ws,
    uint32_t header_row,
    uint32_t first_data_row,
    const std::string &path
)
{
    std::ofstream out(path);
    if (!out.is_open())
        throw std::runtime_error("Cannot write JSON: " + path);

    // Get sheet dimensions
    xlnt::range_reference dims = ws.calculate_dimension();

    uint32_t first_row = dims.top_left().row();
    uint32_t last_row  = dims.bottom_right().row();
    uint32_t first_col = dims.top_left().column().index;
    uint32_t last_col  = dims.bottom_right().column().index;

    if (header_row < first_row || header_row > last_row)
        throw std::runtime_error("Configured header row is outside worksheet bounds");

    // ------------------------------------------------------
    // Extract header row (using user-defined header_row)
    // ------------------------------------------------------
    std::vector<std::string> headers;
    headers.reserve(last_col - first_col + 1);

    for (uint32_t col = first_col; col <= last_col; ++col)
    {
        std::string h = ws.cell(col, header_row).to_string();
        if (h.empty())
            throw std::runtime_error("Header row contains empty column titles");
        headers.push_back(h);
    }

    out << "[\n";
    bool first_output = true;

    // ------------------------------------------------------
    // Serialize rows FROM first data row
    // ------------------------------------------------------
    for (uint32_t row = first_data_row; row <= last_row; ++row)
    {
        // Skip fully empty rows
        bool empty = true;
        for (uint32_t col = first_col; col <= last_col; ++col)
        {
            if (!ws.cell(col, row).to_string().empty())
            {
                empty = false;
                break;
            }
        }
        if (empty)
            continue;

        if (!first_output) out << ",\n";
        first_output = false;

        out << "  {\n";

        // Write columns according to header titles
        for (std::size_t i = 0; i < headers.size(); i++)
        {
            uint32_t col = first_col + i;
            std::string key = headers[i];
            std::string val = ws.cell(col, row).to_string();

            out << "    \"" << json_escape(key) << "\": "
                << "\"" << json_escape(val) << "\"";

            if (i + 1 < headers.size())
                out << ",";
            out << "\n";
        }

        out << "  }";
    }

    out << "\n]\n";
}
