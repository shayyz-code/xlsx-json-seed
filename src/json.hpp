#pragma once
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <string>
#include <vector>
#include <cstdint>

// ------------------------------------------------------
// Escape JSON string characters safely
// ------------------------------------------------------
inline std::string json_escape(const std::string &s)
{
    std::string out;
    out.reserve(s.size() + 16);

    for (unsigned char c : s)
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
            if (c < 0x20)
            {
                char buf[7];
                std::snprintf(buf, sizeof(buf), "\\u%04x", c);
                out += buf;
            }
            else out += static_cast<char>(c);
        }
    }
    return out;
}

// ------------------------------------------------------
// JSON Exporter with configurable header + start row
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

    // Calculate real used range
    xlnt::range_reference dims = ws.calculate_dimension();

    const uint32_t first_row = dims.top_left().row();
    const uint32_t last_row  = dims.bottom_right().row();
    const uint32_t first_col = dims.top_left().column().index;
    const uint32_t last_col  = dims.bottom_right().column().index;

    if (header_row < first_row || header_row > last_row)
        throw std::runtime_error("Configured header_row is outside worksheet bounds.");

    if (first_data_row < first_row)
        first_data_row = first_row + 1; // safety fallback

    // ------------------------------------------------------
    // Read header row into vector
    // ------------------------------------------------------
    std::vector<std::string> headers;
    headers.reserve(last_col - first_col + 1);

    for (uint32_t col = first_col; col <= last_col; ++col)
    {
        auto cell = ws.cell(col, header_row);
        std::string h = cell.has_value() ? cell.to_string() : "";

        if (h.empty())
            throw std::runtime_error(
                "Header row contains an empty column title at column "
                + std::to_string(col)
            );

        headers.push_back(h);
    }

    // ------------------------------------------------------
    // Begin JSON array output
    // ------------------------------------------------------
    out << "[\n";
    bool first_output = true;

    // ------------------------------------------------------
    // Iterate over data rows
    // ------------------------------------------------------
    for (uint32_t row = first_data_row; row <= last_row; ++row)
    {
        // Skip completely empty rows quickly
        bool empty_row = true;
        for (uint32_t col = first_col; col <= last_col; ++col)
        {
            if (ws.cell(col, row).has_value())
            {
                empty_row = false;
                break;
            }
        }
        if (empty_row)
            continue;

        if (!first_output)
            out << ",\n";
        first_output = false;

        out << "  {\n";

        // Each column
        for (std::size_t i = 0; i < headers.size(); ++i)
        {
            uint32_t col = first_col + i;

            std::string key = headers[i];
            xlnt::cell cell = ws.cell(col, row);

            std::string val = cell.has_value() ? cell.to_string() : "";

            out << "    \"" << json_escape(key) << "\": ";

            // Detect Firestore timestamp placeholder `{ "__fire_ts_from_date__": "..."}`
            auto emit_value = [&](const std::string &v)
            {
                if (
                    v.size() > 2 &&
                    v.front() == '{' &&
                    v.back() == '}' &&
                    (v.find("__fire_ts_from_date__") != std::string::npos)
                ) {
                    out << v;   // raw JSON
                }
                else {
                    out << "\"" << json_escape(v) << "\"";
                }
            };

            emit_value(val);

            if (i + 1 < headers.size())
                out << ",";
            out << "\n";
        }


        out << "  }";
    }

    out << "\n]\n";
}
