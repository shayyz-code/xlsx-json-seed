#pragma once
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <string>
#include <vector>
#include <cstdint>

// ------------------------------------------------------
// Escape CSV fields according to RFC 4180
// ------------------------------------------------------
inline std::string csv_escape(const std::string &s)
{
    bool needs_quotes =
        s.find(',') != std::string::npos ||
        s.find('"') != std::string::npos ||
        s.find('\n') != std::string::npos ||
        s.find('\r') != std::string::npos;

    if (!needs_quotes)
        return s;

    std::string out = "\"";
    for (char c : s)
    {
        if (c == '"')
            out += "\"\"";  // escape double-quote
        else
            out += c;
    }
    out += "\"";

    return out;
}

// ------------------------------------------------------
// CSV Exporter with header & start row
// ------------------------------------------------------
inline void save_csv(
    const xlnt::worksheet &ws,
    uint32_t header_row,
    uint32_t first_data_row,
    const std::string &path
)
{
    std::ofstream out(path);
    if (!out.is_open())
        throw std::runtime_error("Cannot write CSV: " + path);

    // Get sheet bounds
    xlnt::range_reference dims = ws.calculate_dimension();

    uint32_t first_row = dims.top_left().row();
    uint32_t last_row  = dims.bottom_right().row();
    uint32_t first_col = dims.top_left().column().index;
    uint32_t last_col  = dims.bottom_right().column().index;

    if (header_row < first_row || header_row > last_row)
        throw std::runtime_error("Configured header_row is outside worksheet bounds");

    if (first_data_row < first_row)
        first_data_row = first_row + 1; // safety fallback

    // ------------------------------------------------------
    // Extract header titles
    // ------------------------------------------------------
    std::vector<std::string> headers;
    headers.reserve(last_col - first_col + 1);

    for (uint32_t col = first_col; col <= last_col; ++col)
    {
        xlnt::cell c = ws.cell(col, header_row);
        std::string h = c.has_value() ? c.to_string() : "";

        if (h.empty())
            throw std::runtime_error(
                "Header row contains an empty column title at column " + std::to_string(col)
            );

        headers.push_back(h);
    }

    // ------------------------------------------------------
    // Write header line
    // ------------------------------------------------------
    for (std::size_t i = 0; i < headers.size(); ++i)
    {
        out << csv_escape(headers[i]);
        if (i + 1 < headers.size()) out << ",";
    }
    out << "\n";

    // ------------------------------------------------------
    // Write data rows
    // ------------------------------------------------------
    for (uint32_t row = first_data_row; row <= last_row; ++row)
    {
        // Skip empty rows
        bool empty = true;
        for (uint32_t col = first_col; col <= last_col; ++col)
        {
            if (ws.cell(col, row).has_value())
            {
                empty = false;
                break;
            }
        }
        if (empty)
            continue;

        // Write each column
        for (std::size_t i = 0; i < headers.size(); ++i)
        {
            uint32_t col = first_col + i;
            xlnt::cell c = ws.cell(col, row);

            std::string val = c.has_value() ? c.to_string() : "";

            // We allow Firestore JSON objects inside CSV cells
            // e.g. { "__fire_ts_from_date__": "..." }
            // They get escaped normally with csv_escape().
            out << csv_escape(val);

            if (i + 1 < headers.size())
                out << ",";
        }

        out << "\n";
    }
}
