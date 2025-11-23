#pragma once
#include <string>
#include <fstream>
#include <stdexcept>
#include <vector>

// RFC-4180 CSV escape
inline std::string csv_escape(const std::string &s)
{
    bool needs_quotes = false;

    for (char c : s)
    {
        if (c == '"' || c == ',' || c == '\n' || c == '\r')
        {
            needs_quotes = true;
            break;
        }
    }

    if (!needs_quotes)
        return s;

    // Escape double quotes
    std::string out;
    out.reserve(s.size() + 8);
    out.push_back('"');

    for (char c : s)
    {
        if (c == '"')
            out += "\"\"";   // escape "
        else
            out.push_back(c);
    }

    out.push_back('"');
    return out;
}

inline void save_csv_nitro(
    const NitroSheet &sheet,
    const std::string &path,
    size_t flush_threshold = 1 << 20
)
{
    if (sheet.cols.empty())
        throw std::runtime_error("CSV export failed: sheet has no columns.");

    const size_t rows = sheet.num_rows;
    const size_t cols = sheet.cols.size();

    // Validate column structures
    for (size_t c = 0; c < cols; ++c)
    {
        if (sheet.cols[c].header.empty())
            throw std::runtime_error("CSV export: column " +
                                     std::to_string(c) +
                                     " has an empty header.");

        if (sheet.cols[c].vals.size() < rows)
            throw std::runtime_error("CSV export: column " +
                                     std::to_string(c) +
                                     " vals smaller than sheet.num_rows.");
    }

    std::ofstream out(path, std::ios::binary);
    if (!out.is_open())
        throw std::runtime_error("Cannot open CSV file: " + path);

    std::string buf;
    buf.reserve(1024 * 1024);

    auto flush_buf = [&](bool force=false){
        if (!buf.empty()) {
            out.write(buf.data(), (std::streamsize)buf.size());
            buf.clear();
        }
        if (force) out.flush();
    };

    // --------------------------
    // Write header row
    // --------------------------
    for (size_t c = 0; c < cols; ++c)
    {
        buf += csv_escape(sheet.cols[c].header);
        if (c + 1 < cols) buf += ",";
    }
    buf += "\n";
    flush_buf();

    // --------------------------
    // Write data rows
    // --------------------------
    for (size_t r = 0; r < rows; ++r)
    {
        bool empty = true;
        for (size_t c = 0; c < cols; ++c)
            if (!sheet.cols[c].vals[r].empty())
                empty = false;

        if (empty) continue;

        for (size_t c = 0; c < cols; ++c)
        {
            buf += csv_escape(sheet.cols[c].vals[r]);
            if (c + 1 < cols) buf += ",";
        }
        buf += "\n";

        if (buf.size() >= flush_threshold)
            flush_buf();
    }

    flush_buf(true);
    out.close();
}
