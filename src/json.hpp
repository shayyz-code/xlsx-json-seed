#pragma once
#include <string>
#include <vector>
#include <fstream>
#include <algorithm>
#include <cctype>
#include <stdexcept>

// Assumes Column { std::string header; std::vector<std::string> vals; bool dirty; }
// and NitroSheet { std::vector<Column> cols; uint32_t first_row; uint32_t data_row_start; uint32_t num_rows; }

// trim helper
inline std::string trim_copy(const std::string &s) {
    size_t b = 0, e = s.size();
    while (b < e && std::isspace((unsigned char)s[b])) ++b;
    while (e > b && std::isspace((unsigned char)s[e-1])) --e;
    return s.substr(b, e - b);
}

// json escape helper (same as before)
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
            else out += (char)c;
        }
    }
    return out;
}

inline bool looks_like_obj(const std::string &trimmed)
{
    if (trimmed.size() < 2) return false;
    return (trimmed.front() == '{' || trimmed.back() == '}');
}

inline bool looks_like_array(const std::string &trimmed)
{
    if (trimmed.size() < 2) return false;
    return (trimmed.front() == '[' && trimmed.back() == ']');
}

inline bool is_valid_number(const std::string& trimmed)
{
    if (trimmed.size() < 1) return false;

    // Fast check: reject trailing dot
    if (trimmed.back() == '.') 
        return false;

    char* end = nullptr;
    std::strtod(trimmed.c_str(), &end);

    // true only if full string parsed as number
    return end == trimmed.c_str() + trimmed.size();
}


// ---------- Robust save_json_nitro that matches NitroSheet layout ----------
inline void save_json_nitro(
    const NitroSheet &sheet,
    uint32_t header_row,        // kept for API compatibility; not used to index vals
    uint32_t first_data_row,    // kept for API compatibility
    const std::string &path,
    bool pretty = true,
    size_t flush_threshold = 1 << 20
)
{
    // print_nitro_sheet(sheet); // uncomment this for debugging

    // Basic validation
    if (sheet.cols.empty())
        throw std::runtime_error("Cannot export JSON: sheet has no columns.");

    const size_t data_rows = sheet.num_rows; // number of data rows in each column (vals.size())
    if (data_rows == 0)
        throw std::runtime_error("Sheet has 0 data rows.");

    const size_t cols = sheet.cols.size();

    

    // Ensure every column has header (we rely on Column::header)
    for (size_t c = 0; c < cols; ++c)
    {
        if (sheet.cols[c].header.empty())
            throw std::runtime_error("Header missing for column index " + std::to_string(c));
    }

    // Ensure vals vectors are large enough for safe indexing
    for (size_t c = 0; c < cols; ++c)
    {
        // const_cast because function receives const NitroSheet; if you want to mutate, accept non-const.
        // Better: require non-const NitroSheet or ensure loader already resized. Here we'll require non-const from caller.
        // To keep signature const, we will check only — but to avoid segfaults the loader MUST ensure sizing.
        if (sheet.cols[c].vals.size() < data_rows)
            throw std::runtime_error("Column " + std::to_string(c) + " vals size (" +
                                     std::to_string(sheet.cols[c].vals.size()) +
                                     ") is smaller than sheet.num_rows (" + std::to_string(data_rows) + ").");
    }

    // Open output stream
    std::ofstream out(path, std::ios::binary);
    if (!out.is_open())
        throw std::runtime_error("Cannot write JSON: " + path);

    std::string buf;
    buf.reserve(1024 * 1024);
    auto flush_buf = [&](bool force=false){
        if (!buf.empty()) {
            out.write(buf.data(), static_cast<std::streamsize>(buf.size()));
            buf.clear();
        }
        if (force) out.flush();
    };

    const std::string nl      = pretty ? "\n" : "";
    const std::string ind1    = pretty ? "  " : "";
    const std::string ind2    = pretty ? "    " : "";

    buf += "[" + nl;
    bool first_obj = true;

    // Iterate data rows: Nitro stores only data rows in vals[0..data_rows-1]
    for (size_t r = 0; r < data_rows; ++r)
    {
        // skip fully empty row
        bool empty = true;
        for (size_t c = 0; c < cols; ++c)
        {
            if (!sheet.cols[c].vals[r].empty()) { empty = false; break; }
        }
        if (empty) continue;

        if (!first_obj) buf += "," + nl;
        first_obj = false;

        buf += ind1 + "{" + nl;

        for (size_t c = 0; c < cols; ++c)
        {
            const std::string &key = sheet.cols[c].header;
            const std::string &raw = sheet.cols[c].vals[r];
            std::string trimmed = trim_copy(raw);

            buf += ind2 + "\"" + json_escape(key) + "\": ";

            if (!trimmed.empty() && (looks_like_obj(trimmed) || looks_like_array(trimmed)))
            {
                buf += trimmed;
            }
            else if (!trimmed.empty() && (trimmed == "null" || trimmed == "true" || trimmed == "false"
                     || is_valid_number(trimmed)))
            {
                // raw number, bool, or null
                buf += to_clean_number(trimmed);
            }
            else
            {
                buf += "\"" + json_escape(raw) + "\"";
            }

            if (c + 1 < cols) buf += ",";
            buf += nl;

            if (buf.size() >= flush_threshold) flush_buf();
        }

        buf += ind1 + "}";
        if (buf.size() >= flush_threshold) flush_buf();
    }

    buf += nl + "]" + nl;
    flush_buf(true);
    out.close();
}
