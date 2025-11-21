#pragma once
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <string>

inline void save_csv(const xlnt::worksheet &ws, const std::string &path)
{
    std::ofstream out(path);
    if (!out.is_open())
        throw std::runtime_error("Cannot write CSV: " + path);

    for (auto row : ws.rows(false))   // false = skip merged cell expansion
    {
        bool first = true;
        for (auto cell : row)
        {
            if (!first)
                out << ",";
            first = false;

            std::string v = cell.to_string();

            // Escape quotes & commas
            bool needs_quotes =
                v.find(',') != std::string::npos ||
                v.find('"') != std::string::npos;

            if (needs_quotes)
            {
                std::string esc = "\"";
                for (char c : v)
                {
                    if (c == '"') esc += "\"\"";
                    else esc += c;
                }
                esc += "\"";
                out << esc;
            }
            else
            {
                out << v;
            }
        }
        out << "\n";
    }
}
