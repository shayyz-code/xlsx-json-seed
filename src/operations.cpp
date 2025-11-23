#include "operations.hpp"
#include <algorithm>
#include <sstream>
#include "utils.hpp"


// ops

void fill_column_xlsx(
    xlnt::worksheet &ws,
    const std::uint32_t header_row,
    const std::uint32_t first_data_row,
    const std::string &column,
    const std::string &fill_with,
    const std::string &new_header)
{
    auto max_row = ws.highest_row();
    if (first_data_row > max_row)
        return;

    const std::string prefix = "firestore-random-past-date-n-year-";

    for (std::uint32_t row = first_data_row; row <= max_row; ++row)
    {
        // Force-create cell so it exists even if empty
        xlnt::cell c = ws.cell(column + std::to_string(row));

        if (fill_with == "firestore-now")
        {
            c.value("__fire_ts_now__");
        }
        else if (str_starts_with(fill_with, prefix))
        {
            std::string years_part = str_slice_from(fill_with, prefix.size());
            std::optional<uint32_t> n_years = std::nullopt;

            try {
                n_years = std::stoul(years_part);
            }
            catch (...) {
                std::cerr << "WARNING: Could not parse N years: " << fill_with << "\n";
            }

            std::string ts = random_past_utc_date_within_n_years(n_years);
            c.value("{ \"__fire_ts_from_date__\": \"" + ts + "\" }");
        }
        else
        {
            c.value(fill_with);
        }
    }

    // Update header
    if (!new_header.empty())
    {
        ws.cell(column + std::to_string(header_row)).value(new_header);
    }
}

void add_column_xlsx(
    xlnt::worksheet &ws,
    const std::uint32_t header_row,
    const std::uint32_t first_data_row,
    const std::string &at,             // either "end" or column letters like "C" or "AA"
    const std::string &fill_with,
    const std::string &new_header
)
{
    // determine used range
    xlnt::range_reference dims = ws.calculate_dimension();
    const std::uint32_t first_row = dims.top_left().row();
    const std::uint32_t last_row  = dims.bottom_right().row();
    const std::uint32_t first_col = dims.top_left().column().index;
    const std::uint32_t last_col  = dims.bottom_right().column().index;

    // figure out insertion index (1-based)
    std::uint32_t insert_at = 0;
    if (at == "end")
    {
        insert_at = last_col + 1; // append after last column
    }
    else if (at == "beginning" || at == "start")
    {
        insert_at = 1; // insert at column A
    }
    else
    {
        insert_at = col_to_index(at);
        if (insert_at < 1) throw std::invalid_argument("Invalid insert column: " + at);
    }

    // if insert_at is beyond last_col+1, we can just create an empty column at that index
    // ensure we move columns only if insert_at <= last_col
    if (insert_at <= last_col)
    {
        // Shift columns right: for each row, move values from last_col..insert_at to +1
        // iterate rows from bottom to top to avoid overwriting
        for (std::uint32_t row = first_row; row <= last_row; ++row)
        {
            // move right-to-left
            for (std::int64_t col = static_cast<std::int64_t>(last_col); col >= static_cast<std::int64_t>(insert_at); --col)
            {
                xlnt::cell src = ws.cell(static_cast<std::uint32_t>(col), row);
                xlnt::cell dst = ws.cell(static_cast<std::uint32_t>(col + 1), row);

                // Copy contents
                if (src.has_formula())
                {
                    dst.formula(src.formula());
                }
                else if (src.data_type() == xlnt::cell_type::number)
                {
                    dst.value(src.value<double>());
                }
                else if (src.data_type() == xlnt::cell_type::boolean)
                {
                    dst.value(src.value<bool>());
                }
                else if (src.data_type() == xlnt::cell_type::inline_string ||
                         src.data_type() == xlnt::cell_type::shared_string ||
                         src.data_type() == xlnt::cell_type::formula_string)
                {
                    dst.value(src.to_string());
                }
                else
                {
                    dst.value(src.to_string()); // fallback for dates, errors, etc.
                }

                // Clear source
                if (src.has_formula())
                    src.formula("");

                src.value("");   // Valid for all xlnt versions

            }
        }
    }
    else
    {
        // insert_at is after last_col: nothing to shift; we just create empty cells if needed
        // nothing needed here; writing later will create cells
    }

    // Column letters for new column
    const std::string new_col_letters = index_to_col(insert_at);

    // Fill the column using your existing helper (which expects column letters)
    fill_column_xlsx(ws, header_row, first_data_row, new_col_letters, fill_with, new_header);
}


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

