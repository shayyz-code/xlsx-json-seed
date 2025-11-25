// nitro_sheet.hpp - OpenXLSX backed Nitro engine
#pragma once
#include "openxlsx_adapter.hpp"
#include <string>
#include <vector>
#include <thread>
#include <future>
#include <random>
#include <optional>
#include <functional>
#include <algorithm>
#include <cctype>
#include <chrono>
#include <sstream>
#include <fstream>
#include <stdexcept>
#include "utils/utils.hpp"
#include <iostream>
#include <iomanip>

// Column + NitroSheet
struct Column {
    std::string header;
    std::vector<std::string> vals;
    bool dirty = false;
};

struct NitroSheet {
    std::vector<Column> cols;
    uint32_t first_row = 1;
    uint32_t data_row_start = 2;
    uint32_t num_rows = 0;
};

// helper functions (to_snake_single, split_to_parts, random_past_utc_date_within_n_years_opt)
// copy them from the previous nitro implementation exactly (kept concise here)

// For brevity in this message: copy implementations from the previous nitro_sheet.hpp (to_snake_single, split_to_parts, random_past..., apply_unary_to_column, split_column_into_targets, export_csv_buffered, write_back_to_xlsx) but **use the adapter** for write_back_to_xlsx below:

inline NitroSheet load_sheet_vectorized_from_openxlsx(ox::XLWorksheet &ws, uint32_t header_row, uint32_t first_data_row) {
    NitroSheet s;
    SheetDimensions dims = sheet_dimensions(ws);
    uint32_t first_row = dims.first_row;
    uint32_t last_row  = dims.last_row;
    uint32_t first_col = dims.first_col;
    uint32_t last_col  = dims.last_col;

    if (header_row < first_row || header_row > last_row) header_row = first_row;
    if (first_data_row < first_row) first_data_row = header_row + 1;

    s.first_row = first_row;
    s.data_row_start = first_data_row;
    s.num_rows = (first_data_row > last_row) ? 0 : (last_row - first_data_row + 1);

    s.cols.reserve(last_col - first_col + 1);

    for (uint32_t col = first_col; col <= last_col; ++col) {
        Column c;
        c.header = sheet_cell_get(ws, col, header_row);
        c.vals.resize(s.num_rows);
        for (uint32_t r = 0; r < s.num_rows; ++r) {
            c.vals[r] = sheet_cell_get(ws, col, first_data_row + r);
        }
        s.cols.emplace_back(std::move(c));
    }
    return s;
}

// write_back_to_xlsx uses adapter sheet_cell_set
inline void write_back_to_xlsx(
    const NitroSheet &sheet,
    ox::XLWorksheet &ws,
    std::uint32_t header_row,
    std::uint32_t first_data_row)
{
    const size_t num_cols = sheet.cols.size();
    if (num_cols == 0) return;

    const size_t num_rows = sheet.num_rows;

    // ---- Write headers ----
    for (size_t c = 0; c < num_cols; ++c)
    {
        const std::string &header = sheet.cols[c].header;
        std::string cell_ref = index_to_col(c) + std::to_string(header_row);
        ws.cell(cell_ref).value() = header;
    }

    // ---- Write all cell values ----
    for (size_t r = 0; r < num_rows; ++r)
    {
        uint32_t excel_row = first_data_row + r;

        for (size_t c = 0; c < num_cols; ++c)
        {
            const std::string &val = sheet.cols[c].vals[r];
            std::string cell_ref = index_to_col(c) + std::to_string(excel_row);
            ws.cell(cell_ref).value() = val;
        }
    }
}


inline void save_as_xlsx(NitroSheet &sheet, const std::uint32_t header_row, const std::uint32_t first_data_row, const std::string &filename)
{
    OpenXLSX::XLDocument output_wb;
    output_wb.create(filename, /* allow overwrite */ true);   
    try
    {
        auto output_ws = output_wb.workbook().worksheet("Sheet1"); // or create new sheet if not exist

        // Write back NitroSheet data:
        write_back_to_xlsx(sheet, output_ws, header_row, first_data_row);

        output_wb.save();
    }
    catch (const std::exception &e)
    {
        std::cerr << "Failed to save output XLSX file: " << filename
        << "  Error: " << e.what() << "\n";
        throw;
    }
    output_wb.close();
}



// Pretty-print NitroSheet for debugging
inline void print_nitro_sheet(const NitroSheet &sheet, std::ostream &os = std::cout)
{
    os << "NitroSheet Debug Dump\n";
    os << "----------------------------------------\n";
    os << "Columns: " << sheet.cols.size() << "\n";
    os << "Data rows (sheet.num_rows): " << sheet.num_rows << "\n\n";

    if (sheet.cols.empty()) {
        os << "(empty sheet)\n";
        return;
    }

    // Print headers
    os << "Headers:\n";
    for (size_t c = 0; c < sheet.cols.size(); ++c) {
        os << "  [" << c << "] \"" << sheet.cols[c].header << "\"\n";
    }
    os << "\n";

    // Print column sizes
    os << "Column sizes:\n";
    for (size_t c = 0; c < sheet.cols.size(); ++c) {
        os << "  col " << c << ": vals.size() = " << sheet.cols[c].vals.size();
        if (sheet.cols[c].vals.size() != sheet.num_rows)
            os << "  <-- MISMATCH!";
        os << "\n";
    }
    os << "\n";

    // Print row data
    os << "Data:\n";

    for (size_t r = 0; r < sheet.num_rows; ++r) {
        os << "Row " << std::setw(4) << r << ": ";
        for (size_t c = 0; c < sheet.cols.size(); ++c) {
            if (r < sheet.cols[c].vals.size())
                os << "\"" << sheet.cols[c].vals[r] << "\"";
            else
                os << "(OOB!)";

            if (c + 1 < sheet.cols.size()) os << " | ";
        }
        os << "\n";
    }

    os << "----------------------------------------\n";
}
