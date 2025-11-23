// openxlsx_adapter.hpp
#pragma once
#include <string>
#include <vector>
#include <optional>
#include <cstdint>
#include <sstream>
#include <OpenXLSX.hpp>

namespace ox = OpenXLSX;

struct SheetDimensions {
    uint32_t first_row;
    uint32_t last_row;
    uint32_t first_col; // 1-based
    uint32_t last_col;  // 1-based
};

inline ox::XLDocument open_workbook(const std::string &path) {
    ox::XLDocument doc;
    doc.open(path);
    return doc;
}

inline void close_workbook(ox::XLDocument &doc) {
    doc.close();
}

// get active sheet (first worksheet)
inline ox::XLWorksheet worksheet_active(ox::XLDocument &doc) {
    return doc.workbook().worksheet(1); // workbook().worksheet(1) is the first sheet in many versions
}

// safe get value from cell (col = 1-based index, row = 1-based index)
// returns empty string if no value / blank cell
inline std::string sheet_cell_get(ox::XLWorksheet &ws, uint32_t col, uint32_t row) {
    try {
        // OpenXLSX allows access via cell reference: e.g. ws.cell("A1")
        // Build a cell reference like "B2"
        std::string ref;
        uint32_t col_n = col;
        // convert col number -> letters
        while (col_n > 0) {
            --col_n;
            char ch = char('A' + (col_n % 26));
            ref.insert(ref.begin(), ch);
            col_n /= 26;
        }
        ref += std::to_string(row);
        auto cell = ws.cell(ref);
        // value() returns XLValue; use get<std::string>() if available
        if (cell.value().type() == ox::XLValueType::Empty) return std::string();
        // Many OpenXLSX versions allow cell.value().get<std::string>()
        try {
            return cell.value().get<std::string>();
        } catch(...) {
            // fallback: to_string via streaming
            std::ostringstream oss;
            oss << cell.value();
            return oss.str();
        }
    } catch(...) {
        return std::string();
    }
}

// set a cell to a string
inline void sheet_cell_set(ox::XLWorksheet &ws, uint32_t col, uint32_t row, const std::string &v) {
    try {
        uint32_t col_n = col;
        std::string ref;
        while (col_n > 0) {
            --col_n;
            char ch = char('A' + (col_n % 26));
            ref.insert(ref.begin(), ch);
            col_n /= 26;
        }
        ref += std::to_string(row);
        auto cell = ws.cell(ref);
        cell.value() = v;
    } catch(...) {
        // best-effort: ignore on failure
    }
}

// Return sheet dimensions using OpenXLSX workbook APIs.
// We attempt to read used range; fallback to simple scan if not supported.
inline SheetDimensions sheet_dimensions(ox::XLWorksheet &ws) {
    // Try to use "dimension" APIs if available. OpenXLSX typically has rowCount/columnCount or dimension info.
    // We'll attempt to query lastRow/lastColumn via 'rowCount' and 'columnCount' or scan until empty.
    SheetDimensions dims{1,1,1,1};
    try {
        // Some versions of OpenXLSX provide usedRange() on the worksheet.
        // We try a few operations with fallbacks.
        // Attempt method: ws.rowCount() and ws.columnCount()
        uint32_t rc = 0, cc = 0;
        try {
            rc = static_cast<uint32_t>(ws.rowCount());
            cc = static_cast<uint32_t>(ws.columnCount());
        } catch(...) {
            rc = 0; cc = 0;
        }

        if (rc > 0 && cc > 0) {
            // assume used area starts at 1
            dims.first_row = 1;
            dims.last_row = rc;
            dims.first_col = 1;
            dims.last_col = cc;
            return dims;
        }

        // Fallback: scan to find last row/col by probing some cells (not super fast but safe)
        // We'll probe a reasonable upper limit e.g., 10000 rows / 1024 cols if rowCount not available.
        const uint32_t MAX_ROWS = 100000; // tuned
        const uint32_t MAX_COLS = 16384;  // Excel max columns
        uint32_t lastR = 1;
        uint32_t lastC = 1;
        // find last column by checking first 100 rows for data across columns (fast heuristic)
        for (uint32_t r = 1; r <= 100 && r <= MAX_ROWS; ++r) {
            for (uint32_t c = 1; c <= 256 && c <= MAX_COLS; ++c) {
                if (!sheet_cell_get(ws, c, r).empty())
                    lastC = std::max(lastC, c);
            }
        }
        // find last row by scanning first few columns
        for (uint32_t c = 1; c <= std::max<uint32_t>(lastC, 1); ++c) {
            for (uint32_t r = 1; r <= MAX_ROWS; ++r) {
                if (!sheet_cell_get(ws, c, r).empty())
                    lastR = std::max(lastR, r);
            }
        }
        dims.first_row = 1;
        dims.first_col = 1;
        dims.last_row = lastR;
        dims.last_col = lastC;
        return dims;
    } catch(...) {
        // ultimate fallback
        dims.first_row = 1;
        dims.first_col = 1;
        dims.last_row = 1;
        dims.last_col = 1;
        return dims;
    }
}
