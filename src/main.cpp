#include <iostream>
#include <chrono>
#include <OpenXLSX.hpp>
#include <CLI/CLI.hpp>
#include "config.hpp"
#include "operations.hpp"
#include "csv.hpp"
#include "json.hpp"
#include "nitro_sheet.hpp"
#include "progress.hpp"

#define FMT_HEADER_ONLY
#include "fmt/core.h"

// color macros
#define RESET   "\033[0m"
#define BOLD    "\033[1m"

#define RED     "\033[31m"
#define GREEN   "\033[32m"
#define YELLOW  "\033[33m"
#define BLUE    "\033[34m"
#define PURPLE  "\033[35m"
#define CYAN    "\033[36m"
#define WHITE   "\033[37m"
#define MAGENTA "\033[95m"


int main(int argc, char **argv)
{
    CLI::App app { BOLD CYAN "XLSX JSON Seed - A tool to process XLSX files using YAML scripts, primarily for Firestore and other databases seeding" RESET };

    std::string script_path;

    app.add_option("-s, --script", script_path, "Path to YAML script")
        ->required()
        ->check(CLI::ExistingFile);

    CLI11_PARSE(app, argc, argv);


    std::cout << BOLD PURPLE                      
    R"(
██  ██ ██     ▄█████ ██  ██      ██ ▄█████ ▄████▄ ███  ██   ▄█████ ██████ ██████ ████▄  
 ████  ██     ▀▀▀▄▄▄  ████       ██ ▀▀▀▄▄▄ ██  ██ ██ ▀▄██   ▀▀▀▄▄▄ ██▄▄   ██▄▄   ██  ██ 
██  ██ ██████ █████▀ ██  ██   ████▀ █████▀ ▀████▀ ██   ██   █████▀ ██▄▄▄▄ ██▄▄▄▄ ████▀  
    )" "\n" RESET;
    std::cout << BOLD           "by " RESET;
    std::cout << BOLD PURPLE    "shayyz-code\n\n" RESET << std::flush;

    Config cfg = load_script(script_path);

    std::cout << BOLD WHITE "- Input File: " RESET << GREEN << cfg.input_file << RESET << "\n";
    std::cout << BOLD WHITE "- Output File: " RESET << GREEN << cfg.output_file << RESET << "\n";
    std::cout << BOLD WHITE "- Header Row: " RESET << GREEN << cfg.header_row << RESET << "\n";
    std::cout << BOLD WHITE "- First Data Row: " RESET << GREEN << cfg.first_data_row << RESET << "\n\n";
    std::cout << BOLD WHITE "- Export CSV: " RESET << GREEN << (cfg.export_csv ? "yes" : "No") << RESET << "\n";
    std::cout << BOLD WHITE "- Export XLSX (Excel): " RESET << GREEN << (cfg.export_xlsx ? "yes" : "No") << RESET << "\n\n";
    

    // Open with OpenXLSX
    ox::XLDocument wb = open_workbook(cfg.input_file);
    auto ws = worksheet_active(wb);

    auto sheet = load_sheet_vectorized_from_openxlsx(ws, cfg.header_row, cfg.first_data_row);
    std::cout << "# Loaded: cols=" << sheet.cols.size() << " rows=" << sheet.num_rows << "\n\n" << std::flush;

    std::cout << BOLD WHITE "#  Running operations..." RESET << "\n\n" << std::flush;

    std::vector<std::string> logs;



    std::size_t total_ops = cfg.operations.size();

    logs.reserve(total_ops);

    std::size_t op_idx = 0;

    // initial 0%
    progress_bar(0, total_ops);

    auto start = std::chrono::high_resolution_clock::now(); // to measure operation time

    for (auto &op : cfg.operations)
    {
        std::string msg; // message to log

        if (op.type == "fill-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto fill_with = op.node["fill-with"].as<std::string>();
            auto new_header = op.node["new-header"].as<std::string>("");

            auto col_index = col_to_index(column);
             
            fill_column_nitro(sheet, cfg.header_row, cfg.first_data_row, col_index, fill_with, new_header);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "fill-column" RESET
                " (" CYAN "{}" RESET ") with "
                MAGENTA "\"{}\"" RESET
                " → by header "
                GREEN "\"{}\"" RESET,
                column, fill_with, new_header
            );
        }
        else if (op.type == "add-column")
        {
            auto at = op.node["at"].as<std::string>();
            auto fill_with = op.node["fill-with"].as<std::string>();
            auto new_header = op.node["new-header"].as<std::string>("");

            add_column_nitro(sheet, cfg.header_row, cfg.first_data_row, at, fill_with, new_header);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "add-column" RESET
                " (" CYAN "{}" RESET ") with "
                MAGENTA "\"{}\"" RESET
                " → by header "
                GREEN "\"{}\"" RESET,
                at, fill_with, new_header
            );
        }
        else if (op.type == "split-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>();
            auto targetNodes = op.node["split-to"];
            auto newHeaderNodes = op.node["new-headers"];
            auto properPositionNodes = op.node["proper-positions"];


            std::vector<std::size_t> targets;

            for (const auto &t : targetNodes)
                targets.push_back(col_to_index(t.as<std::string>()));

            std::vector<std::string> newHeaders;

            for (const auto &h : newHeaderNodes)
                newHeaders.push_back(h.as<std::string>());

            std::vector<std::uint32_t> properPositions;

            for (const auto &p : properPositionNodes)
                properPositions.push_back(p.as<std::uint32_t>());

            split_column_nitro(sheet, cfg.header_row, cfg.first_data_row, col_to_index(column), delim[0], targets, newHeaders, properPositions);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "split-column" RESET
                " (" CYAN "{}" RESET ")",
                column
            );
        }
        else if (op.type == "uppercase-column")
        {
            auto column = op.node["column"].as<std::string>();

            uppercase_column_nitro(sheet, cfg.first_data_row, col_to_index(column));

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "uppercase-column" RESET
                " (" CYAN "{}" RESET ")",
                column
            );
        }
        else if (op.type == "replace-in-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto f = op.node["find"].as<std::string>();
            auto r = op.node["replace"].as<std::string>();

            replace_in_column_nitro(sheet, cfg.first_data_row, col_to_index(column), f, r);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "replace-in-column" RESET
                " (" CYAN "{}" RESET ") "
                MAGENTA "\"{}\"" RESET
                " → "
                GREEN "\"{}\"" RESET,
                column, f, r
            );
        }
        else if (op.type == "transform-row")
        {
            auto row = op.node["row"].as<std::uint32_t>();
            auto to = op.node["to"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>("");

            auto row_index = row - 1; // convert to 0-based

            if (!delim.empty())
                transform_row_nitro(sheet, row_index, to, delim[0]);
            else
                transform_row_nitro(sheet, row_index, to);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "transform-row" RESET
                " (" CYAN "{}" RESET ") → {}{}",
                row, to,
                delim.empty() ? "" : (" (delim=" + delim + ")")
            );
        }
        else if (op.type == "transform-header")
        {
            auto to = op.node["to"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>("");


            if (!delim.empty())
                transform_header_nitro(sheet, to, delim[0]);
            else
                transform_header_nitro(sheet, to);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "transform-header" RESET
                " → {}{}",
                to,
                delim.empty() ? "" : (" (delim=" + delim + ")")
            );
        }
        else if (op.type == "rename-header")
        {
            auto column = op.node["column"].as<std::string>();
            auto new_name = op.node["new-name"].as<std::string>("");

            rename_header_nitro(sheet, col_to_index(column), new_name);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "rename-header" RESET
                " (" CYAN "{}" RESET ") → {}",
                column, new_name
            );
        }
        else if (op.type == "sort-rows-by-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto ascending = op.node["ascending"].as<bool>(true);

            sort_rows_by_column_nitro(sheet, col_to_index(column), ascending);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "sort-rows-by-column" RESET
                " (" CYAN "{}" RESET ") → {}",
                column, ascending ? "ascending" : "descending"
            );
        }
        else if (op.type == "group-collect")
        {
            auto group_by_column = op.node["group-by"].as<std::string>();
            auto collect_columns = op.node["to-array-columns"];
            auto output_columns = op.node["to-array-output-columns"];
            auto marked_unique = op.node["mark-unique-items"].as<bool>(false);
            auto do_maths_columns = op.node["do-maths-columns"];
            auto do_maths_operations = op.node["do-maths-operations"];

            std::vector<std::size_t> collect_columns_indices;

            for (const auto &t : collect_columns)
                collect_columns_indices.push_back(col_to_index(t.as<std::string>()));

            std::vector<std::size_t> output_columns_indices;

            for (const auto &t : output_columns)
                output_columns_indices.push_back(col_to_index(t.as<std::string>()));

            std::vector<std::size_t> do_maths_columns_indices;

            for (const auto &t : do_maths_columns)
                do_maths_columns_indices.push_back(col_to_index(t.as<std::string>()));

            std::vector<std::string> do_maths_ops;

            for (const auto &t : do_maths_operations)
                do_maths_ops.push_back(t.as<std::string>());

            group_collect_nitro(sheet, col_to_index(group_by_column), collect_columns_indices, output_columns_indices, marked_unique, do_maths_columns_indices, do_maths_ops);

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "group-collect-to" RESET
                " (group=" CYAN "{}" RESET ")",
                group_by_column
            );
        }
        else if (op.type == "reassign-numbering")
        {
            auto column = op.node["column"].as<std::string>();
            auto prefix = op.node["prefix"].as<std::string>();
            auto suffix = op.node["suffix"].as<std::string>();
            auto start_from = op.node["start-from"].as<std::uint32_t>(1);
            auto step = op.node["step"].as<std::uint32_t>(1);

            reassign_numbering_nitro(sheet, col_to_index(column), prefix, suffix, start_from, step );

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "reassign-numbering" RESET
                " (" CYAN "{}" RESET ")",
                column
            );
        }
        else if (op.type == "remove-column")
        {
            auto column = op.node["column"].as<std::string>();

            remove_column_nitro(sheet, col_to_index(column));

            msg = fmt::format(
                GREEN "✔ " RESET YELLOW "remove-column" RESET
                " (" CYAN "{}" RESET ")",
                column
            );
        }
        else {
            op_idx--; // negate the upcoming increment
            
            msg = fmt::format(
                RED "✘ Unknown operation type: {}" RESET,
                op.type
            );
        }

        // ---- refresh the progress bar once ----
        op_idx++;

        // Store the formatted message
        logs.push_back(msg);

        progress_bar(op_idx, total_ops);
    }

    auto end = std::chrono::high_resolution_clock::now(); // to measure operation time
    auto ms = std::chrono::duration<double, std::milli>(end - start).count();

    std::cout << "\n# Computed " << total_ops << " operations in " << ms << " ms\n" << std::flush;

    // ======================================================
    //  PRINT ALL LOGS AFTER PROCESSING
    // ======================================================
    std::cout << "\n";
    for (auto &s : logs)
        std::cout << s << "\n";
    std::cout << std::flush;
    


    save_json_nitro(sheet, cfg.header_row, cfg.first_data_row, cfg.output_file + ".json");

    if (cfg.export_csv)
    {
        save_csv_nitro(sheet, cfg.output_file + ".csv");
    }
    if (cfg.export_xlsx)
    {
        save_as_xlsx(sheet, cfg.header_row, cfg.first_data_row, cfg.output_file + ".xlsx");
    }

    close_workbook(wb);
    
    std::cout << "\n" << BOLD GREEN "✨ Finished seeding!" RESET "\n";
    std::cout << std::flush;

    return 0;
}
