#include <iostream>
#include <OpenXLSX.hpp>
#include <CLI/CLI.hpp>
#include "config.hpp"
#include "operations.hpp"
#include "csv.hpp"
#include "json.hpp"
#include "nitro_sheet.hpp"
#include "progress.hpp"

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

    std::string config_path;

    app.add_option("-c, --config", config_path, "Path to YAML config")
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

    Config cfg = load_config(config_path);

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

    std::size_t total_ops = cfg.operations.size();
    std::size_t op_idx = 0;

    // initial 0%
    progress_bar(0, total_ops);

    for (auto &op : cfg.operations)
    {

        if (op.type == "fill-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto fill_with = op.node["fill-with"].as<std::string>();
            auto new_header = op.node["new-header"].as<std::string>("");

            auto col_index = col_to_index(column);

             
            fill_column_nitro(sheet, cfg.header_row, cfg.first_data_row, col_index, fill_with, new_header);
        }
        else if (op.type == "add-column")
        {
            auto at = op.node["at"].as<std::string>();
            auto fill_with = op.node["fill-with"].as<std::string>();
            auto new_header = op.node["new-header"].as<std::string>("");


            add_column_nitro(sheet, cfg.header_row, cfg.first_data_row, at, fill_with, new_header);
        }
        else if (op.type == "split-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>();
            auto targetNodes = op.node["split-to"];
            auto newHeaderNodes = op.node["new-headers"];


            std::vector<std::size_t> targets;

            for (const auto &t : targetNodes)
                targets.push_back(col_to_index(t.as<std::string>()));

            std::vector<std::string> newHeaders;

            for (const auto &h : newHeaderNodes)
                newHeaders.push_back(h.as<std::string>());

            split_column_nitro(sheet, cfg.header_row, cfg.first_data_row, col_to_index(column), delim[0], targets, newHeaders);
        }
        else if (op.type == "uppercase-column")
        {
            auto column = op.node["column"].as<std::string>();

            uppercase_column_nitro(sheet, cfg.first_data_row, col_to_index(column));
        }
        else if (op.type == "replace-in-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto f = op.node["find"].as<std::string>();
            auto r = op.node["replace"].as<std::string>();

            replace_in_column_nitro(sheet, cfg.first_data_row, col_to_index(column), f, r);
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
        }
        else if (op.type == "transform-header")
        {
            auto to = op.node["to"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>("");


            if (!delim.empty())
                transform_header_nitro(sheet, to, delim[0]);
            else
                transform_header_nitro(sheet, to);
        }
        else if (op.type == "rename-header")
        {
            auto column = op.node["column"].as<std::string>();
            auto new_name = op.node["new-name"].as<std::string>("");

            rename_header_nitro(sheet, col_to_index(column), new_name);
        }

        // ---- refresh the progress bar once ----
        op_idx++;
        progress_bar(op_idx, total_ops);
    }

    std::cout << "\n";

    for (auto &op : cfg.operations)
    {

        if (op.type == "fill-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto fill_with = op.node["fill-with"].as<std::string>();
            auto new_header = op.node["new-header"].as<std::string>("");

            std::cout << GREEN "✔ " RESET YELLOW "fill-column" RESET
                      << " (" << CYAN << column << RESET << ") with "
                      << MAGENTA << "\"" << fill_with << "\"" << RESET
                      << " → by header "
                      << GREEN << "\"" << new_header << "\"" << RESET << "\n";
        }
        else if (op.type == "add-column")
        {
            auto at = op.node["at"].as<std::string>();
            auto fill_with = op.node["fill-with"].as<std::string>();
            auto new_header = op.node["new-header"].as<std::string>("");

            std::cout << GREEN "✔ " RESET YELLOW "add-column" RESET
                      << " (" << CYAN << at << RESET << ") with "
                      << MAGENTA << "\"" << fill_with << "\"" << RESET
                      << " → by header "
                      << GREEN << "\"" << new_header << "\"" << RESET << "\n";
        }
        else if (op.type == "split-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>();
            auto targetNodes = op.node["split-to"];
            auto newHeaderNodes = op.node["new-headers"];

            std::cout << GREEN "✔ " RESET YELLOW "split-column" RESET
                      << " (" << CYAN << column << RESET << ")\n";
        }
        else if (op.type == "uppercase-column")
        {
            auto column = op.node["column"].as<std::string>();
            
            std::cout << GREEN "✔ " RESET YELLOW "uppercase-column" RESET
                      << " (" << CYAN << column << RESET << ")\n";
        }
        else if (op.type == "replace-in-column")
        {
            auto column = op.node["column"].as<std::string>();
            auto f = op.node["find"].as<std::string>();
            auto r = op.node["replace"].as<std::string>();

            std::cout << GREEN "✔ " RESET YELLOW "replace-in-column" RESET
                     << " (" << CYAN << column << RESET << ") "
                     << MAGENTA << "\"" << f << "\"" << RESET
                     << " → "
                     << GREEN << "\"" << r << "\"" << RESET << "\n";
        }
        else if (op.type == "transform-row")
        {
            auto row = op.node["row"].as<std::uint32_t>();
            auto to = op.node["to"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>("");

            std::cout << GREEN "✔ " RESET YELLOW "transform-row" RESET
                     << " (" << CYAN << row << RESET << ") "
                     << MAGENTA << "\"" << "to" << "\"" << RESET
                     << " → "
                     << GREEN << "\"" << "with " + delim << "\"" << RESET << "\n";
        }
        else if (op.type == "transform-header")
        {
            auto to = op.node["to"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>("");

            std::cout << GREEN "✔ " RESET YELLOW "transform-header" RESET
                     << " (" << CYAN << "header" << RESET << ") "
                     << MAGENTA << "to" << RESET
                     << " → "
                     << GREEN << "\"" << "with " + delim << "\"" << RESET << "\n";
        }
        else if (op.type == "rename-header")
        {
            auto column = op.node["column"].as<std::string>();
            auto new_name = op.node["new-name"].as<std::string>("");

            std::cout << GREEN "✔ " RESET YELLOW "rename-header" RESET
                     << " (" << CYAN << column << RESET << ") "
                     << MAGENTA << "to" << RESET
                     << " → "
                     << GREEN << "\"" << new_name << "\"" << RESET << "\n";
        }
        else
        {
            std::cerr << RED "✘ Unknown operation type: " << op.type << RESET << "\n";
        }
    }

    std::cout << "\n" << std::flush;


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
