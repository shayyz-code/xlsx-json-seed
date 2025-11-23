#include <iostream>
#include <xlnt/xlnt.hpp>
#include "config.hpp"
#include "operations.hpp"
#include "csv.hpp"
#include "json.hpp"
#include <CLI/CLI.hpp>

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
    CLI::App app { BOLD CYAN "XLSX JSON Seed - A tool to process XLSX files using YAML scripts, primarily for Firestore Seeding" RESET };

    std::string config_path;

    app.add_option("-c, --config", config_path, "Path to YAML config")
        ->required()
        ->check(CLI::ExistingFile);

    CLI11_PARSE(app, argc, argv);


    std::cout << BOLD BLUE      "──────────────────────────────────────────────────────────────\n" RESET;
    std::cout << BOLD CYAN      "   XLSX-JSON-SEED — Process Excel for Data Seeding via YAML   \n" RESET;
    std::cout << BOLD PURPLE    "                         by shayyz-code                       \n" RESET;
    std::cout << BOLD BLUE      "──────────────────────────────────────────────────────────────\n" RESET;

    Config cfg = load_config(config_path);

    std::cout << BOLD WHITE "- Input File: " RESET << GREEN << cfg.input_file << RESET << "\n";
    std::cout << BOLD WHITE "- Output File: " RESET << GREEN << cfg.output_file << RESET << "\n";
    std::cout << BOLD WHITE "- Header Row: " RESET << GREEN << cfg.header_row << RESET << "\n";
    std::cout << BOLD WHITE "- First Data Row: " RESET << GREEN << cfg.first_data_row << RESET << "\n\n";
    std::cout << BOLD WHITE "- Export CSV: " RESET << GREEN << (cfg.export_csv ? "yes" : "No") << RESET << "\n";
    std::cout << BOLD WHITE "- Export XLSX (Excel): " RESET << GREEN << (cfg.export_xlsx ? "yes" : "No") << RESET << "\n\n";
    std::cout << BOLD WHITE "#  Running operations..." RESET << "\n\n";

    xlnt::workbook wb;
    wb.load(cfg.input_file);
    auto ws = wb.active_sheet();

    for (auto &op : cfg.operations)
    {
        if (op.type == "fill-column")
        {
            auto col = op.node["column"].as<std::string>();
            auto fill_with = op.node["fill-with"].as<std::string>();
            auto new_header = op.node["new-header"].as<std::string>("");

             std::cout << GREEN "✔ " RESET YELLOW "fill-column" RESET
                      << " (" << CYAN << col << RESET << ") with "
                      << MAGENTA << "\"" << fill_with << "\"" << RESET
                      << " → by header "
                      << GREEN << "\"" << new_header << "\"" << RESET << "\n";

            fill_column_xlsx(ws, cfg.header_row, cfg.first_data_row, col, fill_with, new_header);
        }
        else if (op.type == "split-column")
        {
            auto src = op.node["source"].as<std::string>();
            auto delim = op.node["delimiter"].as<std::string>();
            auto targetNodes = op.node["split-to"];
            auto newHeaderNodes = op.node["new-headers"];

            std::vector<std::string> targets;

            for (const auto &t : targetNodes)
                targets.push_back(t.as<std::string>());

            std::vector<std::string> newHeaders;

            for (const auto &h : newHeaderNodes)
                newHeaders.push_back(h.as<std::string>());

            std::cout << GREEN "✔ " RESET YELLOW "split-column" RESET
                      << " (" << CYAN << src << RESET << ")\n";

            split_column_xlsx(ws, cfg.header_row, cfg.first_data_row, src, delim, targets, newHeaders);
        }
        else if (op.type == "uppercase-column")
        {
            auto col = op.node["column"].as<std::string>();

            std::cout << GREEN "✔ " RESET YELLOW "uppercase-column" RESET
                      << " (" << CYAN << col << RESET << ")\n";
            uppercase_column_xlsx(ws, cfg.first_data_row, col);
        }
        else if (op.type == "replace-in-column")
        {
            auto col = op.node["column"].as<std::string>();
            auto f = op.node["find"].as<std::string>();
            auto r = op.node["replace"].as<std::string>();

             std::cout << GREEN "✔ " RESET YELLOW "replace-in-column" RESET
                      << " (" << CYAN << col << RESET << ") "
                      << MAGENTA << "\"" << f << "\"" << RESET
                      << " → "
                      << GREEN << "\"" << r << "\"" << RESET << "\n";
            replace_in_column_xlsx(ws, cfg.first_data_row, col, f, r);
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

            if (!delim.empty())
                transform_row_xlsx(ws, row, to, delim[0]);
            else
                transform_row_xlsx(ws, row, to);
        }
        else
        {
            std::cerr << RED "✘ Unknown operation type: " << op.type << RESET << "\n";
        }
    }


    save_json(ws, cfg.header_row, cfg.first_data_row, cfg.output_file + ".json");

    if (cfg.export_csv)
    {
        save_csv(ws, cfg.header_row, cfg.first_data_row, cfg.output_file + ".csv");
    }
    if (cfg.export_xlsx)
    {
        wb.save(cfg.output_file + ".xlsx");
    }
    
    std::cout << "\n" << BOLD GREEN "✨ Finished seeding!" RESET "\n";
    std::cout << BOLD BLUE "──────────────────────────────────────────────────────────────\n" RESET;
}
