#include "config.hpp"

Config load_script(const std::string &path)
{
    YAML::Node root = YAML::LoadFile(path);
    Config cfg;

    cfg.input_file = root["input"].as<std::string>();
    cfg.output_file = root["output"].as<std::string>();
    cfg.export_csv = root["export-csv"].as<bool>(false);
    cfg.export_xlsx = root["export-xlsx"].as<bool>(false);
    cfg.header_row = root["header-row"].as<std::uint32_t>(1);
    cfg.first_data_row = root["first-data-row"].as<std::uint32_t>(2);

    for (const auto &op : root["operations"])
    {
        Operation o;
        o.type = op["type"].as<std::string>();
        o.node = YAML::Clone(op);
        cfg.operations.push_back(o);
    }

    return cfg;
}