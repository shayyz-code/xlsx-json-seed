#include "config.hpp"

Config load_config(const std::string &path)
{
    YAML::Node root = YAML::LoadFile(path);
    Config cfg;

    cfg.input_file = root["input"].as<std::string>();
    cfg.output_file = root["output"].as<std::string>();
    cfg.export_csv = root["export_csv"].as<bool>(false);

    for (const auto &op : root["operations"])
    {
        Operation o;
        o.type = op["type"].as<std::string>();
        o.node = YAML::Clone(op);
        cfg.operations.push_back(o);
    }

    return cfg;
}