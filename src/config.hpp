#pragma once
#include <string>
#include <vector>
#include <yaml-cpp/yaml.h>

struct Operation
{
    std::string type;
    YAML::Node node;
};

struct Config
{
    std::string input_file;
    std::string output_file;
    bool export_csv = false;
    bool export_xlsx = false;
    std::uint32_t header_row = 1;
    std::uint32_t first_data_row = 2;
    std::vector<Operation> operations;
};

Config load_script(const std::string &path);