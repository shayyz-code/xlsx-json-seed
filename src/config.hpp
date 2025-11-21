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
    std::vector<Operation> operations;
};

Config load_config(const std::string &path);