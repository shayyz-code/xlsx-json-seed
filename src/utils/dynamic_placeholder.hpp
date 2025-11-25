#include <string>
#include <vector>


struct PlaceholderSpan {
    std::size_t start;
    std::size_t end;
    std::string key;
};

std::vector<PlaceholderSpan> scan_placeholders(const std::string &s);