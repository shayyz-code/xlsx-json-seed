#include "dynamic_placeholder.hpp"


std::vector<PlaceholderSpan> scan_placeholders(const std::string &s) {
    std::vector<PlaceholderSpan> spans;
    spans.reserve(16); // small optimization

    std::size_t pos = 0;
    while (true) {
        // find start of ${...}
        std::size_t start = s.find("${", pos);
        if (start == std::string::npos)
            break;

        // find the closing brace
        std::size_t end = s.find('}', start + 2);
        if (end == std::string::npos)
            break; // malformed — ignore / stop

        // extract key inside ${ ... }
        std::string key = s.substr(start + 2, end - (start + 2));

        if (key.empty()) {
            pos = end + 1; // continue scanning after '}'
            continue;      // skip empty keys
        }

        spans.push_back({ start, end, std::move(key) });

        pos = end + 1;  // continue scanning after '}'
    }

    return spans;
}