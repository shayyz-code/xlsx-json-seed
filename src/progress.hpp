#pragma once
#include <iostream>
#include <string>
#include <cmath>

inline void progress_bar(size_t current, size_t total)
{
    const int bar_width = 50;

    double ratio = (double)current / total;
    int filled = (int)std::round(ratio * bar_width);

    std::cout << "\r[";

    for (int i = 0; i < bar_width; i++)
        std::cout << (i < filled ? "#" : "-");

    std::cout << "] " << (int)(ratio * 100) << "%";
    std::cout.flush();

    if (current == total)
        std::cout << "\n";
}
