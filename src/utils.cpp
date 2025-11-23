#include <random>
#include <chrono>
#include "utils.hpp"
#include <iostream>
#include <string>

#define RESET   "\033[0m"

#ifdef _WIN32
#include <windows.h>
#else
#include <sys/ioctl.h>
#include <unistd.h>
#endif

// helpers

std::string str_slice_from(const std::string &s, size_t start)
{
    if (start > s.size()) return "";   // avoid out-of-range
    return s.substr(start);
}

bool str_starts_with(const std::string &s, const std::string &prefix)
{
    if (s.size() < prefix.size()) return false;
    return std::equal(prefix.begin(), prefix.end(), s.begin());
}

// -----------------------------
// Column name <-> index helpers
// -----------------------------
// "A" -> 1, "Z" -> 26, "AA" -> 27
std::size_t col_to_index(const std::string &letters)
{
    size_t index = 0;
    for (char c : letters)
    {
        if (c >= 'A' && c <= 'Z') index = index * 26 + (c - 'A' + 1);
        else if (c >= 'a' && c <= 'z') index = index * 26 + (c - 'a' + 1);
        else throw std::invalid_argument("Invalid column letter: " + letters);
    }
    return index - 1; // convert 1-based -> 0-based
}

// ----------------------
// Convert 0-based index to column letters
// ----------------------
std::string index_to_col(size_t index)
{
    std::string letters;
    index += 1; // make 1-based
    while (index > 0)
    {
        size_t rem = (index - 1) % 26;
        letters = static_cast<char>('A' + rem) + letters;
        index = (index - 1) / 26;
    }
    return letters;
}

std::string to_upper(const std::string &str)
{
    std::string r = str;
    std::transform(r.begin(), r.end(), r.begin(), [](unsigned char c) { return std::toupper(c);});

    return r;
}

std::string to_lower(const std::string &str)
{
    std::string r = str;
    std::transform(r.begin(), r.end(), r.begin(), [](unsigned char c) { return std::tolower(c);});

    return r;
}

std::string to_camel(const std::string &str, char &delim)
{
    std::string r = to_lower(str);
    bool cap_next = false;

    for (size_t i = 0; i < r.size(); i++)
    {
        if (r[i] == delim)
        {
            cap_next = true;
        }
        else if (cap_next)
        {
            r[i] = std::toupper(r[i]);
            cap_next = false;
        }
    }

    r.erase(std::remove(r.begin(), r.end(), delim), r.end());

    return r;
}


std::string to_pascal(const std::string &str, char &delim)
{
    std::string r = to_camel(str, delim);
    if (!r.empty())
    {
        r[0] = std::toupper(r[0]);
    }
    return r;
}

// helper for to_snake
static inline bool is_sep(char c)
{
    return std::isspace((unsigned char)c) || std::ispunct((unsigned char)c);
}

std::string to_snake(const std::string &str)
{
    std::string out;
    out.reserve(str.size() * 2);

    bool last_was_sep = false;

    for (size_t i = 0; i < str.size(); ++i)
    {
        unsigned char c = str[i];

        // ---- Separator / punctuation / space ----
        if (is_sep(c))
        {
            if (!out.empty() && !last_was_sep)
            {
                out += '_';
                last_was_sep = true;
            }
            continue;
        }

        // ---- Uppercase letter and not first char ----
        if (std::isupper(c))
        {
            if (!out.empty() && !last_was_sep)
            {
                out += '_';
            }
            out += (char)std::tolower(c);
            last_was_sep = false;
            continue;
        }

        // ---- Lowercase or digit ----
        out += (char)std::tolower(c);
        last_was_sep = false;
    }

    // Trim trailing underscore
    while (!out.empty() && out.back() == '_')
        out.pop_back();

    return out;
}

// json-firestore-seed

std::string random_past_utc_date_within_n_years(
    std::optional<uint32_t> n_years
)
{

    // Current UTC time point
    std::chrono::system_clock::time_point now = std::chrono::system_clock::now();

    // One year in seconds (365 days)
    const int64_t one_year_seconds = 365LL * 24 * 60 * 60 * *n_years;

    // Random generator
    static thread_local std::mt19937_64 rng(std::random_device{}());
    std::uniform_int_distribution<int64_t> dist(0, one_year_seconds);

    // Pick a random offset backwards from "now"
    int64_t random_offset = dist(rng);

    std::chrono::system_clock::time_point random_time = now - std::chrono::seconds(random_offset);

    // Convert to time_t (UTC)
    std::time_t tt = std::chrono::system_clock::to_time_t(random_time);

    // Format as UTC ISO 8601
    std::tm tm_utc;
    #if defined(_WIN32)
        gmtime_s(&tm_utc, &tt);
    #else
        gmtime_r(&tt, &tm_utc);
    #endif

    char buffer[32];
    std::strftime(buffer, sizeof(buffer), "%Y-%m-%dT%H:%M:%SZ", &tm_utc);

    return std::string(buffer);
}

// os terminal helpers
int get_terminal_width()
{
#ifdef _WIN32
    CONSOLE_SCREEN_BUFFER_INFO csbi;
    if (GetConsoleScreenBufferInfo(GetStdHandle(STD_OUTPUT_HANDLE), &csbi))
        return csbi.srWindow.Right - csbi.srWindow.Left + 1;
    else
        return 80; // fallback
#else
    struct winsize ws;
    if (ioctl(STDOUT_FILENO, TIOCGWINSZ, &ws) == 0)
        return ws.ws_col;
    else
        return 80; // fallback
#endif
}

void print_full_line_utf8(const std::string &color, const std::string &glyph)
{
    int width = get_terminal_width();
    if (width < 1) width = 80;

    std::cout << color;

    // Print glyph 'width' times
    for (int i = 0; i < width; ++i)
        std::cout << glyph;

    std::cout << RESET << "\n";
}
