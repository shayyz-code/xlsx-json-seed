#define CATCH_CONFIG_MAIN
#include <catch2/catch_all.hpp>
#include "utils/utils.hpp"


TEST_CASE("str_slice_from returns correct substring", "[str_slice_from]")
{
    REQUIRE(str_slice_from("Hello, World!", 7) == "World!");
    REQUIRE(str_slice_from("Hello, World!", 0) == "Hello, World!");
    REQUIRE(str_slice_from("Hello, World!", 13) == "");
    REQUIRE(str_slice_from("Hello, World!", 20) == "");
}

TEST_CASE("str_starts_with detects prefixes correctly", "[str_starts_with]")
{
    REQUIRE(str_starts_with("Hello, World!", "Hello") == true);
    REQUIRE(str_starts_with("Hello, World!", "World") == false);
    REQUIRE(str_starts_with("Hello, World!", "") == true);
    REQUIRE(str_starts_with("Short", "LongerPrefix") == false);
}