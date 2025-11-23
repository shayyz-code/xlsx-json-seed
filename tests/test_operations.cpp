#define CATCH_CONFIG_MAIN
#include <catch2/catch_all.hpp>
#include "operations.hpp"
#include "utils.hpp"


TEST_CASE("to_lower converts strings to lowercase", "[to_lower]")
{
    REQUIRE(to_lower("Hello World") == "hello world");
    REQUIRE(to_lower("TESTING") == "testing");
    REQUIRE(to_lower("already lower") == "already lower");
}

TEST_CASE("to_camel converts strings to camelCase", "[to_camel]")
{
    char delims[] = {' ', '-', '_'};
    std::cout << "Hello" + std::to_string(' ') + "World";
    for (auto delim : delims)
    {
        std::string delim_str { delim };

        REQUIRE(to_camel("Hello" + delim_str + "World", delim) == "helloWorld");
        REQUIRE(to_camel("testing" + delim_str + "Testing", delim) == "testingTesting");
        REQUIRE(to_camel("test" + delim_str + "case", delim) == "testCase");
    }
}

TEST_CASE("to_pascal converts strings to PascalCase", "[to_pascal]")
{
    char delims[] = {' ', '-', '_'};
    for (auto delim : delims)
    {
        std::string delim_str { delim };

        REQUIRE(to_pascal("hello" + delim_str + "world", delim) == "HelloWorld");
        REQUIRE(to_pascal("testing" + delim_str + "Testing", delim) == "TestingTesting");
        REQUIRE(to_pascal("test" + delim_str + "case", delim) == "TestCase");
    }
}

TEST_CASE("to_snake converts strings to snake_case", "[to_snake]")
{
    REQUIRE(to_snake("HelloWorld") == "hello_world");
    REQUIRE(to_snake("testingTesting") == "testing_testing");
    REQUIRE(to_snake("TestCase") == "test_case");
    REQUIRE(to_snake("Test-Delim") == "test_delim");
    REQUIRE(to_snake("Test Space") == "test_space");
    REQUIRE(to_snake("DUNS number") == "d_u_n_s_number");
}