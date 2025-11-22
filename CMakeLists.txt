cmake_minimum_required(VERSION 3.21)

project(xlsx_seed LANGUAGES CXX)

set(CMAKE_CXX_COMPILER "clang++")

# Set the C++ standard to C++17
set(CMAKE_CXX_STANDARD 17)

# Ensure the compiler *must* support C++17, otherwise CMake will error
set(CMAKE_CXX_STANDARD_REQUIRED ON)

enable_testing()

find_package(xlnt CONFIG REQUIRED)
find_package(yaml-cpp CONFIG REQUIRED)
find_package(CLI11 CONFIG REQUIRED)

add_library(xlsx_seed_lib
    src/config.cpp
    src/operations.cpp
    src/csv.hpp
)

target_include_directories(xlsx_seed_lib PUBLIC src)

target_link_libraries(xlsx_seed_lib
    PRIVATE
        yaml-cpp
        xlnt::xlnt
)

add_executable(xlsx_seed
    src/main.cpp
)

target_link_libraries(xlsx_seed 
    PRIVATE
        xlsx_seed_lib
        xlnt::xlnt
        CLI11::CLI11
)

# ---- Tests ----
find_package(Catch2 CONFIG REQUIRED)

add_executable(test_operations tests/test_operations.cpp)
target_link_libraries(test_operations PRIVATE xlsx_seed_lib Catch2::Catch2WithMain)

add_test(NAME operations_test COMMAND test_operations)
