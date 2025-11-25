cmake_minimum_required(VERSION 3.21)

project(xlsx_json_seed LANGUAGES CXX)

set(CMAKE_CXX_COMPILER "clang++")

# Set the C++ standard to C++17
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_EXTENSIONS OFF)

# Ensure the compiler *must* support C++17, otherwise CMake will error
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_compile_options(-O3 -march=native -DNDEBU)

enable_testing()

find_package(OpenXLSX CONFIG REQUIRED)
find_package(yaml-cpp CONFIG REQUIRED)
find_package(CLI11 CONFIG REQUIRED)
find_package(fmt CONFIG REQUIRED)

add_library(xlsx_json_seed_lib
    src/openxlsx_adapter.hpp
    src/nitro_sheet.hpp
    src/config.cpp
    src/operations.cpp
    src/utils.cpp
    src/csv.hpp
    src/json.hpp
    src/progress.hpp
)

target_include_directories(xlsx_json_seed_lib PUBLIC src)

target_link_libraries(xlsx_json_seed_lib
    PRIVATE
        OpenXLSX::OpenXLSX
        yaml-cpp
)

add_executable(xlsx_json_seed
    src/main.cpp
)

target_link_libraries(xlsx_json_seed 
    PRIVATE
        xlsx_json_seed_lib
        OpenXLSX::OpenXLSX
        CLI11::CLI11
        fmt::fmt
)

# ---- Tests ----
find_package(Catch2 CONFIG REQUIRED)

add_executable(test_operations
    tests/test_operations.cpp
)
target_link_libraries(test_operations PRIVATE xlsx_json_seed_lib OpenXLSX::OpenXLSX Catch2::Catch2WithMain)
add_test(NAME operations_test COMMAND test_operations)

add_executable(test_utils
    tests/test_utils.cpp
)
target_link_libraries(test_utils PRIVATE xlsx_json_seed_lib Catch2::Catch2WithMain)
add_test(NAME utils_test COMMAND test_utils)


