cmake_minimum_required(VERSION 3.21)

include(~/vcpkg/scripts/buildsystems/vcpkg.cmake)

project(xlsx_seed)

add_executable(xlsx_seed
    src/main.cpp
    src/config.cpp
    src/config.hpp
    src/operations.cpp
    src/operations.hpp
    src/csv.hpp
)

find_package(xlnt CONFIG REQUIRED)
find_package(yaml-cpp CONFIG REQUIRED)
find_package(CLI11 CONFIG REQUIRED)

target_link_libraries(xlsx_seed 
    PRIVATE
        yaml-cpp
        xlnt::xlnt
        CLI11::CLI11
)
