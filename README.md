# xlsx-seed

<p align="left">
  <img src="https://img.shields.io/badge/language-C++20-blue.svg" />
  <img src="https://img.shields.io/badge/build-CMake-informational.svg" />
  <img src="https://img.shields.io/badge/platform-macOS%20%7C%20Linux%20%7C%20Windows-lightgrey.svg" />
  <img src="https://img.shields.io/badge/license-MIT-green.svg" />
  <img src="https://img.shields.io/badge/status-active-brightgreen.svg" />
  <img src="https://img.shields.io/github/last-commit/shayyz-code/xlsx-seed" />
</p>

A Blazingly Fast and Flexible **CLI tool for processing and automating Excel (.xlsx) data for Database Seeding** using a simple **YAML instruction file**.

Built in **C++**, powered by **xlnt**, and **yaml-cpp**.

## Features

- Read `.xlsx` Excel files
- Export to **XLSX** or **CSV**
- Automate transformations via **YAML scripting**
- Split, uppercase, replace, and more
- Colorized CLI
- Blazingly Fast native runtime
- Easy to extend with new operations

## Installation

### 1. Clone the repo

```bash
git clone https://github.com/shayyz-code/xlsx-seed
cd xlsx-seed
```

### 2. Install dependencies via vcpkg (not with manifest mode)

```bash
vcpkg install xlnt yaml-cpp
```

Or with manifest mode:

```bash
vcpkg install
```

### 3. Build the tool

```bash
cmake -B build -DCMAKE_TOOLCHAIN_FILE=path/to/vcpkg.cmake
cmake --build build --config Release
```

CLI binary will be at:

```
build/xlsx_seed
```

## Usage

```
./build/xlsx_seed --config config.yaml
```

## Example

config.yaml and input.xlsx can be found in [./example](./example).

```yaml
input: "example/input.xlsx"
output: "example/result"
export_csv: true

operations:
  - type: split-column
    source: B
    delimiter: "-"
    targets: [D, E]

  - type: uppercase-column
    column: C

  - type: replace-in-column
    column: C
    find: " "
    replace: " - "
```

## How to Contribute

Pull requests and feature suggestions are welcome!
Feel free to open an issue to discuss ideas.

## Contributors

Thanks goes to these wonderful people ✨

![Contributors](https://contrib.rocks/image?repo=shayyz-code/xlsx-seed)

## License

MIT License — free for personal & commercial use.

Copyright (c) 2025 shayyz-code.
