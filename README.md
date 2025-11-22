# xlsx-json-seed

![c++](https://img.shields.io/badge/C%2B%2B-00599C?style=for-the-badge&logo=c%2B%2B&logoColor=white)
![cmake](https://img.shields.io/badge/CMake-064F8C?style=for-the-badge&logo=cmake&logoColor=white)
![MIT](https://img.shields.io/badge/MIT-green?style=for-the-badge)

A Blazingly Fast and Flexible **CLI tool for processing and automating Excel (.xlsx) data for Firestore Seeding** using a simple **YAML instruction file**.

Can be used together with [json-firestore-seed](https://github.com/shayyz-code/json-firestore-seed) to process firestore data entry.

Built in **C++**, powered by **xlnt**, and **yaml-cpp**.

## Features

- Read `.xlsx` Excel files
- Export to **JSON**, **CSV** or back to **XLSX**
- Automate transformations via **YAML scripting**
- Split, uppercase, replace, and more
- Colorized CLI
- Blazingly Fast native runtime
- Easy to extend with new operations

## Installation

### 1. Clone the repo

```bash
git clone https://github.com/shayyz-code/xlsx-json-seed
cd xlsx-json-seed
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
build/xlsx_json_seed
```

## Usage

```
./build/xlsx_json_seed --config config.yaml
```

## Example

config.yaml and input.xlsx can be found in [./example](./example).

```yaml
input: "example/input.xlsx"
output: "example/result"
header-row: 1
first-data-row: 2

export-csv: true

operations:
  - type: split-column
    source: B
    delimiter: "-"
    split-to: [D, E]
    new-headers: ["Part1", "Part2"]

  # - type: uppercase-column
  #   column: C

  - type: replace-in-column
    column: C
    find: "-"
    replace: ","

  - type: transform-row
    row: 1
    to: "snake_case"
    delimiter: " "
```

## Result

JSON:

```json
[
  {
    "no": "1",
    "product_code": "VG-WHITE",
    "product_name": "V Shirt",
    "part1": "VG",
    "part2": "WHITE"
  },
  {
    "no": "2",
    "product_code": "GH-BLUE",
    "product_name": "G Handbag",
    "part1": "GH",
    "part2": "BLUE"
  },
  {
    "no": "3",
    "product_code": "BN-PURPLE",
    "product_name": "B Necklace",
    "part1": "BN",
    "part2": "PURPLE"
  }
]
```

CSV:

```csv
no,product_code,product_name,part1,part2
1,VG-WHITE,V Shirt,VG,WHITE
2,GH-BLUE,G Handbag,GH,BLUE
3,BN-PURPLE,B Necklace,BN,PURPLE

```

## How to Contribute

Pull requests and feature suggestions are welcome!
Feel free to open an issue to discuss ideas.

## Contributors

Thanks goes to these wonderful people ✨

![Contributors](https://contrib.rocks/image?repo=shayyz-code/xlsx-json-seed)

## License

MIT License — free for personal & commercial use.

Copyright (c) 2025 shayyz-code.
