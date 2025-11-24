```
██  ██ ██     ▄█████ ██  ██      ██ ▄█████ ▄████▄ ███  ██   ▄█████ ██████ ██████ ████▄
 ████  ██     ▀▀▀▄▄▄  ████       ██ ▀▀▀▄▄▄ ██  ██ ██ ▀▄██   ▀▀▀▄▄▄ ██▄▄   ██▄▄   ██  ██
██  ██ ██████ █████▀ ██  ██   ████▀ █████▀ ▀████▀ ██   ██   █████▀ ██▄▄▄▄ ██▄▄▄▄ ████▀
```

![c++](https://img.shields.io/badge/C%2B%2B-00599C?style=for-the-badge&logo=c%2B%2B&logoColor=white)
![cmake](https://img.shields.io/badge/CMake-064F8C?style=for-the-badge&logo=cmake&logoColor=white)
![MIT](https://img.shields.io/badge/MIT-green?style=for-the-badge)

# XLSX JSON Seed

A **Blazingly Fast, Nitro-Boosted** tool for processing and automating Excel (.xlsx) data for any databases using a simple **YAML instruction file**.

Can be used together with [**json-firestore-seed**](https://github.com/shayyz-code/json-firestore-seed) to process **firestore data entry**.

Built in **C++**, powered by [**OpenXLSX**](https://github.com/troldal/OpenXLSX), and [**yaml-cpp**](https://github.com/jbeder/yaml-cpp).

## NOTE: Compactibility, Performance, and Memory Usage

See at OpenXLSX's repo [https://github.com/troldal/OpenXLSX](https://github.com/troldal/OpenXLSX) (since it may be updated at anytime)

_OpenXLSX version used in this project is (aral-matrix) 14 July 2025._

## Features

- Supports Firestore timestamps (works with [json-firestore-seed](https://github.com/shayyz-code/json-firestore-seed))
- Read `.xlsx` Excel files your client (customer) gives and convert to json for database schema
- Export to **JSON**, **CSV** or back to **XLSX**
- Automate via **YAML scripting**
- Add, fill, split, uppercase, replace, and more
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
vcpkg install openxlsx yaml-cpp
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

_config.yaml_ and _input.xlsx_ can be found in [./example](./example).

```yaml
input: "example/input.xlsx"
output: "example/result"
header-row: 1
first-data-row: 2

export-csv: true

operations:
  - type: split-column
    column: B
    delimiter: "-"
    split-to: [D, E]
    new-headers: ["Part1", "Part2"]

  # - type: uppercase-column
  #   column: G

  - type: replace-in-column
    column: C
    find: "-"
    replace: ","

  - type: fill-column
    column: F
    fill-with: "firestore-random-past-date-n-year-2"
    new-header: "Created At"

  - type: add-column
    at: "end"
    fill-with: "firestore-now"
    new-header: "Updated At"

  - type: add-column
    at: "start"
    fill-with: "prefix"
    new-header: "Prefix"

  # - type: transform-row
  #   row: 1
  #   to: "camelCase"
  #   delimiter: " "

  - type: rename-header
    column: A
    new-name: "New Prefix"

  - type: transform-header
    to: "snake_case"
    delimiter: " "
```

## Result

JSON:

```json
[
  {
    "new_prefix": "prefix",
    "no": "1",
    "product_code": "VG-WHITE",
    "product_name": "V Shirt",
    "part1": "VG",
    "part2": "WHITE",
    "created_at": { "__fire_ts_from_date__": "2025-08-19T16:46:05Z" },
    "updated_at": "__fire_ts_now__"
  },
  {
    "new_prefix": "prefix",
    "no": "2",
    "product_code": "GH-BLUE",
    "product_name": "G Handbag",
    "part1": "GH",
    "part2": "BLUE",
    "created_at": { "__fire_ts_from_date__": "2024-12-10T22:55:11Z" },
    "updated_at": "__fire_ts_now__"
  },
  {
    "new_prefix": "prefix",
    "no": "3",
    "product_code": "BN-PURPLE",
    "product_name": "B Necklace",
    "part1": "BN",
    "part2": "PURPLE",
    "created_at": { "__fire_ts_from_date__": "2024-02-07T22:53:14Z" },
    "updated_at": "__fire_ts_now__"
  }
]
```

CSV:

```csv
new_prefix,no,product_code,product_name,part1,part2,created_at,updated_at
prefix,1,VG-WHITE,V Shirt,VG,WHITE,"{ ""__fire_ts_from_date__"": ""2025-08-19T16:46:05Z"" }",__fire_ts_now__
prefix,2,GH-BLUE,G Handbag,GH,BLUE,"{ ""__fire_ts_from_date__"": ""2024-12-10T22:55:11Z"" }",__fire_ts_now__
prefix,3,BN-PURPLE,B Necklace,BN,PURPLE,"{ ""__fire_ts_from_date__"": ""2024-02-07T22:53:14Z"" }",__fire_ts_now__
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
