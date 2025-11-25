```
██  ██ ██     ▄█████ ██  ██      ██ ▄█████ ▄████▄ ███  ██   ▄█████ ██████ ██████ ████▄
 ████  ██     ▀▀▀▄▄▄  ████       ██ ▀▀▀▄▄▄ ██  ██ ██ ▀▄██   ▀▀▀▄▄▄ ██▄▄   ██▄▄   ██  ██
██  ██ ██████ █████▀ ██  ██   ████▀ █████▀ ▀████▀ ██   ██   █████▀ ██▄▄▄▄ ██▄▄▄▄ ████▀
```

![c++](https://img.shields.io/badge/C%2B%2B-00599C?style=for-the-badge&logo=c%2B%2B&logoColor=white)
![cmake](https://img.shields.io/badge/CMake-064F8C?style=for-the-badge&logo=cmake&logoColor=white)
![MIT](https://img.shields.io/badge/MIT-green?style=for-the-badge)

# XLSX JSON Seed

A **Blazingly Fast, Nitro-Boosted** tool for processing and automating Excel (.xlsx) data for any databases (including Firestore) using a simple **YAML instruction file**.

Can be used together with [**json-firestore-seed**](https://github.com/shayyz-code/json-firestore-seed) to process **firestore data entry**.

Built in **C++**, powered by [**OpenXLSX**](https://github.com/troldal/OpenXLSX), and [**yaml-cpp**](https://github.com/jbeder/yaml-cpp).

## NOTE: Compactibility, Performance, and Memory Usage

See at OpenXLSX's repo [https://github.com/troldal/OpenXLSX](https://github.com/troldal/OpenXLSX) (since it may be updated at anytime)

_OpenXLSX version used in this project is (aral-matrix) 14 July 2025._

# Table of Contents

- [XLSX JSON Seed](#xlsx-json-seed)
  - [Features](#features)
  - [Installation](#installation)
    - [1. Clone the repo](#1-clone-the-repo)
    - [2. Install dependencies via vcpkg](#2-install-dependencies-via-vcpkg-not-with-manifest-mode)
    - [3. Build the tool](#3-build-the-tool)
  - [Usage](#usage)
  - [Example](#example)
  - [Result](#result)
    - [JSON](#json)
    - [CSV](#csv)
  - [Operations Reference Table](#operations-reference-table)
  - [How to Contribute](#how-to-contribute)
  - [Contributors](#contributors)
  - [License](#license)

## Features

- Supports Firestore timestamps (works with [json-firestore-seed](https://github.com/shayyz-code/json-firestore-seed))
- Read `.xlsx` Excel files your client (customer) gives and convert to json for database schema
- Export to **JSON**, **CSV** or back to **XLSX**
- Automate via **YAML scripting**
- Add, remove, group-as-array, renumbering, fill, split, uppercase, replace, and more
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
    split-to: [E, F]
    new-headers: ["Part1", "Part2"]

  # - type: uppercase-column
  #   column: G

  - type: replace-in-column
    column: C
    find: "-"
    replace: ","

  - type: fill-column
    column: G
    fill-with: "firestore-random-past-date-n-year-2"
    new-header: "Created At"

  - type: add-column
    at: "end"
    fill-with: "firestore-now"
    new-header: "Updated At"

  - type: add-column
    at: "start"
    fill-with: column[F] # fill with values from column F
    new-header: "Dynamic Value"

  - type: sort-rows-by-column
    column: F
    ascending: true

  - type: group-collect-to-array
    group-by: F # column with BN (type column)
    collect-column: G # colors
    output-column: G # write array back to color column

  - type: reassign-numbering
    column: B
    prefix: ""
    suffix: "."
    start-from: 1
    step: 1

  - type: remove-column
    column: C

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
    "new_prefix": "BN",
    "no": "1.",
    "product_name": "B Necklace",
    "price": 1000,
    "part1": "BN",
    "part2": ["PURPLE", "RED"],
    "created_at": { "__fire_ts_from_date__": "2023-12-21T12:58:37Z" },
    "updated_at": "__fire_ts_now__"
  },
  {
    "new_prefix": "GH",
    "no": "2.",
    "product_name": "G Handbag",
    "price": 200,
    "part1": "GH",
    "part2": ["BLUE", "BROWN"],
    "created_at": { "__fire_ts_from_date__": "2024-02-27T19:27:53Z" },
    "updated_at": "__fire_ts_now__"
  },
  {
    "new_prefix": "VG",
    "no": "3.",
    "product_name": "V Shirt",
    "price": 60,
    "part1": "VG",
    "part2": ["WHITE"],
    "created_at": { "__fire_ts_from_date__": "2024-01-29T10:42:42Z" },
    "updated_at": "__fire_ts_now__"
  }
]
```

CSV:

```csv
new_prefix,no,product_name,price,part1,part2,created_at,updated_at
BN,1.,B Necklace,1000,BN,"[""PURPLE"",""RED""]","{ ""__fire_ts_from_date__"": ""2023-12-21T12:58:37Z"" }",__fire_ts_now__
GH,2.,G Handbag,200,GH,"[""BLUE"",""BROWN""]","{ ""__fire_ts_from_date__"": ""2024-02-27T19:27:53Z"" }",__fire_ts_now__
VG,3.,V Shirt,60,VG,"[""WHITE""]","{ ""__fire_ts_from_date__"": ""2024-01-29T10:42:42Z"" }",__fire_ts_now__
```

## Operations Reference Table

| **Operation Type**       | **Description**                                                                    | **Required Fields**                                 | **Optional Fields**                  |
| ------------------------ | ---------------------------------------------------------------------------------- | --------------------------------------------------- | ------------------------------------ |
| `split-column`           | Splits a column into multiple parts by a delimiter.                                | `column`, `delimiter`, `split-to`, `new-headers`    | —                                    |
| `replace-in-column`      | Replaces occurrences of a substring within a column.                               | `column`, `find`, `replace`                         | —                                    |
| `fill-column`            | Fills a column with a constant or dyanmic value and optionally renames the header. | `column`, `fill-with` <br />// Dynamic -> column[F] | `new-header`                         |
| `add-column`             | Adds a column at the start, end, before, or after another column.                  | `at`, `fill-with`, `new-header`                     | —                                    |
| `uppercase-column`       | Converts the entire column to uppercase.                                           | `column`                                            | —                                    |
| `sort-rows-by-column`    | Sorts rows by a given column (ascending/descending).                               | `column`                                            | `ascending` (default `true`)         |
| `group-collect-to-array` | Groups rows by one column and creates an array of values from another column.      | `group-by`, `collect-column`, `output-column`       | —                                    |
| `reassign-numbering`     | Replaces a numeric column with a new sequence number format.                       | `column`, `prefix`, `suffix`                        | `start-from` (default 1), `step` (1) |
| `remove-column`          | Deletes a column entirely.                                                         | `column`                                            | —                                    |
| `rename-header`          | Renames a column header.                                                           | `column`, `new-name`                                | —                                    |
| `transform-row`          | Transforms one row into another format (camelCase, snake_case, etc.).              | `row`, `to`                                         | `delimiter`                          |
| `transform-header`       | Transforms all headers (camelCase, snake_case, etc.).                              | `to`                                                | `delimiter`                          |

## How to Contribute

Pull requests and feature suggestions are welcome!
Feel free to open an issue to discuss ideas.

## Contributors

Thanks goes to these wonderful people ✨

![Contributors](https://contrib.rocks/image?repo=shayyz-code/xlsx-json-seed)

## License

MIT License — free for personal & commercial use.

Copyright (c) 2025 shayyz-code.
