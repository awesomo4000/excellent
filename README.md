# Excel-lent

A high-level Zig wrapper for zig-xlswriter, which calls libxlsxwriter.

## Features

- Make xlsx files

## Requirements

- Zig 0.14.0 

## Installation

```zig fetch --save=excellent https://github.com/USER/REPO```

## Build

#### Everything
```zig build``

#### Single example
```zig build -Dexample=<name>```

(`zig build --help` will show the examples list)


## Quick Start

```zig
const xlsx = @import("excellent");

pub fn main() !void {
    var workbook = try xlsx.Workbook.init("example.xlsx");
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet("Sheet1");
    try worksheet.writeString(0, 0, "Hello, Excel!");
    try worksheet.writeNumber(1, 0, 42);

    try workbook.close();
}
```
