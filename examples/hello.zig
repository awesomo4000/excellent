const std = @import("std");
const excellent = @import("excellent");

pub fn main() !void {
    var workbook = try excellent.Workbook.init("hello.xlsx");
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);
    try worksheet.writeString(0, 0, "Hello, Excel!");
    try worksheet.writeNumber(1, 0, 123);
}
