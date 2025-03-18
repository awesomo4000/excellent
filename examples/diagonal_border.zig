//
// A simple formatting example that demonstrates how to add diagonal
// cell borders using the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(
        allocator,
        "diagonal_border.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add some diagonal border formats
    var format1 = try workbook.addFormat();
    _ = format1.setDiagonalType(.up);

    var format2 = try workbook.addFormat();
    _ = format2.setDiagonalType(.down);

    var format3 = try workbook.addFormat();
    _ = format3.setDiagonalType(.up_down);

    var format4 = try workbook.addFormat();
    _ = format4.setDiagonalType(.up_down)
        .setDiagonalBorder(.hair)
        .setDiagonalColor(0xFF0000); // Red color

    // Write formatted strings using cell references
    try worksheet.writeStringCell("B3", "Text", format1);
    try worksheet.writeStringCell("B6", "Text", format2);
    try worksheet.writeStringCell("B9", "Text", format3);
    try worksheet.writeStringCell("B12", "Text", format4);

    try workbook.close();
}
