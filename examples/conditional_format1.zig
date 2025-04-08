const std = @import("std");
const excel = @import("excellent");
const cf = excel.cf;

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(
        allocator,
        "conditional_format1.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet("Sheet1");

    // Write some sample numbers
    try worksheet.writeNumber(0, 1, 10, null); // B1
    try worksheet.writeNumber(1, 1, 20, null); // B2
    try worksheet.writeNumber(2, 1, 30, null); // B3
    try worksheet.writeNumber(3, 1, 40, null); // B4
    try worksheet.writeNumber(4, 1, 50, null); // B5
    try worksheet.writeNumber(5, 1, 60, null); // B6
    try worksheet.writeNumber(6, 1, 70, null); // B7
    try worksheet.writeNumber(7, 1, 80, null); // B8
    try worksheet.writeNumber(8, 1, 90, null); // B9

    // Create a format with red text
    var format = try workbook.addFormat();
    _ = format.setFontColor(excel.Colors.red);

    try worksheet.conditionalFormatRange(
        "B1:B9",
        cf.cellLessThan(33, format),
    );

    // Save the workbook
    try workbook.close();
}
