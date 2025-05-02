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
    try worksheet.writeNumber(0, 1, 34, null); // B1
    try worksheet.writeNumber(1, 1, 32, null); // B2
    try worksheet.writeNumber(2, 1, 31, null); // B3
    try worksheet.writeNumber(3, 1, 35, null); // B4
    try worksheet.writeNumber(4, 1, 36, null); // B5
    try worksheet.writeNumber(5, 1, 30, null); // B6
    try worksheet.writeNumber(6, 1, 38, null); // B7
    try worksheet.writeNumber(7, 1, 38, null); // B8
    try worksheet.writeNumber(8, 1, 32, null); // B9

    // Create a format with red text
    var format = try workbook.addFormat();
    _ = format.setFontColor(excel.Colors.red);
    const condition = cf.cellLessThan(33, format);
    try worksheet.conditionalFormatRange("B1:B9", condition);

    // Save the workbook
    try workbook.close();
}
