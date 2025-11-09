const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    // Create a new workbook and add a worksheet.
    var workbook = try excel.Workbook.create(
        allocator,
        "dates_and_times02.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // A datetime to display.
    const datetime = excel.DateTime{
        .year = 2013,
        .month = 2,
        .day = 28,
        .hour = 12,
        .minute = 0,
        .second = 0.0,
    };

    // Add a format with date formatting.
    var format = try workbook.addFormat();
    _ = try format.setNumFormat("mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    worksheet.setColumnWidth(0, 0, 20);

    // Write the datetime without formatting.
    // 41333.5
    try worksheet.writeDateTime(0, 0, datetime, null);

    // Write the datetime with formatting.
    // Feb 28 2013 12:00 PM

    try worksheet.writeDateTime(1, 0, datetime, format);

    try workbook.close();
}
