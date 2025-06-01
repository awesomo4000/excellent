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
        "dates_and_times03.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add a format with date formatting.
    var format = try workbook.addFormat();
    _ = try format.setNumFormat("mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    worksheet.setColumnWidth(0, 0, 20);

    // Write some unix datetimes with formatting.
    // 1970-01-01. The Unix epoch.
    try worksheet.writeUnixTime(0, 0, 0, format);

    // 2000-01-01.
    try worksheet.writeUnixTime(1, 0, 1577836800, format);

    // 1900-01-01.
    try worksheet.writeUnixTime(2, 0, -2208988800, format);

    try workbook.close();
}
