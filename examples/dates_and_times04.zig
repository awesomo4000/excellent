const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    // A datetime to display.
    const datetime = excel.DateTime{
        .year = 2013,
        .month = 1,
        .day = 23,
        .hour = 12,
        .minute = 30,
        .second = 5.123,
    };

    var row: u32 = 0;
    const col: u16 = 0;

    // Examples date and time formats. In the output file compare how changing
    // the format strings changes the appearance of the date.
    const date_formats = [_][]const u8{
        "dd/mm/yy",
        "mm/dd/yy",
        "dd m yy",
        "d mm yy",
        "d mmm yy",
        "d mmmm yy",
        "d mmmm yyy",
        "d mmmm yyyy",
        "dd/mm/yy hh:mm",
        "dd/mm/yy hh:mm:ss",
        "dd/mm/yy hh:mm:ss.000",
        "hh:mm",
        "hh:mm:ss",
        "hh:mm:ss.000",
    };

    // Create a new workbook and add a worksheet.
    var workbook = try excel.Workbook.create(
        allocator,
        "dates_and_times04.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add a bold format.
    var bold = try workbook.addFormat();
    _ = bold.setBold();

    // Write the column headers.
    try worksheet.writeString(row, col, "Formatted date", bold);
    try worksheet.writeString(row, col + 1, "Format", bold);

    // Widen the first column to make the text clearer.
    worksheet.setColumnWidth(0, 1, 20);

    // Write the same date and time using each of the above formats.
    for (date_formats, 0..) |format_string, i| {
        _ = i;
        row += 1;

        // Create a format for the date or time.
        var format = try workbook.addFormat();
        _ = try format.setNumFormat(format_string);
        _ = format.setAlign(.left);

        // Write the datetime with each format.
        try worksheet.writeDateTime(row, col, datetime, format);

        // Also write the format string for comparison.
        try worksheet.writeString(row, col + 1, format_string, null);
    }

    try workbook.close();
}
