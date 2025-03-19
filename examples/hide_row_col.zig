const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    // Create a new workbook and add a worksheet
    var workbook = try excel.Workbook.create(
        allocator,
        "hide_row_col.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Write some data
    try worksheet.writeString(0, 3, "Some hidden columns.", null);
    try worksheet.writeString(7, 0, "Some hidden rows.", null);

    // Hide all rows without data
    worksheet.setDefaultRow(15, true);

    // Set the height of empty rows that we want to display even if it is
    // the default height
    var row: u32 = 1;
    while (row <= 6) : (row += 1) {
        worksheet.setRowHeight(row, 15);
    }

    // Columns can be hidden explicitly using column range string
    try worksheet.setColumnOptRange("G:XFD", 8.43, null, true);
}
