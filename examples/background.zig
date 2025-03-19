const std = @import("std");
const excel = @import("excellent");
const Workbook = excel.Workbook;
const Worksheet = excel.Worksheet;

pub fn main() !void {
    // Create a temporary file for the logo using the TmpFile API
    var arena = std.heap.ArenaAllocator.init(
        std.heap.page_allocator,
    );
    defer arena.deinit();
    const allocator = arena.allocator();

    // Create a new workbook
    var workbook = try Workbook.create(allocator, "background.xlsx");
    defer workbook.deinit();

    // Add a worksheet
    var worksheet = try workbook.addWorksheet(null);

    // Set the background using the logo from the root directory
    try worksheet.setBackground("logo.png");

    // Close the workbook to ensure it's written to disk
    try workbook.close();
}
