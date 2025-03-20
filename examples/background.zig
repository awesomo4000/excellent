const std = @import("std");
const excel = @import("excellent");
const TmpFile = excel.TmpFile;
const Workbook = excel.Workbook;
const Worksheet = excel.Worksheet;

const logo_data = @embedFile("logo.png");
pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var tmp_file = try TmpFile.create(
        allocator,
        "logo_",
    );
    defer tmp_file.cleanUp();

    try tmp_file.write(logo_data);

    // Create a new workbook
    var workbook = try Workbook.create(
        allocator,
        "background.xlsx",
    );
    defer workbook.deinit();

    // Add a worksheet
    var worksheet = try workbook.addWorksheet(null);

    // Set the background using the logo from the root directory
    try worksheet.setBackground(tmp_file.path);

    // Close the workbook to ensure it's written to disk
    try workbook.close();
}
