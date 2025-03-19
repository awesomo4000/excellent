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
        "utf8.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);
    try worksheet.writeString(2, 1, "Это фраза на русском!", null);
}
