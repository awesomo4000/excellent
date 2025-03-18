const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa =
        std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(
        allocator,
        "hello.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);
    try worksheet.writeString(0, 0, "Hello, Excel!", null);
    try worksheet.writeNumber(1, 0, 123, null);
}
