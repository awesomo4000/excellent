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
        "anatomy.xlsx",
    );
    defer workbook.deinit();

    var worksheet1 = try workbook.addWorksheet("Fruits");
    var worksheet2 = try workbook.addWorksheet("Vegetables");

    try worksheet1.writeString(0, 0, "Peach", null);
    try worksheet1.writeString(1, 0, "Plum", null);
    try worksheet1.writeString(2, 0, "Grape", null);

    try worksheet2.writeString(0, 0, "Carrot", null);
    try worksheet2.writeString(1, 0, "Potato", null);
    try worksheet2.writeString(2, 0, "Spinach", null);

    try workbook.close();
}
