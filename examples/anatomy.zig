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
    // this will deinit the formats, so we don't need to do it manually
    defer workbook.deinit();
    var my_format1 = try workbook.addFormat();

    _ = my_format1.setBold();

    var my_format2 = try workbook.addFormat();
    _ = try my_format2.setNumFormat("$#,##0.00");

    var worksheet1 = try workbook.addWorksheet("Demo");
    var worksheet2 = try workbook.addWorksheet("Sheet2");

    try worksheet1.writeString(0, 0, "Peach", null);
    try worksheet1.writeString(1, 0, "Plum", my_format1);
    try worksheet1.writeString(2, 0, "Pear", my_format1);
    try worksheet1.writeString(3, 0, "Persimmon", my_format1);
    try worksheet1.writeNumber(5, 0, 123, null);
    try worksheet1.writeNumber(6, 0, 4567.555, my_format2);

    try worksheet2.writeString(0, 0, "Some text", my_format1);
    try workbook.close();
}
