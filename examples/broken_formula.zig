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
        "broken_formula.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add some data
    try worksheet.writeString(0, 0, "Value A", null);
    try worksheet.writeString(0, 1, "Value B", null);
    try worksheet.writeString(0, 2, "Sum", null);

    try worksheet.writeNumber(1, 0, 10, null);
    try worksheet.writeNumber(1, 1, 20, null);

    // This is a broken formula where we're including the result cell in its own calculation
    // which should be detected as a circular reference in the autocheck
    try worksheet.writeFormula(1, 2, "=SUM(A2:C2)", null);

    // This formula includes a null terminator which should be detected as an issue
    try worksheet.writeFormula(2, 2, "=A2+B2\x00", null);

    // This formula has incorrect range syntax which should be detected
    try worksheet.writeFormula(3, 2, "=SUM(A2:)", null);

    // This string contains a null terminator which should be detected
    try worksheet.writeString(4, 0, "String with null\x00", null);

    // This string is truncated which should be detected
    try worksheet.writeString(4, 1, "Truncated...", null);
}
