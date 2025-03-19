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
        "array_formula.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Write some data for the formulas
    try worksheet.writeNumber(0, 1, 500, null);
    try worksheet.writeNumber(1, 1, 10, null);
    try worksheet.writeNumber(4, 1, 1, null);
    try worksheet.writeNumber(5, 1, 2, null);
    try worksheet.writeNumber(6, 1, 3, null);

    try worksheet.writeNumber(0, 2, 300, null);
    try worksheet.writeNumber(1, 2, 15, null);
    try worksheet.writeNumber(4, 2, 20234, null);
    try worksheet.writeNumber(5, 2, 21003, null);
    try worksheet.writeNumber(6, 2, 10000, null);

    // Write an array formula that returns a single value
    try worksheet.writeArrayFormula(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", null);

    // Similar to above but using a cell range reference
    try worksheet.writeArrayFormulaRange("A2:A2", "{=SUM(B1:C1*B2:C2)}", null);

    // Write an array formula that returns a range of values
    try worksheet.writeArrayFormula(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", null);

    try workbook.close();
}
