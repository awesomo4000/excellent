const std = @import("std");
const excel = @import("excellent");

const Expense = struct {
    item: []const u8,
    cost: i32,
};

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    const expenses = [_]Expense{
        .{ .item = "Rent", .cost = 1000 },
        .{ .item = "Gas", .cost = 100 },
        .{ .item = "Food", .cost = 300 },
        .{ .item = "Gym", .cost = 50 },
    };

    var workbook = try excel.Workbook.create(
        allocator,
        "tutorial2.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add a bold format to use to highlight cells
    var bold = try workbook.addFormat();

    _ = bold.setBold();

    // Add a number format for cells with money
    var money = try workbook.addFormat();
    _ = try money.setNumFormat("$#,##0");

    // Write the header row
    try worksheet.writeString(0, 0, "Item", bold);
    try worksheet.writeString(0, 1, "Cost", bold);

    // Write the expense data
    for (expenses, 0..) |expense, i| {
        const row = i + 1;
        try worksheet.writeString(row, 0, expense.item, null);
        try worksheet.writeNumber(row, 1, @floatFromInt(expense.cost), money);
    }

    // Write the total row
    const total_row = expenses.len + 1;
    try worksheet.writeString(total_row, 0, "Total", bold);

    // Be very explicit about the exact cell range we want to sum
    // Create a formula that matches exactly what the rows are
    var formula_buf: [64]u8 = undefined;
    const formula = try std.fmt.bufPrint(&formula_buf, "=SUM(B2:B{d})", .{
        expenses.len + 1,
    });

    try worksheet.writeFormula(total_row, 1, formula, money);

    try workbook.close();
}
