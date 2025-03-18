const std = @import("std");
const excel = @import("excellent");

const Expense = struct {
    item: []const u8,
    cost: f64,
};

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    const expenses = [_]Expense{
        .{ .item = "Rent", .cost = 1000.0 },
        .{ .item = "Gas", .cost = 100.0 },
        .{ .item = "Food", .cost = 300.0 },
        .{ .item = "Gym", .cost = 50.0 },
    };

    var workbook = try excel.Workbook.create(
        allocator,
        "tutorial1.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);
    var row: u32 = 0;

    // Write expenses
    for (expenses) |expense| {
        try worksheet.writeString(row, 0, expense.item, null);
        try worksheet.writeNumber(row, 1, expense.cost, null);
        row += 1;
    }

    // Write total
    try worksheet.writeString(row, 0, "Total", null);
    try worksheet.writeFormula(row, 1, "=SUM(B1:B4)", null);

    try workbook.close();
}
