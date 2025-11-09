const std = @import("std");
const excel = @import("excellent");

// Some data we want to write to the worksheet.
const Expense = struct {
    item: []const u8,
    cost: i32,
    date: excel.DateTime,
};

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(
        allocator,
        "tutorial3.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add a bold format to use to highlight cells.
    var bold = try workbook.addFormat();

    _ = bold.setBold();

    // Add a number format for cells with money.
    var money = try workbook.addFormat();
    _ = try money.setNumFormat("$#,##0");

    // Add an Excel date format.
    var date_format = try workbook.addFormat();
    _ = try date_format.setNumFormat("mmmm d yyyy");

    // Adjust the column width.
    worksheet.setColumnWidth(0, 0, 15);

    // Create our expense data
    const expenses = [_]Expense{
        .{ .item = "Rent", .cost = 1000, .date = .{ .year = 2013, .month = 1, .day = 13 } },
        .{ .item = "Gas", .cost = 100, .date = .{ .year = 2013, .month = 1, .day = 14 } },
        .{ .item = "Food", .cost = 300, .date = .{ .year = 2013, .month = 1, .day = 16 } },
        .{ .item = "Gym", .cost = 50, .date = .{ .year = 2013, .month = 1, .day = 20 } },
    };

    // Write some data header.
    try worksheet.writeString(0, 0, "Item", bold);
    try worksheet.writeString(0, 1, "Cost", bold);

    // Iterate over the data and write it out element by element.
    for (expenses, 0..) |expense, i| {
        const row = i + 1;

        try worksheet.writeString(row, 0, expense.item, null);

        // Write the date
        try worksheet.writeDateTime(row, 1, expense.date, date_format);

        try worksheet.writeNumber(row, 2, @floatFromInt(expense.cost), money);
    }

    // Write a total using a formula.
    const total_row = expenses.len + 1;
    try worksheet.writeString(total_row, 0, "Total", bold);
    try worksheet.writeFormula(total_row, 2, "=SUM(C2:C5)", money);

    try workbook.close();
}
