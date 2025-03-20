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
        "chart.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet("Sheet1");

    // Write some data for the chart
    const data = [_][3]f64{
        [3]f64{ 1, 2, 3 },
        [3]f64{ 2, 4, 6 },
        [3]f64{ 3, 6, 9 },
        [3]f64{ 4, 8, 12 },
        [3]f64{ 5, 10, 15 },
    };

    // Write the data to the worksheet
    for (data, 0..) |row_data, row| {
        for (row_data, 0..) |value, col| {
            try worksheet.writeNumber(
                row,
                col,
                value,
                null,
            );
        }
    }

    // Create a column chart
    var chart = try workbook.addChart(.column);

    // Add data series to the chart
    _ = try chart.addSeries(null, "Sheet1!$A$1:$A$5");
    _ = try chart.addSeries(null, "Sheet1!$B$1:$B$5");
    _ = try chart.addSeries(null, "Sheet1!$C$1:$C$5");

    // Set chart title and font
    try chart.setTitle("Year End Results");
    const font = excel.ChartFont{
        .name = "Chart example",
        .size = 18,
        .color = 0x0000FF, // Blue color
        .bold = false,
    };
    chart.setTitleFont(font);

    // Insert the chart into the worksheet at B7
    try worksheet.insertChart(6, 1, chart);
}
