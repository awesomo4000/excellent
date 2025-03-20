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
        "chart_bar.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet("Sheet1");

    // Add a bold format for headers
    var bold_format = try workbook.addFormat();
    _ = bold_format.setBold();

    // Write the headers
    try worksheet.writeString(0, 0, "Number", bold_format);
    try worksheet.writeString(0, 1, "Batch 1", bold_format);
    try worksheet.writeString(0, 2, "Batch 2", bold_format);

    // Write the data
    const data = [_][3]u8{
        [_]u8{ 2, 10, 30 },
        [_]u8{ 3, 40, 60 },
        [_]u8{ 4, 50, 70 },
        [_]u8{ 5, 20, 50 },
        [_]u8{ 6, 10, 40 },
        [_]u8{ 7, 50, 30 },
    };

    for (data, 0..) |row_data, row_idx| {
        for (row_data, 0..) |cell_value, col_idx| {
            try worksheet.writeNumber(
                @intCast(row_idx + 1),
                @intCast(col_idx),
                @floatFromInt(cell_value),
                null,
            );
        }
    }

    // Chart 1: Create a bar chart
    var chart1 = try workbook.addChart(.bar);

    // Add the first series
    var c1series1 = try chart1.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$B$2:$B$7");
    try c1series1.setName("Batch 1");
    // Add the second series
    var c1series2 = try chart1.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$C$2:$C$7");
    try c1series2.setName("Batch 2");

    // Set chart title
    try chart1.setTitle("Results of sample analysis");

    // Set chart style
    chart1.setStyle(11);
    try chart1.setXAxisName("Test number");
    try chart1.setYAxisName("Sample length (mm)");
    // Insert the chart into the worksheet
    try worksheet.insertChart(1, 4, chart1);

    // Chart 2: Create a stacked bar chart
    var chart2 = try workbook.addChart(.bar_stacked);

    // Add the first series
    var c2series1 = try chart2.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$B$2:$B$7");
    try c2series1.setName("Batch 1");
    // Add the second series
    var c2series2 = try chart2.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$C$2:$C$7");
    try c2series2.setName("Batch 2");
    // Set chart title
    try chart2.setTitle("Results of sample analysis");
    try chart2.setXAxisName("Test number");
    try chart2.setYAxisName("Sample length (mm)");
    // Set chart style
    chart2.setStyle(12);

    // Insert the chart into the worksheet
    try worksheet.insertChart(17, 4, chart2);

    // Chart 3: Create a percent stacked bar chart
    var chart3 = try workbook.addChart(.bar_stacked_percent);

    // Add the first series
    var c3series1 = try chart3.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$B$2:$B$7");
    try c3series1.setName("Batch 1");
    // Add the second series
    var c3series2 = try chart3.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$C$2:$C$7");
    try c3series2.setName("Batch 2");

    // Set chart title
    try chart3.setTitle("Results of sample analysis");
    try chart3.setXAxisName("Test number");
    try chart3.setYAxisName("Sample length (mm)");
    // Set chart style
    chart3.setStyle(13);

    // Insert the chart into the worksheet
    try worksheet.insertChart(33, 4, chart3);
}
