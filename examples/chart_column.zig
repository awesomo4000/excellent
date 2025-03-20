const std = @import("std");
const excellent = @import("excellent");
const Workbook = excellent.Workbook;
const Worksheet = excellent.Worksheet;
const Chart = excellent.Chart;
const ChartType = excellent.ChartType;
const Format = excellent.Format;

fn writeWorksheetData(worksheet: *Worksheet, bold: *Format) !void {
    const data = [_][3]u8{
        .{ 2, 10, 30 },
        .{ 3, 40, 60 },
        .{ 4, 50, 70 },
        .{ 5, 20, 50 },
        .{ 6, 10, 40 },
        .{ 7, 50, 30 },
    };

    try worksheet.writeString(0, 0, "Number", bold);
    try worksheet.writeString(0, 1, "Batch 1", bold);
    try worksheet.writeString(0, 2, "Batch 2", bold);

    for (data, 0..) |row, row_num| {
        for (row, 0..) |value, col_num| {
            try worksheet.writeNumber(@intCast(row_num + 1), @intCast(col_num), @floatFromInt(value), null);
        }
    }
}

pub fn main() !void {
    const allocator = std.heap.page_allocator;
    var workbook = try Workbook.create(allocator, "chart_column.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);
    const bold = try workbook.addFormat();
    _ = bold.setBold();

    // Write some data for the chart.
    try writeWorksheetData(&worksheet, bold);

    // Chart 1: Create a column chart.
    var chart1 = try workbook.addChart(.column);

    // Add the first series
    _ = try chart1.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$B$2:$B$7");

    // Add the second series
    _ = try chart1.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$C$2:$C$7");

    // Add a chart title and some axis labels.
    try chart1.setTitle("Results of sample analysis");
    try chart1.setAxisName(.x_axis, "Test number");
    try chart1.setAxisName(.y_axis, "Sample length (mm)");

    // Set an Excel chart style.
    chart1.setStyle(11);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(1, 4, chart1);

    // Chart 2: Create a stacked column chart.
    var chart2 = try workbook.addChart(.column_stacked);

    // Add the first series
    _ = try chart2.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$B$2:$B$7");

    // Add the second series
    _ = try chart2.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$C$2:$C$7");

    // Add a chart title and some axis labels.
    try chart2.setTitle("Results of sample analysis");
    try chart2.setAxisName(.x_axis, "Test number");
    try chart2.setAxisName(.y_axis, "Sample length (mm)");

    // Set an Excel chart style.
    chart2.setStyle(12);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(17, 4, chart2);

    // Chart 3: Create a percent stacked column chart.
    var chart3 = try workbook.addChart(.column_stacked_percent);

    // Add the first series
    _ = try chart3.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$B$2:$B$7");

    // Add the second series
    _ = try chart3.addSeries("Sheet1!$A$2:$A$7", "Sheet1!$C$2:$C$7");

    // Add a chart title and some axis labels.
    try chart3.setTitle("Results of sample analysis");
    try chart3.setAxisName(.x_axis, "Test number");
    try chart3.setAxisName(.y_axis, "Sample length (mm)");

    // Set an Excel chart style.
    chart3.setStyle(13);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(33, 4, chart3);
}
