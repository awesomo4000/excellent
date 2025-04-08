const std = @import("std");
const excellent = @import("excellent");

fn writeWorksheetData(worksheet: *excellent.Worksheet, bold: *excellent.Format) !void {
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

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            try worksheet.writeNumber(
                @intCast(row_idx + 1),
                @intCast(col_idx),
                @floatFromInt(value),
                null,
            );
        }
    }
}

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    // Create a new workbook and add sheets
    var workbook = try excellent.Workbook.create(
        allocator,
        "chartsheet.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);
    var chartsheet = try workbook.addChartsheet(null);

    // Add a bold format for headers
    var bold = try workbook.addFormat();
    _ = bold.setBold();

    // Write the data for the chart
    try writeWorksheetData(&worksheet, bold);

    // Create a bar chart
    var chart = try workbook.addChart(.bar);

    // Add the first series to the chart
    var series = try chart.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Set the name for the first series
    try series.setName("=Sheet1!$B$1");

    // Add a second series and configure it programmatically
    series = try chart.addSeries(null, null);

    try series.setCategories("Sheet1", 1, 0, 6, 0);
    try series.setValues("Sheet1", 1, 2, 6, 2);
    try series.setNameRange("Sheet1", 0, 2);

    // Add chart title and axis labels
    try chart.setTitle("Results of sample analysis");
    try chart.setXaxisName("Test number");
    try chart.setYaxisName("Sample length (mm)");

    // Set chart style
    chart.setStyle(11);

    // Add the chart to the chartsheet
    try chartsheet.setChart(chart);

    // Make the chartsheet active
    try chartsheet.activate();
}
