const std = @import("std");
const excellent = @import("excellent");
const Workbook = excellent.Workbook;
const Worksheet = excellent.Worksheet;
const Chart = excellent.Chart;
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

    for (data, 0..) |row, i| {
        for (row, 0..) |val, j| {
            try worksheet.writeNumber(@intCast(i + 1), @intCast(j), @as(f64, @floatFromInt(val)), null);
        }
    }
}

pub fn main() !void {
    var workbook = try Workbook.create(std.heap.page_allocator, "chart_data_table.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);
    var bold = try workbook.addFormat();
    _ = bold.setBold();

    try writeWorksheetData(&worksheet, bold);

    // Chart 1: Column chart with data table
    var chart1 = try workbook.addChart(.column);
    var series1 = try chart1.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series1.setName("=Sheet1!$B$1");

    var series2 = try chart1.addSeries(null, null);
    try series2.setCategories("Sheet1", 1, 0, 6, 0);
    try series2.setValues("Sheet1", 1, 2, 6, 2);
    try series2.setNameRange("Sheet1", 0, 2);

    try chart1.setTitle("Chart with Data Table");
    try chart1.setXAxisName("Test number");
    try chart1.setYAxisName("Sample length (mm)");

    chart1.setTable();
    try worksheet.insertChart(1, 4, chart1);

    // Chart 2: Column chart with data table and legend keys
    var chart2 = try workbook.addChart(.column);
    var series3 = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series3.setName("=Sheet1!$B$1");

    var series4 = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series4.setName("=Sheet1!$C$1");

    try chart2.setTitle("Data Table with legend keys");
    try chart2.setXAxisName("Test number");
    try chart2.setYAxisName("Sample length (mm)");

    chart2.setTable();
    chart2.setTableGrid(true, true, true, true);
    chart2.setLegendPosition(.none);

    try worksheet.insertChart(17, 4, chart2);
}
