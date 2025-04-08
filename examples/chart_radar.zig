const std = @import("std");
const excel = @import("excellent");

const Workbook = excel.Workbook;
const Worksheet = excel.Worksheet;
const Chart = excel.Chart;
const Format = excel.Format;
const Colors = excel.Colors;

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try Workbook.create(
        allocator,
        "chart_radar.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);
    const bold = try workbook.addFormat();
    _ = bold.setBold();

    // Write data for the chart
    try writeWorksheetData(&worksheet, bold);

    // Chart 1: Create a simple radar chart
    var chart1 = try workbook.addChart(.radar);
    var series = try chart1.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined
    // Configure the series using a programmatic approach
    series = try chart1.addSeries(null, null);
    try series.setCategories("Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
    try series.setValues("Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
    try series.setNameRange("Sheet1", 0, 2); // "=Sheet1!$C$1"

    try chart1.setTitle("Results of sample analysis");
    _ = chart1.setStyle(11);
    try worksheet.insertChart(1, 4, chart1);

    // Note: The high-level API currently supports only the basic radar chart type
    // For demonstration purposes, we'll create three different instances with the same type
    // but in a real implementation, we would use the specific radar chart types

    // Chart 2: Create a radar chart (which would ideally be a radar chart with markers)
    var chart2 = try workbook.addChart(.radar_with_markers);
    series = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    series = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    try chart2.setTitle("Results of sample analysis");
    _ = chart2.setStyle(12);
    try worksheet.insertChart(17, 4, chart2);

    // Chart 3: Create a radar chart (which would ideally be a filled radar chart)
    var chart3 = try workbook.addChart(.radar_filled);
    series = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    series = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    try chart3.setTitle("Results of sample analysis");
    _ = chart3.setStyle(13);
    try worksheet.insertChart(33, 4, chart3);
}

fn writeWorksheetData(worksheet: *Worksheet, bold: *Format) !void {
    const data = [_][3]f64{
        // Three columns of data
        [_]f64{ 2, 30, 25 },
        [_]f64{ 3, 60, 40 },
        [_]f64{ 4, 70, 50 },
        [_]f64{ 5, 50, 30 },
        [_]f64{ 6, 40, 50 },
        [_]f64{ 7, 30, 40 },
    };

    try worksheet.writeString(0, 0, "Number", bold);
    try worksheet.writeString(0, 1, "Batch 1", bold);
    try worksheet.writeString(0, 2, "Batch 2", bold);

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            try worksheet.writeNumber(
                @intCast(row_idx + 1),
                @intCast(col_idx),
                value,
                null,
            );
        }
    }
}
