const std = @import("std");
const excel = @import("excellent");
const chart = @import("excellent").chart;

const Workbook = excel.Workbook;
const Worksheet = excel.Worksheet;
const Chart = excel.Chart;
const Format = excel.Format;
const ChartFill = excel.ChartFill;
const ChartPoint = chart.ChartPoint;

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try Workbook.create(
        allocator,
        "chart_pie.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);
    const bold = try workbook.addFormat();
    _ = bold.setBold();

    // Write data for the charts
    try writeWorksheetData(&worksheet, bold);

    // Chart 1: Simple pie chart
    var chart1 = try workbook.addChart(.pie);
    var series = try chart1.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    try series.setName("Pie sales data");
    try chart1.setTitle("Popular Pie Types");
    _ = chart1.setStyle(10);
    try worksheet.insertChart(1, 3, chart1);

    // Chart 2: Pie chart with custom colors
    var chart2 = try workbook.addChart(.pie);
    series = try chart2.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    try series.setName("Pie sales data");
    try chart2.setTitle("Pie Chart with user defined colors");

    // Create custom fills for the segments
    var fill1 = ChartFill{ .color = 0x5ABA10 };
    var fill2 = ChartFill{ .color = 0xFE110E };
    var fill3 = ChartFill{ .color = 0xCA5C05 };

    // Create points with the fills
    var point1 = ChartPoint{ .fill = &fill1 };
    var point2 = ChartPoint{ .fill = &fill2 };
    var point3 = ChartPoint{ .fill = &fill3 };

    // Set the points for the series
    try series.setPoints(&[_]*ChartPoint{ &point1, &point2, &point3 });
    try worksheet.insertChart(17, 3, chart2);

    // Chart 3: Pie chart with rotation
    var chart3 = try workbook.addChart(.pie);
    series = try chart3.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    try series.setName("Pie sales data");
    try chart3.setTitle("Pie Chart with segment rotation");
    _ = chart3.setRotation(90);
    try worksheet.insertChart(33, 3, chart3);
}

fn writeWorksheetData(worksheet: *Worksheet, bold: *Format) !void {
    try worksheet.writeString(0, 0, "Category", bold);
    try worksheet.writeString(1, 0, "Apple", null);
    try worksheet.writeString(2, 0, "Cherry", null);
    try worksheet.writeString(3, 0, "Pecan", null);

    try worksheet.writeString(0, 1, "Values", bold);
    try worksheet.writeNumber(1, 1, 60, null);
    try worksheet.writeNumber(2, 1, 30, null);
    try worksheet.writeNumber(3, 1, 10, null);
}
