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
    var workbook = try Workbook.create(std.heap.page_allocator, "chart_doughnut.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);
    const bold = try workbook.addFormat();
    _ = bold.setBold();

    // Write data for the charts
    try writeWorksheetData(&worksheet, bold);

    // Chart 1: Simple doughnut chart
    var chart1 = try workbook.addChart(.doughnut);
    var series = try chart1.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    try series.setName("Doughnut sales data");
    try chart1.setTitle("Popular Doughnut Types");
    _ = chart1.setStyle(10);
    try worksheet.insertChart(1, 3, chart1);

    // Chart 2: Doughnut chart with custom colors
    chart1 = try workbook.addChart(.doughnut);
    series = try chart1.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    try series.setName("Doughnut sales data");
    try chart1.setTitle("Doughnut Chart with user defined colors");

    // Create custom fills for the segments
    var fill1 = ChartFill{ .color = 0xFA58D0 };
    var fill2 = ChartFill{ .color = 0x61210B };
    var fill3 = ChartFill{ .color = 0xF5F6CE };

    // Create points with the fills
    var point1 = ChartPoint{ .fill = &fill1 };
    var point2 = ChartPoint{ .fill = &fill2 };
    var point3 = ChartPoint{ .fill = &fill3 };

    // Set the points for the series
    try series.setPoints(&[_]*ChartPoint{ &point1, &point2, &point3 });
    try worksheet.insertChart(17, 3, chart1);

    // Chart 3: Doughnut chart with rotation
    chart1 = try workbook.addChart(.doughnut);
    series = try chart1.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    try series.setName("Doughnut sales data");
    try chart1.setTitle("Doughnut Chart with segment rotation");
    _ = chart1.setRotation(90);
    try worksheet.insertChart(33, 3, chart1);

    // Chart 4: Doughnut chart with hole size and other options
    chart1 = try workbook.addChart(.doughnut);
    series = try chart1.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    try series.setName("Doughnut sales data");
    try chart1.setTitle("Doughnut Chart with options applied.");
    try series.setPoints(&[_]*ChartPoint{ &point1, &point2, &point3 });
    _ = chart1.setStyle(26);
    _ = chart1.setRotation(28);
    _ = chart1.setHoleSize(33);
    try worksheet.insertChart(49, 3, chart1);
}

fn writeWorksheetData(worksheet: *Worksheet, bold: *Format) !void {
    try worksheet.writeString(0, 0, "Category", bold);
    try worksheet.writeString(1, 0, "Glazed", null);
    try worksheet.writeString(2, 0, "Chocolate", null);
    try worksheet.writeString(3, 0, "Cream", null);

    try worksheet.writeString(0, 1, "Values", bold);
    try worksheet.writeNumber(1, 1, 50, null);
    try worksheet.writeNumber(2, 1, 35, null);
    try worksheet.writeNumber(3, 1, 15, null);
}
