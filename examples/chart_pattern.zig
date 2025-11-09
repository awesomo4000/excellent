const std = @import("std");
const excellent = @import("excellent");
const Colors = excellent.Colors;

pub fn main() !void {
    const allocator = std.heap.page_allocator;

    // Create a new workbook
    var workbook = try excellent.Workbook.create(allocator, "chart_pattern.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    // Add a worksheet
    var worksheet = try workbook.addWorksheet("Sheet1");
    defer worksheet.deinit();
    // Add a bold format to use to highlight the header cells
    var bold = try workbook.addFormat();
    defer bold.deinit();
    _ = bold.setBold();

    // Write some data for the chart
    try worksheet.writeString(0, 0, "Shingle", bold);
    try worksheet.writeNumber(1, 0, 105, null);
    try worksheet.writeNumber(2, 0, 150, null);
    try worksheet.writeNumber(3, 0, 130, null);
    try worksheet.writeNumber(4, 0, 90, null);

    try worksheet.writeString(0, 1, "Brick", bold);
    try worksheet.writeNumber(1, 1, 50, null);
    try worksheet.writeNumber(2, 1, 120, null);
    try worksheet.writeNumber(3, 1, 100, null);
    try worksheet.writeNumber(4, 1, 110, null);

    // Create a chart
    var chart = try workbook.addChart(.column);

    // Configure the chart series
    var series1 = try chart.addSeries(null, "Sheet1!$A$2:$A$5");
    var series2 = try chart.addSeries(null, "Sheet1!$B$2:$B$5");

    try series1.setName("=Sheet1!$A$1");
    try series2.setName("=Sheet1!$B$1");

    // Set the chart title and axis names
    try chart.setTitle("Cladding types");
    try chart.setXAxisName("Region");
    try chart.setYAxisName("Number of houses");

    // Configure and add the chart series patterns
    try series1.setPattern(.{
        .pattern_type = .shingle,
        .fg_color = Colors.saddlebrown,
        .bg_color = Colors.aztecgold,
    });

    try series2.setPattern(.{
        .pattern_type = .horizontal_brick,
        .fg_color = Colors.artfulred,
        .bg_color = Colors.pompelmo,
    });

    // Configure and set the chart series borders
    try series1.setLine(.{
        .color = Colors.saddlebrown,
        .width = 0.0,
        .dash_type = .solid,
    });

    try series2.setLine(.{
        .color = Colors.artfulred,
        .width = 0.0,
        .dash_type = .solid,
    });
    chart.setSeriesGap(70);

    // Insert the chart into the worksheet
    try worksheet.insertChart(1, 3, chart);

    _ = try workbook.close();
}
