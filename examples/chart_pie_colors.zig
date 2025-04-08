const std = @import("std");
const excel = @import("excellent");
const chart = @import("excellent").chart;

const Workbook = excel.Workbook;
const Worksheet = excel.Worksheet;
const Chart = excel.Chart;
const Format = excel.Format;
const ChartFill = excel.ChartFill;
const ChartPoint = chart.ChartPoint;
const Colors = excel.Colors;

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try Workbook.create(
        allocator,
        "chart_pie_colors.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);

    // Write data for the chart
    try worksheet.writeString(0, 0, "Pass", null);
    try worksheet.writeString(1, 0, "Fail", null);
    try worksheet.writeNumber(0, 1, 90, null);
    try worksheet.writeNumber(1, 1, 10, null);

    // Create a pie chart
    var chart1 = try workbook.addChart(.pie);

    // Add the data series to the chart
    var series = try chart1.addSeries("=Sheet1!$A$1:$A$2", "=Sheet1!$B$1:$B$2");

    // Create fills for chart segments
    var green_fill = ChartFill{ .color = Colors.green };
    var red_fill = ChartFill{ .color = Colors.red };

    // Create points with fills
    var green_point = ChartPoint{ .fill = &green_fill };
    var red_point = ChartPoint{ .fill = &red_fill };

    // Set the points for the series
    try series.setPoints(&[_]*ChartPoint{ &green_point, &red_point });

    // Insert chart into worksheet
    try worksheet.insertChart(1, 3, chart1);
}
