const std = @import("std");
const excellent = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    // Create a new workbook
    var workbook = try excellent.Workbook.create(
        allocator,
        "chart_working_with_example.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    // Add a worksheet
    var worksheet = try workbook.addWorksheet(null);

    // Write some data for the chart
    try worksheet.writeNumber(0, 0, 10, null);
    try worksheet.writeNumber(1, 0, 40, null);
    try worksheet.writeNumber(2, 0, 50, null);
    try worksheet.writeNumber(3, 0, 20, null);
    try worksheet.writeNumber(4, 0, 10, null);
    try worksheet.writeNumber(5, 0, 50, null);

    // Create a line chart
    var chart = try workbook.addChart(.line);

    // Add a data series to the chart
    const series = try chart.addSeries(null, "Sheet1!$A$1:$A$6");
    _ = series; // Used in other examples

    // Insert the chart into the worksheet
    try worksheet.insertChart(0, 2, chart);
}
