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
        "chart_clustered.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet("Sheet1");

    // Add a bold format for headers
    var bold_format = try workbook.addFormat();
    _ = bold_format.setBold();

    // Write the headers
    try worksheet.writeString(0, 0, "Types", bold_format);
    try worksheet.writeString(0, 1, "Sub Type", bold_format);
    try worksheet.writeString(0, 2, "Value 1", bold_format);
    try worksheet.writeString(0, 3, "Value 2", bold_format);
    try worksheet.writeString(0, 4, "Value 3", bold_format);

    // Write Type 1 data
    try worksheet.writeString(1, 0, "Type 1", null);
    try worksheet.writeString(1, 1, "Sub Type A", null);
    try worksheet.writeNumber(1, 2, 5000, null);
    try worksheet.writeNumber(1, 3, 8000, null);
    try worksheet.writeNumber(1, 4, 6000, null);

    try worksheet.writeString(2, 1, "Sub Type B", null);
    try worksheet.writeNumber(2, 2, 2000, null);
    try worksheet.writeNumber(2, 3, 3000, null);
    try worksheet.writeNumber(2, 4, 4000, null);

    try worksheet.writeString(3, 1, "Sub Type C", null);
    try worksheet.writeNumber(3, 2, 250, null);
    try worksheet.writeNumber(3, 3, 1000, null);
    try worksheet.writeNumber(3, 4, 2000, null);

    // Write Type 2 data
    try worksheet.writeString(4, 0, "Type 2", null);
    try worksheet.writeString(4, 1, "Sub Type D", null);
    try worksheet.writeNumber(4, 2, 6000, null);
    try worksheet.writeNumber(4, 3, 6000, null);
    try worksheet.writeNumber(4, 4, 6500, null);

    try worksheet.writeString(5, 1, "Sub Type E", null);
    try worksheet.writeNumber(5, 2, 500, null);
    try worksheet.writeNumber(5, 3, 300, null);
    try worksheet.writeNumber(5, 4, 200, null);

    // Create a column chart
    var chart = try excel.Chart.init(
        allocator,
        workbook.workbook,
        .column,
    );

    // Configure the series with 2D ranges for categories (A2:B6)
    // This creates the clusters
    try chart.addSeries("Sheet1!$A$2:$B$6", "Sheet1!$C$2:$C$6");
    try chart.addSeries("Sheet1!$A$2:$B$6", "Sheet1!$D$2:$D$6");
    try chart.addSeries("Sheet1!$A$2:$B$6", "Sheet1!$E$2:$E$6");

    // Set chart title and style
    try chart.setTitle("Clustered Chart");
    chart.setStyle(37);

    // Turn off the legend
    chart.setLegendPosition(.none);

    // Insert the chart into the worksheet at cell G3
    try worksheet.insertChart(2, 6, chart);
}
