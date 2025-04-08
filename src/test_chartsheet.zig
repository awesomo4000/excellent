const std = @import("std");
const testing = std.testing;
const excellent = @import("excellent.zig");
const Chart = excellent.Chart;
const ChartType = excellent.ChartType;
const Workbook = excellent.Workbook;
const Chartsheet = excellent.Chartsheet;

test "Chartsheet - basic operations" {
    const allocator = testing.allocator;
    const filename = "test_chartsheet.xlsx";
    // Ensure file doesn't exist at start
    std.fs.cwd().deleteFile(filename) catch |err| switch (err) {
        error.FileNotFound => {},
        else => return err,
    };

    var workbook = try Workbook.create(allocator, filename);
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
        std.fs.cwd().deleteFile(filename) catch |err| {
            std.debug.print("Failed to clean up test file: {}\n", .{err});
        };
    }

    // Add a worksheet for data
    const worksheet = try workbook.addWorksheet("Sheet1");
    _ = worksheet;

    // Add a chartsheet
    const chartsheet = try workbook.addChartsheet("Chart1");
    defer chartsheet.deinit();

    // Create a chart
    var chart = try Chart.init(allocator, workbook.workbook, .bar);
    defer chart.deinit();

    // Set up the chart
    try chart.setTitle("Test Chart");
    chart.setStyle(11);

    // Set the chart on the chartsheet
    try chartsheet.setChart(&chart);

    // Make the chartsheet active
    try chartsheet.activate();

    // Set zoom level
    chartsheet.setZoom(100);
}

test "Chartsheet - multiple chartsheets" {
    const allocator = testing.allocator;
    const filename = "test_multiple_chartsheets.xlsx";
    // Ensure file doesn't exist at start
    std.fs.cwd().deleteFile(filename) catch |err| switch (err) {
        error.FileNotFound => {},
        else => return err,
    };

    var workbook = try Workbook.create(allocator, filename);
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
        std.fs.cwd().deleteFile(filename) catch |err| {
            std.debug.print("Failed to clean up test file: {}\n", .{err});
        };
    }

    // Add a worksheet for data
    const worksheet = try workbook.addWorksheet("Sheet1");
    _ = worksheet;

    // Add multiple chartsheets
    const chartsheet1 = try workbook.addChartsheet("Chart1");
    defer chartsheet1.deinit();

    const chartsheet2 = try workbook.addChartsheet("Chart2");
    defer chartsheet2.deinit();

    // Create charts for each chartsheet
    var chart1 = try Chart.init(allocator, workbook.workbook, .bar);
    defer chart1.deinit();
    try chart1.setTitle("First Chart");
    chart1.setStyle(11);

    var chart2 = try Chart.init(allocator, workbook.workbook, .line);
    defer chart2.deinit();
    try chart2.setTitle("Second Chart");
    chart2.setStyle(12);

    // Set charts on their respective chartsheets
    try chartsheet1.setChart(&chart1);
    try chartsheet2.setChart(&chart2);

    // Make one chartsheet active
    try chartsheet1.activate();
}
