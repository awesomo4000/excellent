const std = @import("std");
const testing = std.testing;
const excellent = @import("excellent.zig");
const Chart = excellent.Chart;
const ChartType = excellent.ChartType;
const ChartFont = excellent.ChartFont;
const ChartLegendPosition = excellent.ChartLegendPosition;
const Workbook = excellent.Workbook;

test "Chart - creation and basic operations" {
    const allocator = testing.allocator;
    const filename = "test_chart.xlsx";
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

    // Add a worksheet before creating charts
    const worksheet = try workbook.addWorksheet("Sheet1");
    _ = worksheet;

    var chart = try Chart.init(allocator, workbook.workbook, .column);
    defer chart.deinit();
    try chart.addSeries("=Sheet1!$A$1:$A$5", "=Sheet1!$B$1:$B$5");
    try chart.setTitle("Test Chart");
    chart.setStyle(2);
    chart.setLegendPosition(.right);
}

test "Chart - font customization" {
    const allocator = testing.allocator;
    const filename = "test_chart_font.xlsx";
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

    // Add a worksheet before creating charts
    const worksheet = try workbook.addWorksheet("Sheet1");
    _ = worksheet;

    var chart = try Chart.init(allocator, workbook.workbook, .line);
    defer chart.deinit();
    try chart.setTitle("Test Chart");

    const font = ChartFont{
        .name = "Arial",
        .size = 12.0,
        .bold = true,
        .italic = false,
    };
    chart.setTitleFont(font);
}

test "Chart - all types support" {
    const allocator = testing.allocator;
    const filename = "test_chart_types.xlsx";
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

    // Add a worksheet before creating charts
    const worksheet = try workbook.addWorksheet("Sheet1");
    _ = worksheet;

    const chart_types = [_]ChartType{ .column, .bar, .line, .pie, .scatter, .area, .radar, .doughnut };
    for (chart_types) |chart_type| {
        var chart = try Chart.init(allocator, workbook.workbook, chart_type);
        defer chart.deinit();
        try chart.setTitle("Test Chart");
        chart.setStyle(1);
    }
}
