const std = @import("std");
const excellent = @import("excellent");
const Chart = excellent.Chart;
const ChartLine = excellent.ChartLine;
const ChartFill = excellent.ChartFill;
const Colors = excellent.Colors;
const Worksheet = excellent.Worksheet;
const Workbook = excellent.Workbook;
const Format = excellent.Format;
const DashType = excellent.ChartLine.DashType;

fn writeWorksheetData(
    worksheet: *Worksheet,
    bold: *Format,
) !void {
    const data = [_][3]u8{
        // Three columns of data
        [_]u8{ 2, 10, 30 },
        [_]u8{ 3, 40, 60 },
        [_]u8{ 4, 50, 70 },
        [_]u8{ 5, 20, 50 },
        [_]u8{ 6, 10, 40 },
        [_]u8{ 7, 50, 30 },
    };

    try worksheet.writeString(
        0,
        0,
        "Number",
        bold,
    );
    try worksheet.writeString(
        0,
        1,
        "Batch 1",
        bold,
    );
    try worksheet.writeString(
        0,
        2,
        "Batch 2",
        bold,
    );

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            try worksheet.writeNumber(
                @intCast(row_idx + 1),
                @intCast(col_idx),
                @floatFromInt(value),
                null,
            );
        }
    }
}

pub fn main() !void {
    // Create a new workbook and work   sheet
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};

    defer if (gpa.deinit() == .leak) {
        std.log.warn("Memory leak detected", .{});
    };
    const allocator = gpa.allocator();
    var workbook = try Workbook.create(
        allocator,
        "chart_data_tools.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet(null);
    defer worksheet.deinit();

    // Add a bold format for headers
    var bold = try workbook.addFormat();
    defer bold.deinit();
    _ = bold.setBold();

    // Write the data
    try writeWorksheetData(&worksheet, bold);

    // Chart 1: High-Low Lines
    var chart1 = try workbook.addChart(.line);
    try chart1.setTitle("Chart with High-Low Lines");

    _ = try chart1.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );
    _ = try chart1.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );
    try chart1.setHighLowLines(null);
    try worksheet.insertChart(1, 4, chart1);

    // Chart 2: Drop Lines
    var chart2 = try workbook.addChart(.line);
    try chart2.setTitle("Chart with Drop Lines");
    _ = try chart2.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );
    _ = try chart2.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );
    try chart2.setDropLines(null);
    try worksheet.insertChart(17, 4, chart2);

    // Chart 3: Up-Down Bars
    var chart3 = try workbook.addChart(.line);
    try chart3.setTitle("Chart with Up-Down bars");
    _ = try chart3.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );
    _ = try chart3.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );
    try chart3.setUpDownBars();
    try worksheet.insertChart(33, 4, chart3);

    // Chart 4: Up-Down Bars with Formatting
    var chart4 = try workbook.addChart(.line);
    try chart4.setTitle("Chart with Up-Down bars");
    _ = try chart4.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );
    _ = try chart4.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );

    var line = ChartLine{ .color = Colors.black, .width = 0.75 };
    var up_fill = ChartFill{ .color = Colors.jade };
    var down_fill = ChartFill{ .color = Colors.red };
    try chart4.setUpDownBarsFormat(&line, &up_fill, &line, &down_fill);
    try worksheet.insertChart(49, 4, chart4);

    // Chart 5: Markers and Data Labels
    var chart5 = try workbook.addChart(.line);
    try chart5.setTitle("Chart with Data Labels and Markers");
    var series5 = try chart5.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );
    _ = try chart5.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );
    try series5.setMarkerType(.circle);
    try series5.setLabels();
    try worksheet.insertChart(65, 4, chart5);

    // Chart 6: Error Bars
    var chart6 = try workbook.addChart(.line);
    try chart6.setTitle("Chart with Error Bars");
    var series6 = try chart6.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );
    _ = try chart6.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );
    try series6.setErrorBars(.std_error, 0);
    try series6.setLabels();
    try worksheet.insertChart(81, 4, chart6);

    // Chart 7: Trendline
    var chart7 = try workbook.addChart(.line);
    try chart7.setTitle("Chart with a Trendline");
    var series7 = try chart7.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    _ = try chart7.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$C$2:$C$7",
    );

    var poly_line = ChartLine{
        .color = Colors.gray,
        .dash_type = .long_dash,
    };

    try series7.setTrendline(.poly, 3);
    try series7.setTrendlineLine(&poly_line);
    try worksheet.insertChart(97, 4, chart7);
}
