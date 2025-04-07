const std = @import("std");
const excel = @import("excellent");
const Workbook = excel.Workbook;
const Worksheet = excel.Worksheet;
const Format = excel.Format;
const Chart = excel.Chart;
const ChartType = excel.ChartType;

// Write some data to the worksheet
fn writeWorksheetData(worksheet: *Worksheet, bold: *Format) !void {
    const data = [_][3]u8{
        // Three columns of data
        [_]u8{ 2, 10, 30 },
        [_]u8{ 3, 40, 60 },
        [_]u8{ 4, 50, 70 },
        [_]u8{ 5, 20, 50 },
        [_]u8{ 6, 10, 40 },
        [_]u8{ 7, 50, 30 },
    };

    try worksheet.writeString(0, 0, "Number", bold);
    try worksheet.writeString(0, 1, "Batch 1", bold);
    try worksheet.writeString(0, 2, "Batch 2", bold);

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
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    var workbook = try Workbook.create(allocator, "chart_line.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet("Sheet1");
    defer worksheet.deinit();

    // Add a bold format to use to highlight the header cells
    var bold = try workbook.addFormat();
    defer bold.deinit();
    _ = bold.setBold();

    // Write some data for the chart
    try writeWorksheetData(&worksheet, bold);

    //
    // Chart 1. Create a line chart
    //
    var chart1 = try workbook.addChart(.line);

    // Add the first series to the chart
    var series = try chart1.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );
    try series.setName("=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined
    series = try chart1.addSeries(null, null);

    // Configure the series using a syntax that is easier to define programmatically
    try series.setCategories("Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
    try series.setValues("Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
    try series.setNameRange("Sheet1", 0, 2); // "=Sheet1!$C$1"

    // Add a chart title and some axis labels
    try chart1.setTitle("Results of sample analysis");
    try chart1.setXaxisName("Test number");
    try chart1.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    chart1.setStyle(10);

    // Insert the chart into the worksheet
    try worksheet.insertChart(1, 4, chart1);

    //
    // Chart 2. Create a stacked line chart.
    //
    var chart2 = try workbook.addChart(.line_stacked);

    // Add the first series to the chart
    series = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    series = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    // Add a chart title and some axis labels
    try chart2.setTitle("Results of sample analysis");
    try chart2.setXaxisName("Test number");
    try chart2.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    chart2.setStyle(12);

    // Insert the chart into the worksheet
    try worksheet.insertChart(17, 4, chart2);

    //
    // Chart 3. Create a percent stacked line chart.
    //
    var chart3 = try workbook.addChart(.line_stacked_percent);

    // Add the first series to the chart
    series = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    series = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    // Add a chart title and some axis labels
    try chart3.setTitle("Results of sample analysis");
    try chart3.setXaxisName("Test number");
    try chart3.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    chart3.setStyle(13);

    // Insert the chart into the worksheet
    try worksheet.insertChart(33, 4, chart3);
}
