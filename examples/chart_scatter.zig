const std = @import("std");
const excel = @import("excellent");

const Workbook = excel.Workbook;
const Worksheet = excel.Worksheet;
const Chart = excel.Chart;
const Format = excel.Format;
const ChartSeries = excel.ChartSeries;
const Colors = excel.Colors;

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try Workbook.create(
        allocator,
        "chart_scatter.xlsx",
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

    //
    // Chart 1: Create a basic scatter chart (points only)
    //
    var chart1 = try workbook.addChart(.scatter);

    // Add the first series to the chart
    var series = try chart1.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined
    // Configure the series using a programmatic approach
    series = try chart1.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");
    // Add chart title and axis labels
    try chart1.setTitle("Results of sample analysis");
    try chart1.setXaxisName("Test number");
    try chart1.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    _ = chart1.setStyle(11);

    // Insert the chart into the worksheet
    try worksheet.insertChart(1, 4, chart1);

    //
    // Chart 2: Create another scatter chart with a different style
    //
    var chart2 = try workbook.addChart(.scatter_straight_with_markers);

    // Add the first series to the chart
    series = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    series = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    // Add chart title and axis labels
    try chart2.setTitle("Results of sample analysis");
    try chart2.setXaxisName("Test number");
    try chart2.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    _ = chart2.setStyle(12);

    // Insert the chart into the worksheet
    try worksheet.insertChart(17, 4, chart2);

    //
    // Chart 3: Create a scatter chart with a different style
    //
    var chart3 = try workbook.addChart(.scatter_straight);

    // Add the first series to the chart
    series = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    series = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    // Add chart title and axis labels
    try chart3.setTitle("Results of sample analysis");
    try chart3.setXaxisName("Test number");
    try chart3.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    _ = chart3.setStyle(13);

    // Insert the chart into the worksheet
    try worksheet.insertChart(33, 4, chart3);

    //
    // Chart 4: Create a scatter chart with custom markers
    //
    var chart4 = try workbook.addChart(.scatter_smooth_with_markers);

    // Add the first series to the chart with custom markers
    series = try chart4.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add the second series with different markers
    series = try chart4.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    // Add chart title and axis labels
    try chart4.setTitle("Results of sample analysis");
    try chart4.setXaxisName("Test number");
    try chart4.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    _ = chart4.setStyle(14);

    // Insert the chart into the worksheet
    try worksheet.insertChart(49, 4, chart4);

    //
    // Chart 5: Create a scatter chart with colors
    //
    var chart5 = try workbook.addChart(.scatter_smooth);

    // Add the first series with custom colors
    series = try chart5.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    try series.setName("=Sheet1!$B$1");

    // Add the second series with different colors
    series = try chart5.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    try series.setName("=Sheet1!$C$1");

    // Add chart title and axis labels
    try chart5.setTitle("Results of sample analysis");
    try chart5.setXaxisName("Test number");
    try chart5.setYaxisName("Sample length (mm)");

    // Set an Excel chart style
    _ = chart5.setStyle(15);

    // Insert the chart into the worksheet
    try worksheet.insertChart(65, 4, chart5);
}

fn writeWorksheetData(worksheet: *Worksheet, bold: *Format) !void {
    const data = [_][3]f64{
        // Three columns of data
        [_]f64{ 2, 10, 30 },
        [_]f64{ 3, 40, 60 },
        [_]f64{ 4, 50, 70 },
        [_]f64{ 5, 20, 50 },
        [_]f64{ 6, 10, 40 },
        [_]f64{ 7, 50, 30 },
    };

    // Write the column headers
    try worksheet.writeString(0, 0, "Number", bold);
    try worksheet.writeString(0, 1, "Batch 1", bold);
    try worksheet.writeString(0, 2, "Batch 2", bold);

    // Write the example data
    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            try worksheet.writeNumber(
                @intCast(row_idx + 1),
                @intCast(col_idx),
                value,
                null,
            );
        }
    }
}
