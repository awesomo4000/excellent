const std = @import("std");
const excel = @import("excellent");
const colors = excel.Colors;

pub fn main() !void {
    // Create a new workbook
    var workbook =
        try excel.Workbook.create(
            std.heap.page_allocator,
            "chart_fonts.xlsx",
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

    // Create a chart object
    var chart = try workbook.addChart(.line);

    // Configure the chart series
    _ = try chart.addSeries(null, "Sheet1!$A$1:$A$6");

    // Create fonts for different chart elements
    const titleFont = excel.ChartFont{
        .name = "Calibri",
        .size = 18,
        .color = colors.blue,
    };

    const yAxisNameFont = excel.ChartFont{
        .name = "Courier",
        .color = colors.conifer,
    };

    const yAxisNumFont = excel.ChartFont{
        .name = "Arial",
        .color = colors.deepskyblue,
    };

    const xAxisNameFont = excel.ChartFont{
        .name = "Century",
        .color = colors.red,
    };

    const xAxisNumFont = excel.ChartFont{
        .rotation = -30,
    };

    const legendFont = excel.ChartFont{
        .color = colors.royalpurple,
        .bold = true,
        .italic = true,
        .underline = true,
    };

    // Configure chart title with font
    try chart.setTitle("Test Results");
    chart.setTitleFont(titleFont);

    // Configure Y axis with fonts
    try chart.setYAxisName("Units");
    chart.setYAxisNameFont(yAxisNameFont);
    chart.setYAxisNumFont(yAxisNumFont);

    // Configure X axis with fonts
    try chart.setXAxisName("Month");
    chart.setXAxisNameFont(xAxisNameFont);
    chart.setXAxisNumFont(xAxisNumFont);

    // Configure legend position and font
    chart.setLegendPosition(.bottom);
    chart.setLegendFont(legendFont);

    // Insert the chart into the worksheet
    try worksheet.insertChart(0, 2, chart);
}
