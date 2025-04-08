const std = @import("std");
const excellent = @import("excellent");

pub fn main() !void {
    const allocator = std.heap.page_allocator;

    // Create a new workbook
    var workbook = try excellent.Workbook.create(allocator, "chart_styles.xlsx");
    defer workbook.deinit();

    // Define chart types and names
    const chart_types = [_]excellent.ChartType{
        .column,
        .area,
        .line,
        .pie,
    };
    const chart_names = [_][]const u8{ "Column", "Area", "Line", "Pie" };

    // Create a worksheet for each chart type
    for (chart_types, chart_names) |chart_type, chart_name| {
        var worksheet = try workbook.addWorksheet(chart_name);
        worksheet.setZoom(30); // Set zoom to 30% to show all charts

        // Create 48 charts, each with a different style
        var style_num: u8 = 1;
        var row_num: usize = 0;
        while (row_num < 90) : (row_num += 15) {
            var col_num: usize = 0;
            while (col_num < 64) : (col_num += 8) {
                var chart = try workbook.addChart(chart_type);
                // Note: No defer chart.deinit() here as the workbook owns the chart

                // Create chart title with style number
                var title_buf: [32]u8 = undefined;
                const title = try std.fmt.bufPrintZ(&title_buf, "Style {d}", .{style_num});
                try chart.setTitle(title);

                // Add series data
                _ = try chart.addSeries(null, "=Data!$A$1:$A$6");
                chart.setStyle(style_num);

                // Insert chart into worksheet
                try worksheet.insertChart(row_num, col_num, chart);

                style_num += 1;
            }
        }
    }

    // Create a worksheet with data for the charts
    var data_worksheet = try workbook.addWorksheet("Data");
    try data_worksheet.writeNumber(0, 0, 10, null);
    try data_worksheet.writeNumber(1, 0, 40, null);
    try data_worksheet.writeNumber(2, 0, 50, null);
    try data_worksheet.writeNumber(3, 0, 20, null);
    try data_worksheet.writeNumber(4, 0, 10, null);
    try data_worksheet.writeNumber(5, 0, 50, null);

    try workbook.close();
}
