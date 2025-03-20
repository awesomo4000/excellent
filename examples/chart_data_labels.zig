const std = @import("std");
const excel = @import("excellent");
const colors = excel.Colors;
pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    // Create a new workbook and add a worksheet
    var workbook = try excel.Workbook.create(
        allocator,
        "chart_data_labels.xlsx",
    );
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    var worksheet = try workbook.addWorksheet("Sheet1");

    // Add a bold format for headers
    var bold_format = try workbook.addFormat();
    _ = bold_format.setBold();

    // Write some data for the chart
    try worksheet.writeString(0, 0, "Number", bold_format);
    try worksheet.writeNumber(1, 0, 2, null);
    try worksheet.writeNumber(2, 0, 3, null);
    try worksheet.writeNumber(3, 0, 4, null);
    try worksheet.writeNumber(4, 0, 5, null);
    try worksheet.writeNumber(5, 0, 6, null);
    try worksheet.writeNumber(6, 0, 7, null);

    try worksheet.writeString(0, 1, "Data", bold_format);
    try worksheet.writeNumber(1, 1, 20, null);
    try worksheet.writeNumber(2, 1, 10, null);
    try worksheet.writeNumber(3, 1, 20, null);
    try worksheet.writeNumber(4, 1, 30, null);
    try worksheet.writeNumber(5, 1, 40, null);
    try worksheet.writeNumber(6, 1, 30, null);

    try worksheet.writeString(0, 2, "Text", bold_format);
    try worksheet.writeString(1, 2, "Jan", null);
    try worksheet.writeString(2, 2, "Feb", null);
    try worksheet.writeString(3, 2, "Mar", null);
    try worksheet.writeString(4, 2, "Apr", null);
    try worksheet.writeString(5, 2, "May", null);
    try worksheet.writeString(6, 2, "Jun", null);

    // Chart 1: Standard data labels
    var chart1 = try workbook.addChart(.column);
    try chart1.setTitle("Chart with standard data labels");

    var series1 = try chart1.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add data labels to the series
    try series1.enableDataLabels();

    // Turn off the legend
    chart1.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(1, 3, chart1, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 2: Category and value data labels
    var chart2 = try workbook.addChart(.column);
    try chart2.setTitle("Category and Value data labels");

    var series2 = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add data labels with value and category options
    try series2.enableDataLabels();
    try series2.setDataLabelOptions(.{
        .show_category = true,
        .show_value = true,
    });

    // Turn off the legend
    chart2.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(17, 3, chart2, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 3: Data labels with custom font
    var chart3 = try workbook.addChart(.column);
    try chart3.setTitle("Data labels with user defined font");

    var series3 = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add data labels and customize the font
    try series3.enableDataLabels();
    try series3.setDataLabelFont(.{
        .name = "Arial",
        .bold = true,
        .color = colors.red,
        .rotation = -30,
    });

    // Turn off the legend
    chart3.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(33, 3, chart3, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 4: Data labels with formatting
    var chart4 = try workbook.addChart(.column);
    try chart4.setTitle("Data labels with formatting");

    var series4 = try chart4.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add data labels with border/line and fill
    try series4.enableDataLabels();
    try series4.setDataLabelLine(.{
        .color = colors.red,
        .width = 0.75,
    });
    try series4.setDataLabelFill(.{ .color = colors.yellow });

    // Turn off the legend
    chart4.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(49, 3, chart4, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 5: Custom string data labels
    var chart5 = try workbook.addChart(.column);
    try chart5.setTitle("Chart with custom string data labels");

    var series5 = try chart5.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Enable data labels
    try series5.enableDataLabels();

    // Add custom data labels
    try series5.setCustomDataLabels(&[_]?excel.ChartDataLabel{
        .{ .value = "Amy" },
        .{ .value = "Bea" },
        .{ .value = "Eva" },
        .{ .value = "Fay" },
        .{ .value = "Liv" },
        .{ .value = "Una" },
        null,
    });

    // Turn off the legend
    chart5.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(65, 3, chart5, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 6: Custom data labels from cells
    var chart6 = try workbook.addChart(.column);
    try chart6.setTitle("Chart with custom data labels from cells");

    var series6 = try chart6.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Enable data labels
    try series6.enableDataLabels();

    // Add custom data labels with cell references
    try series6.setCustomDataLabels(&[_]?excel.ChartDataLabel{
        .{ .value = "=Sheet1!$C$2" },
        .{ .value = "=Sheet1!$C$3" },
        .{ .value = "=Sheet1!$C$4" },
        .{ .value = "=Sheet1!$C$5" },
        .{ .value = "=Sheet1!$C$6" },
        .{ .value = "=Sheet1!$C$7" },
        null,
    });

    // Turn off the legend
    chart6.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(81, 3, chart6, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 7: Mixed custom and default data labels
    var chart7 = try workbook.addChart(.column);
    try chart7.setTitle("Mixed custom and default data labels");

    var series7 = try chart7.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Enable data labels
    try series7.enableDataLabels();

    // Create custom font
    const red_font = excel.ChartFont{
        .name = "Arial",
        .color = colors.red,
    };

    // Add mixed custom data labels
    try series7.setCustomDataLabels(&[_]?excel.ChartDataLabel{
        .{ .value = "=Sheet1!$C$2", .font = red_font },
        .{}, // Use default
        .{ .value = "=Sheet1!$C$4", .font = red_font },
        .{ .value = "=Sheet1!$C$5", .font = red_font },
        null,
    });

    // Turn off the legend
    chart7.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(97, 3, chart7, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 8: Hidden/deleted data labels
    var chart8 = try workbook.addChart(.column);
    try chart8.setTitle("Chart with deleted data labels");

    var series8 = try chart8.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Enable data labels
    try series8.enableDataLabels();

    // Hide specific data labels
    try series8.setCustomDataLabels(&[_]?excel.ChartDataLabel{
        .{ .hide = true },
        .{ .hide = false },
        .{ .hide = true },
        .{ .hide = true },
        .{ .hide = false },
        .{ .hide = true },
        null,
    });

    // Turn off the legend
    chart8.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(113, 3, chart8, .{
        .x_offset = 25,
        .y_offset = 10,
    });

    // Chart 9: Custom labels with formatting
    var chart9 = try workbook.addChart(.column);
    try chart9.setTitle("Chart with custom labels and formatting");

    var series9 = try chart9.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Enable data labels
    try series9.enableDataLabels();

    // Set default formatting for all labels
    try series9.setDataLabelLine(.{
        .color = colors.red,
        .width = 0.75,
    });
    try series9.setDataLabelFill(.{ .color = colors.yellow });

    // Add custom data labels with custom formatting that overrides defaults
    try series9.setCustomDataLabels(&[_]?excel.ChartDataLabel{
        .{ .value = "Amy", .line = .{
            .color = colors.blue,
            .width = 0.75,
        } },
        .{ .value = "Bea" },
        .{ .value = "Eva" },
        .{ .value = "Fay" },
        .{ .value = "Liv" },
        .{ .value = "Una", .fill = .{ .color = colors.green } },
        null,
    });

    // Turn off the legend
    chart9.setLegendPosition(.none);

    // Insert the chart into the worksheet with positioning options
    try worksheet.insertChartOpt(129, 3, chart9, .{
        .x_offset = 25,
        .y_offset = 10,
    });
}
