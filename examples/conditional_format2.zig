const std = @import("std");
const excel = @import("excellent");
const Worksheet = excel.Worksheet;
const Workbook = excel.Workbook;
const Format = excel.Format;
const Colors = excel.Colors;
const cf = excel.cf; // conditional formatting

// Write some data to the worksheet.
fn writeWorksheetData(worksheet: *Worksheet) !void {
    const data = [10][10]u8{
        [_]u8{ 34, 72, 38, 30, 75, 48, 75, 66, 84, 86 },
        [_]u8{ 6, 24, 1, 84, 54, 62, 60, 3, 26, 59 },
        [_]u8{ 28, 79, 97, 13, 85, 93, 93, 22, 5, 14 },
        [_]u8{ 27, 71, 40, 17, 18, 79, 90, 93, 29, 47 },
        [_]u8{ 88, 25, 33, 23, 67, 1, 59, 79, 47, 36 },
        [_]u8{ 24, 100, 20, 88, 29, 33, 38, 54, 54, 88 },
        [_]u8{ 6, 57, 88, 28, 10, 26, 37, 7, 41, 48 },
        [_]u8{ 52, 78, 1, 96, 26, 45, 47, 33, 96, 36 },
        [_]u8{ 60, 54, 81, 66, 81, 90, 80, 93, 12, 55 },
        [_]u8{ 70, 5, 46, 14, 71, 19, 66, 36, 41, 21 },
    };

    for (0..10) |row| {
        for (0..10) |col| {
            try worksheet.writeNumber(
                @intCast(row + 2),
                @intCast(col + 1),
                @floatFromInt(data[row][col]),
                null,
            );
        }
    }
}

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    const workbook = try Workbook.create(
        allocator,
        "conditional_format2.xlsx",
    );
    defer workbook.deinit();

    var worksheet1 = try workbook.addWorksheet(null);
    var worksheet2 = try workbook.addWorksheet(null);
    var worksheet3 = try workbook.addWorksheet(null);
    var worksheet4 = try workbook.addWorksheet(null);
    var worksheet5 = try workbook.addWorksheet(null);
    var worksheet6 = try workbook.addWorksheet(null);
    var worksheet7 = try workbook.addWorksheet(null);
    var worksheet8 = try workbook.addWorksheet(null);
    // var worksheet9 = try workbook.addWorksheet(null);

    // Add a format. Light red fill with dark red text.
    const redFormat = try workbook.addFormat();
    _ = redFormat.setBgColor(
        Colors.lightsalmon,
    ).setFontColor(Colors.deepcrimson);

    // Add a format. Green fill with dark green text.
    const greenFormat = try workbook.addFormat();
    _ = greenFormat.setBgColor(
        Colors.grannyapple,
    ).setFontColor(Colors.mosscore);

    // Example 1. Conditional formatting based on simple cell based criteria.
    try writeWorksheetData(&worksheet1);

    try worksheet1.writeStringCell(
        "A1",
        "Cells with values >= 50 are in light red." ++
            " Values < 50 are in light green.",
        null,
    );

    try worksheet1.conditionalFormatRange(
        "B3:K12",
        cf.cellGreaterThanOrEqualTo(50, redFormat),
    );
    try worksheet1.conditionalFormatRange(
        "B3:K12",
        cf.cellLessThan(50, greenFormat),
    );

    // // Example 2. Conditional formatting based on max and min values.
    try writeWorksheetData(&worksheet2);

    try worksheet2.writeStringCell(
        "A1",
        "Values between 30 and 70 are in light red. " ++
            "Values outside that range are in light green.",
        null,
    );

    try worksheet2.conditionalFormatRange(
        "B3:K12",
        cf.cellBetween(30, 70, redFormat),
    );
    try worksheet2.conditionalFormatRange(
        "B3:K12",
        cf.cellNotBetween(30, 70, greenFormat),
    );

    // // Example 3. Conditional formatting with duplicate and unique values.
    try writeWorksheetData(&worksheet3);

    try worksheet3.writeString(
        0,
        0,
        "Duplicate values are in light red. " ++
            "Unique values are in light green.",
        null,
    );

    try worksheet3.conditionalFormatRange(
        "B3:K12",
        cf.duplicate(redFormat),
    );
    try worksheet3.conditionalFormatRange(
        "B3:K12",
        cf.unique(greenFormat),
    );

    // // Example 4. Conditional formatting with above and below average values.
    try writeWorksheetData(&worksheet4);

    try worksheet4.writeString(
        0,
        0,
        "Above average values are in light red. " ++
            "Below average values are in light green.",
        null,
    );

    try worksheet4.conditionalFormatRange(
        "B3:K12",
        cf.averageAbove(redFormat),
    );
    try worksheet4.conditionalFormatRange(
        "B3:K12",
        cf.averageBelow(greenFormat),
    );

    // // Example 5. Conditional formatting with top and bottom values.
    try writeWorksheetData(&worksheet5);

    try worksheet5.writeString(
        0,
        0,
        "Top 10 values are in light red. " ++
            "Bottom 10 values are in light green.",
        null,
    );

    try worksheet5.conditionalFormatRange(
        "B3:K12",
        cf.top(10, redFormat),
    );
    try worksheet5.conditionalFormatRange(
        "B3:K12",
        cf.bottom(10, greenFormat),
    );

    // Example 6. Conditional formatting with multiple ranges.
    try writeWorksheetData(&worksheet6);

    try worksheet6.writeString(
        0,
        0,
        "Cells with values >= 50 are in light red. " ++
            "Values < 50 are in light green. " ++
            "Non-contiguous ranges.",
        null,
    );
    const condFmtGte50 = cf.cellGreaterThanOrEqualTo(50, redFormat);
    const condFmtLt50 = cf.cellLessThan(50, greenFormat);
    try worksheet6.conditionalFormatRange("B3:K6", condFmtGte50);
    try worksheet6.conditionalFormatRange("B9:K12", condFmtGte50);
    try worksheet6.conditionalFormatRange("B3:K6", condFmtLt50);
    try worksheet6.conditionalFormatRange("B9:K12", condFmtLt50);

    // Example 7. Conditional formatting with 2 color scales.
    // Write the worksheet data.
    for (1..13) |i| {
        try worksheet7.writeNumber(@intCast(i + 1), 1, @floatFromInt(i), null);
        try worksheet7.writeNumber(@intCast(i + 1), 3, @floatFromInt(i), null);
        try worksheet7.writeNumber(@intCast(i + 1), 6, @floatFromInt(i), null);
        try worksheet7.writeNumber(@intCast(i + 1), 8, @floatFromInt(i), null);
    }

    try worksheet7.writeString(
        0,
        0,
        "Examples of color scales with default and user colors.",
        null,
    );

    try worksheet7.writeStringCell("B2", "2 Color Scale", null);
    try worksheet7.writeStringCell("D2", "2 Color Scale + user colors", null);
    try worksheet7.writeStringCell("G2", "3 Color Scale", null);
    try worksheet7.writeStringCell("I2", "3 Color Scale + user colors", null);

    // 2 color scale with standard colors.
    try worksheet7.conditionalFormatRange("B3:B14", cf.twoColorScale(null));

    // 2 color scale with user defined colors.
    try worksheet7.conditionalFormatRange("D3:D14", cf.twoColorScale(.{
        .min_color = Colors.red,
        .max_color = Colors.lime,
    }));

    // 3 color scale with standard colors.
    try worksheet7.conditionalFormatRange("G3:G14", cf.threeColorScale(null));

    // 3 color scale with user defined colors.
    try worksheet7.conditionalFormatRange("I3:I14", cf.threeColorScale(.{
        .min_color = Colors.fadedcyan,
        .mid_color = Colors.muffinsky,
        .max_color = Colors.grumbleblue,
    }));

    // Example 8. Conditional formatting with data bars.
    // First data bar example.
    try worksheet8.writeString(
        0,
        0,
        "Examples of data bars.",
        null,
    );

    // // Write the worksheet data.
    for (1..13) |i| {
        try worksheet8.writeNumber(@intCast(i + 1), 1, @floatFromInt(i), null);
        try worksheet8.writeNumber(@intCast(i + 1), 3, @floatFromInt(i), null);
        try worksheet8.writeNumber(@intCast(i + 1), 5, @floatFromInt(i), null);
        try worksheet8.writeNumber(@intCast(i + 1), 7, @floatFromInt(i), null);
        try worksheet8.writeNumber(@intCast(i + 1), 9, @floatFromInt(i), null);
        try worksheet8.writeNumber(@intCast(i + 1), 11, @floatFromInt(i), null);
        try worksheet8.writeNumber(@intCast(i + 1), 13, @floatFromInt(i), null);
    }

    try worksheet8.writeStringCell("B2", "Default data bars", null);
    try worksheet8.writeStringCell("D2", "Data bars + border", null);
    try worksheet8.writeStringCell("F2", "Bars with user color", null);
    try worksheet8.writeStringCell("H2", "Negative same as positive", null);
    try worksheet8.writeStringCell("J2", "Zero axis", null);
    try worksheet8.writeStringCell("L2", "Right to left", null);
    try worksheet8.writeStringCell("N2", "Excel 2010 style", null);

    var sheet8condFmt = cf.dataBar(.{});

    // Default data bars.
    try worksheet8.conditionalFormatRange("B3:B14", sheet8condFmt);

    // Data bars with border.
    sheet8condFmt.bar_border_color = Colors.black;
    try worksheet8.conditionalFormatRange("D3:D14", sheet8condFmt);

    // User defined color.
    sheet8condFmt.bar_color = Colors.wiltedlettuce;
    try worksheet8.conditionalFormatRange("F3:F14", sheet8condFmt);

    // Same color for negative values.
    sheet8condFmt.bar_negative_color_same = true;
    try worksheet8.conditionalFormatRange("H3:H14", sheet8condFmt);

    // Zero axis.
    sheet8condFmt.bar_axis_position = .midpoint;
    sheet8condFmt.bar_axis_color = 0;
    try worksheet8.conditionalFormatRange("J3:J14", sheet8condFmt);

    // Right to left.
    sheet8condFmt.bar_axis_position = .right;
    try worksheet8.conditionalFormatRange("L3:L14", cf.dataBar(null));

    // Excel 2010 style.
    try worksheet8.conditionalFormatRange("N3:N14", cf.dataBar(null));

    // // Example 9. Conditional formatting with icon sets.
    // try worksheet9.writeString(
    //     0,
    //     0,
    //     "Examples of conditional formats with icon sets.",
    //     null,
    // );

    // // Write the worksheet data.
    // for (1..4) |i| {
    //     try worksheet9.writeNumber(
    //         2,
    //         @intCast(i),
    //         @floatFromInt(i),
    //         null,
    //     );
    //     try worksheet9.writeNumber(
    //         3,
    //         @intCast(i),
    //         @floatFromInt(i),
    //         null,
    //     );
    //     try worksheet9.writeNumber(
    //         4,
    //         @intCast(i),
    //         @floatFromInt(i),
    //         null,
    //     );
    //     try worksheet9.writeNumber(
    //         5,
    //         @intCast(i),
    //         @floatFromInt(i),
    //         null,
    //     );
    // }

    // for (1..5) |i| {
    //     try worksheet9.writeNumber(6, @intCast(i), @floatFromInt(i), null);
    //     try worksheet9.writeNumber(7, @intCast(i), @floatFromInt(i), null);
    //     try worksheet9.writeNumber(8, @intCast(i), @floatFromInt(i), null);
    // }

    // try worksheet9.writeString(1, 1, "3 Traffic lights", null);
    // try worksheet9.writeString(1, 3, "3 Traffic lights unrimmed", null);
    // try worksheet9.writeString(1, 5, "3 Arrows", null);
    // try worksheet9.writeString(1, 7, "3 Symbols circled", null);
    // try worksheet9.writeString(1, 9, "3 Symbols", null);
    // try worksheet9.writeString(1, 11, "3 Flags", null);
    // try worksheet9.writeString(1, 13, "3 Traffic lights", null);

    // // Three traffic lights (default style).
    // try worksheet9.conditionalFormatRange("B3:B5", cf.iconSet(
    //     null,
    //     .{ .icon_style = .three_traffic_lights_rimmed },
    // ));

    // // Three traffic lights (unrimmed style).
    // try worksheet9.conditionalFormatRange(
    //     "D3:D5",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .three_traffic_lights_unrimmed },
    //     ),
    // );

    // // Three arrows.
    // try worksheet9.conditionalFormatRange(
    //     "F3:F5",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .three_arrows_colored },
    //     ),
    // );

    // // Three symbols circled.
    // try worksheet9.conditionalFormatRange(
    //     "H3:H5",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .three_symbols_circled },
    //     ),
    // );

    // // Three symbols.
    // try worksheet9.conditionalFormatRange(
    //     "J3:J5",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .three_symbols_uncircled },
    //     ),
    // );

    // // Three flags.
    // try worksheet9.conditionalFormatRange(
    //     "L3:L5",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .three_flags },
    //     ),
    // );

    // // Three traffic lights.
    // try worksheet9.conditionalFormatRange(
    //     "N3:N5",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .three_traffic_lights_rimmed },
    //     ),
    // );

    // // Examples of 4 set icons.
    // try worksheet9.writeString(1, 15, "4 Traffic lights", null);
    // try worksheet9.writeString(1, 17, "4 Arrows", null);
    // try worksheet9.writeString(1, 19, "4 Red-Black", null);
    // try worksheet9.writeString(1, 21, "4 Ratings", null);

    // // Four traffic lights.
    // try worksheet9.conditionalFormatRange(
    //     "P3:P6",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .four_traffic_lights },
    //     ),
    // );

    // // Four arrows.
    // try worksheet9.conditionalFormatRange(
    //     "R3:R6",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .four_arrows_colored },
    //     ),
    // );

    // // Four red to black.
    // try worksheet9.conditionalFormatRange(
    //     "T3:T6",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .four_red_to_black },
    //     ),
    // );

    // // Four ratings.
    // try worksheet9.conditionalFormatRange(
    //     "V3:V6",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .four_ratings },
    //     ),
    // );

    // // Examples of 5 set icons.
    // try worksheet9.writeString(6, 15, "5 Arrows", null);
    // try worksheet9.writeString(6, 17, "5 Ratings", null);
    // try worksheet9.writeString(6, 19, "5 Quarters", null);

    // // Five arrows.
    // try worksheet9.conditionalFormatRange(
    //     "P7:P11",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .five_arrows_colored },
    //     ),
    // );

    // // Five ratings.
    // try worksheet9.conditionalFormatRange(
    //     "R7:R11",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .five_ratings },
    //     ),
    // );

    // // Five quarters.
    // try worksheet9.conditionalFormatRange(
    //     "T7:T11",
    //     cf.iconSet(
    //         null,
    //         .{ .icon_style = .five_quarters },
    //     ),
    // );

    try workbook.close();
}
