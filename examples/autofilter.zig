const std = @import("std");
const excellent = @import("excellent");

const Row = struct {
    region: []const u8,
    item: []const u8,
    volume: i32,
    month: []const u8,
};

const data = [_]Row{
    .{ .region = "East", .item = "Apple", .volume = 9000, .month = "July" },
    .{ .region = "East", .item = "Apple", .volume = 5000, .month = "July" },
    .{ .region = "South", .item = "Orange", .volume = 9000, .month = "September" },
    .{ .region = "North", .item = "Apple", .volume = 2000, .month = "November" },
    .{ .region = "West", .item = "Apple", .volume = 9000, .month = "November" },
    .{ .region = "South", .item = "Pear", .volume = 7000, .month = "October" },
    .{ .region = "North", .item = "Pear", .volume = 9000, .month = "August" },
    .{ .region = "West", .item = "Orange", .volume = 1000, .month = "December" },
    .{ .region = "West", .item = "Grape", .volume = 1000, .month = "November" },
    .{ .region = "South", .item = "Pear", .volume = 10000, .month = "April" },
    .{ .region = "West", .item = "Grape", .volume = 6000, .month = "January" },
    .{ .region = "South", .item = "Orange", .volume = 3000, .month = "May" },
    .{ .region = "North", .item = "Apple", .volume = 3000, .month = "December" },
    .{ .region = "South", .item = "Apple", .volume = 7000, .month = "February" },
    .{ .region = "West", .item = "Grape", .volume = 1000, .month = "December" },
    .{ .region = "East", .item = "Grape", .volume = 8000, .month = "February" },
    .{ .region = "South", .item = "Grape", .volume = 10000, .month = "June" },
    .{ .region = "West", .item = "Pear", .volume = 7000, .month = "December" },
    .{ .region = "South", .item = "Apple", .volume = 2000, .month = "October" },
    .{ .region = "East", .item = "Grape", .volume = 7000, .month = "December" },
    .{ .region = "North", .item = "Grape", .volume = 6000, .month = "April" },
    .{ .region = "East", .item = "Pear", .volume = 8000, .month = "February" },
    .{ .region = "North", .item = "Apple", .volume = 7000, .month = "August" },
    .{ .region = "North", .item = "Orange", .volume = 7000, .month = "July" },
    .{ .region = "North", .item = "Apple", .volume = 6000, .month = "June" },
    .{ .region = "South", .item = "Grape", .volume = 8000, .month = "September" },
    .{ .region = "West", .item = "Apple", .volume = 3000, .month = "October" },
    .{ .region = "South", .item = "Orange", .volume = 10000, .month = "November" },
    .{ .region = "West", .item = "Grape", .volume = 4000, .month = "July" },
    .{ .region = "North", .item = "Orange", .volume = 5000, .month = "August" },
    .{ .region = "East", .item = "Orange", .volume = 1000, .month = "November" },
    .{ .region = "East", .item = "Orange", .volume = 4000, .month = "October" },
    .{ .region = "North", .item = "Grape", .volume = 5000, .month = "August" },
    .{ .region = "East", .item = "Apple", .volume = 1000, .month = "December" },
    .{ .region = "South", .item = "Apple", .volume = 10000, .month = "March" },
    .{ .region = "East", .item = "Grape", .volume = 7000, .month = "October" },
    .{ .region = "West", .item = "Grape", .volume = 1000, .month = "September" },
    .{ .region = "East", .item = "Grape", .volume = 10000, .month = "October" },
    .{ .region = "South", .item = "Orange", .volume = 8000, .month = "March" },
    .{ .region = "North", .item = "Apple", .volume = 4000, .month = "July" },
    .{ .region = "South", .item = "Orange", .volume = 5000, .month = "July" },
    .{ .region = "West", .item = "Apple", .volume = 4000, .month = "June" },
    .{ .region = "East", .item = "Apple", .volume = 5000, .month = "April" },
    .{ .region = "North", .item = "Pear", .volume = 3000, .month = "August" },
    .{ .region = "East", .item = "Grape", .volume = 9000, .month = "November" },
    .{ .region = "North", .item = "Orange", .volume = 8000, .month = "October" },
    .{ .region = "East", .item = "Apple", .volume = 10000, .month = "June" },
    .{ .region = "South", .item = "Pear", .volume = 1000, .month = "December" },
    .{ .region = "North", .item = "Grape", .volume = 10000, .month = "July" },
    .{ .region = "East", .item = "Grape", .volume = 6000, .month = "February" },
};

fn writeWorksheetHeader(ws: *excellent.Worksheet, header_format: *excellent.Format) !void {
    // Make the columns wider for clarity
    ws.setColumnWidth(0, 3, 12);

    // Write the column headers
    ws.setRowHeightFormat(0, 20, header_format);
    try ws.writeString(0, 0, "Region", header_format);
    try ws.writeString(0, 1, "Item", header_format);
    try ws.writeString(0, 2, "Volume", header_format);
    try ws.writeString(0, 3, "Month", header_format);
}

fn writeWorksheetData(ws: *excellent.Worksheet) !void {
    for (data, 0..) |row, i| {
        try ws.writeString(@intCast(i + 1), 0, row.region, null);
        try ws.writeString(@intCast(i + 1), 1, row.item, null);
        try ws.writeNumber(@intCast(i + 1), 2, @floatFromInt(row.volume), null);
        try ws.writeString(@intCast(i + 1), 3, row.month, null);
    }
}

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    var wb = try excellent.Workbook.create(allocator, "autofilter.xlsx");
    defer wb.deinit();

    // Create worksheets for each example
    var ws1 = try wb.addWorksheet("Sheet1");
    var ws2 = try wb.addWorksheet("Sheet2");
    var ws3 = try wb.addWorksheet("Sheet3");
    var ws4 = try wb.addWorksheet("Sheet4");
    var ws5 = try wb.addWorksheet("Sheet5");
    var ws6 = try wb.addWorksheet("Sheet6");
    var ws7 = try wb.addWorksheet("Sheet7");

    var header_format = try wb.addFormat();
    _ = header_format.setBold();

    // Example 1: Basic autofilter
    try writeWorksheetHeader(&ws1, header_format);
    try writeWorksheetData(&ws1);
    try ws1.autofilter(0, 0, 50, 3);

    // Example 2: Autofilter with a single filter
    try writeWorksheetHeader(&ws2, header_format);
    try writeWorksheetData(&ws2);
    // Hide rows that don't match the filter
    for (data, 0..) |row, i| {
        if (!std.mem.eql(u8, row.region, "East")) {
            ws2.hideRow(@intCast(i + 1));
        }
    }
    try ws2.autofilter(0, 0, 50, 3);
    try ws2.filterColumn(0, .{
        .criteria = .equal_to,
        .value_string = "East",
    });

    // Example 3: Autofilter with dual filter
    try writeWorksheetHeader(&ws3, header_format);
    try writeWorksheetData(&ws3);
    // Hide rows that don't match the filter
    for (data, 0..) |row, i| {
        if (!std.mem.eql(u8, row.region, "East") and !std.mem.eql(u8, row.region, "South")) {
            ws3.hideRow(@intCast(i + 1));
        }
    }
    try ws3.autofilter(0, 0, 50, 3);
    try ws3.filterColumn2(
        0,
        .{
            .criteria = .equal_to,
            .value_string = "East",
        },
        .{
            .criteria = .equal_to,
            .value_string = "South",
        },
        .or_op,
    );

    // Example 4: Autofilter with filter conditions in two columns
    try writeWorksheetHeader(&ws4, header_format);
    try writeWorksheetData(&ws4);
    // Hide rows that don't match the filter
    for (data, 0..) |row, i| {
        if (!(std.mem.eql(u8, row.region, "East") and row.volume > 3000 and row.volume < 8000)) {
            ws4.hideRow(@intCast(i + 1));
        }
    }
    try ws4.autofilter(0, 0, 50, 3);
    try ws4.filterColumn(0, .{
        .criteria = .equal_to,
        .value_string = "East",
    });
    try ws4.filterColumn2(
        2,
        .{
            .criteria = .greater_than,
            .value = 3000,
        },
        .{
            .criteria = .less_than,
            .value = 8000,
        },
        .and_op,
    );

    // Example 5: Autofilter with a list filter condition
    try writeWorksheetHeader(&ws5, header_format);
    try writeWorksheetData(&ws5);
    // Hide rows that don't match the filter
    for (data, 0..) |row, i| {
        if (!std.mem.eql(u8, row.region, "East") and
            !std.mem.eql(u8, row.region, "North") and
            !std.mem.eql(u8, row.region, "South"))
        {
            ws5.hideRow(@intCast(i + 1));
        }
    }
    try ws5.autofilter(0, 0, 50, 3);
    try ws5.filterColumn(0, .{
        .criteria = .equal_to,
        .value_string = "East",
    });

    // Example 6: Autofilter with filter for blanks
    try writeWorksheetHeader(&ws6, header_format);

    // Create a copy of data with one blank region
    var data_with_blank = data;
    data_with_blank[5].region = "";

    // Write the row data
    for (data_with_blank, 0..) |row, i| {
        try ws6.writeString(@intCast(i + 1), 0, row.region, null);
        try ws6.writeString(@intCast(i + 1), 1, row.item, null);
        try ws6.writeNumber(@intCast(i + 1), 2, @floatFromInt(row.volume), null);
        try ws6.writeString(@intCast(i + 1), 3, row.month, null);
        // Hide rows that don't match the filter
        if (!std.mem.eql(u8, row.region, "")) {
            ws6.hideRow(@intCast(i + 1));
        }
    }
    try ws6.autofilter(0, 0, 50, 3);
    try ws6.filterColumn(0, .{ .criteria = .blanks });

    // Example 7: Autofilter with filter for non-blanks
    try writeWorksheetHeader(&ws7, header_format);

    // Write the row data using the same data with blank
    for (data_with_blank, 0..) |row, i| {
        try ws7.writeString(@intCast(i + 1), 0, row.region, null);
        try ws7.writeString(@intCast(i + 1), 1, row.item, null);
        try ws7.writeNumber(@intCast(i + 1), 2, @floatFromInt(row.volume), null);
        try ws7.writeString(@intCast(i + 1), 3, row.month, null);
        // Hide rows that don't match the filter
        if (std.mem.eql(u8, row.region, "")) {
            ws7.hideRow(@intCast(i + 1));
        }
    }
    try ws7.autofilter(0, 0, 50, 3);
    try ws7.filterColumn(0, .{ .criteria = .non_blanks });
}
