const std = @import("std");
const testing = std.testing;
const xlsxwriter = @import("xlsxwriter");
const c = xlsxwriter.c;
const workbook = @import("workbook.zig");
const format = @import("format.zig");
const worksheet = @import("worksheet.zig");

const Row = struct {
    region: []const u8,
    item: []const u8,
    volume: i32,
    month: []const u8,
};

const test_data = [_]Row{
    .{ .region = "East", .item = "Apple", .volume = 9000, .month = "July" },
    .{ .region = "East", .item = "Apple", .volume = 5000, .month = "July" },
    .{ .region = "South", .item = "Orange", .volume = 9000, .month = "September" },
    .{ .region = "North", .item = "Apple", .volume = 2000, .month = "November" },
    .{ .region = "West", .item = "Apple", .volume = 9000, .month = "November" },
};

fn writeWorksheetHeader(ws: *worksheet.Worksheet, header_format: *format.Format) !void {
    // Make the columns wider for clarity
    ws.setColumnWidth(0, 3, 12);

    // Write the column headers
    ws.setRowHeightFormat(0, 20, header_format);
    try ws.writeString(0, 0, "Region", header_format);
    try ws.writeString(0, 1, "Item", header_format);
    try ws.writeString(0, 2, "Volume", header_format);
    try ws.writeString(0, 3, "Month", header_format);
}

fn writeWorksheetData(ws: *worksheet.Worksheet) !void {
    for (test_data, 0..) |row, i| {
        try ws.writeString(@intCast(i + 1), 0, row.region, null);
        try ws.writeString(@intCast(i + 1), 1, row.item, null);
        try ws.writeNumber(@intCast(i + 1), 2, @floatFromInt(row.volume), null);
        try ws.writeString(@intCast(i + 1), 3, row.month, null);
    }
}

test "autofilter basic" {
    var wb = try workbook.Workbook.create("zig-test-autofilter.xlsx", null);
    defer wb.close();

    var ws = try wb.addWorksheet("Sheet1");
    var header_format = try wb.addFormat();
    try header_format.setBold(true);

    try writeWorksheetHeader(ws, header_format);
    try writeWorksheetData(ws);

    // Add autofilter to the data range
    try ws.autofilterRange("A1:D6");
}

test "autofilter with single filter" {
    var wb = try workbook.Workbook.create("zig-test-autofilter-single.xlsx", null);
    defer wb.close();

    var ws = try wb.addWorksheet("Sheet1");
    var header_format = try wb.addFormat();
    try header_format.setBold(true);

    try writeWorksheetHeader(ws, header_format);
    try writeWorksheetData(ws);

    // Add autofilter to the data range
    try ws.autofilterRange("A1:D6");

    // Add filter for East region
    try ws.filterColumn(0, .{
        .criteria = .equal_to,
        .value_string = "East",
    });
}

test "autofilter with dual filter" {
    var wb = try workbook.Workbook.create("zig-test-autofilter-dual.xlsx", null);
    defer wb.close();

    var ws = try wb.addWorksheet("Sheet1");
    var header_format = try wb.addFormat();
    try header_format.setBold(true);

    try writeWorksheetHeader(ws, header_format);
    try writeWorksheetData(ws);

    // Add autofilter to the data range
    try ws.autofilterRange("A1:D6");

    // Add filter for volume between 3000 and 8000
    try ws.filterColumn2(
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
}
