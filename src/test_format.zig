const std = @import("std");
const excellent = @import("excellent.zig");
const Format = excellent.Format;
const Alignment = excellent.Alignment;
const BorderStyle = excellent.BorderStyle;
const Workbook = excellent.Workbook;

// Test Format creation and properties
test "format_creation" {
    // Create a workbook first, since Format objects are created from a workbook
    var workbook = try Workbook.create(std.testing.allocator, "/tmp/test_workbook.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    // Create a format from the workbook
    var format = try workbook.addFormat();

    // Test setting properties
    _ = format.setBold();
    _ = format.setItalic();

    // Set alignment
    _ = format.setAlign(.Center);

    // Set border
    _ = format.setBorder(.thin);
}
