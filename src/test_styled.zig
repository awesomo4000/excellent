const std = @import("std");
const excellent = @import("excellent.zig");
const StyledText = excellent.StyledText;
const StyledWriter = excellent.StyledWriter;
const Workbook = excellent.Workbook;

// Test styled text creation
test "styled_text_creation" {
    // Create a workbook first
    var workbook = try Workbook.create(std.testing.allocator, "/tmp/test_styled.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    // Get a format
    var format = try workbook.addFormat();
    _ = format.setBold();

    // Create styled text
    const styled_text = StyledText.init("Test", format);

    // Check properties
    try std.testing.expectEqualStrings("Test", styled_text.text);
}
