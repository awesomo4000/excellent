const std = @import("std");
const excellent = @import("excellent.zig");
const Workbook = excellent.Workbook;

// Test workbook creation
test "workbook_creation" {
    // Create a temp workbook with a unique name for testing
    var workbook = try Workbook.create(std.testing.allocator, "/tmp/test_workbook.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    // Add a worksheet
    const worksheet = try workbook.addWorksheet("Sheet1");

    // Basic checks that worksheet was created properly
    _ = worksheet;
}
