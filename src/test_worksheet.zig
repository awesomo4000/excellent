const std = @import("std");
const excellent = @import("excellent.zig");
const cell_utils = @import("cell_utils.zig");

test "setColumnOptRange validation" {
    const testing = std.testing;
    std.debug.print("\nRunning setColumnOptRange validation test...\n", .{});

    const MockWorksheet = struct {
        // Simplified version of setColumnOptRange that only validates the input
        pub fn validateColumnRange(column_range: []const u8) !void {
            // Split the range into start and end columns
            var iter = std.mem.splitScalar(u8, column_range, ':');
            const start_col_str = iter.next() orelse return error.InvalidRange;
            const end_col_str = iter.next() orelse return error.InvalidRange;

            // Check if there are more parts than expected (i.e., more than one colon)
            if (iter.next() != null) return error.TooManyColonsInRange;

            // Validate column names
            _ = try cell_utils.cell.colToIndex(start_col_str);
            _ = try cell_utils.cell.colToIndex(end_col_str);
        }
    };

    // Test case 1: Valid range with one colon
    try MockWorksheet.validateColumnRange("A:B");

    // Test case 2: Too many colons
    try testing.expectError(error.TooManyColonsInRange, MockWorksheet.validateColumnRange("A:B:C"));

    // Test case 3: Invalid column names
    try testing.expectError(error.InvalidColumn, MockWorksheet.validateColumnRange("123:XYZ"));

    // Test case 4: Empty range
    try testing.expectError(error.InvalidRange, MockWorksheet.validateColumnRange(""));

    std.debug.print("setColumnOptRange validation test completed successfully!\n", .{});
}

// Add more worksheet tests
test "column index conversion" {
    const testing = std.testing;

    // Test column index to string conversion
    const col_a = try cell_utils.cell.indexToCol(0);
    defer std.heap.page_allocator.free(col_a);
    try testing.expectEqualStrings("A", col_a);

    const col_z = try cell_utils.cell.indexToCol(25);
    defer std.heap.page_allocator.free(col_z);
    try testing.expectEqualStrings("Z", col_z);

    const col_aa = try cell_utils.cell.indexToCol(26);
    defer std.heap.page_allocator.free(col_aa);
    try testing.expectEqualStrings("AA", col_aa);
}
