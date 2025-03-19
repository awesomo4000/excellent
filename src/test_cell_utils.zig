const std = @import("std");
const excellent = @import("excellent.zig");
const cell = excellent.cell;

// Test column index conversion
test "column_index_conversion" {
    // Test col to index
    try std.testing.expectEqual(@as(u16, 0), try cell.colToIndex("A"));
    try std.testing.expectEqual(@as(u16, 25), try cell.colToIndex("Z"));
    try std.testing.expectEqual(@as(u16, 26), try cell.colToIndex("AA"));
    try std.testing.expectEqual(@as(u16, 701), try cell.colToIndex("ZZ"));

    // Test invalid column
    try std.testing.expectError(error.InvalidColumn, cell.colToIndex("123"));
    try std.testing.expectError(error.InvalidColumn, cell.colToIndex(""));
}

// Test cell reference parsing
test "cell_reference_parsing" {
    // Test valid references
    const a1 = try cell.strToRowCol("A1");
    try std.testing.expectEqual(@as(u32, 0), a1.row);
    try std.testing.expectEqual(@as(u16, 0), a1.col);

    const b2 = try cell.strToRowCol("B2");
    try std.testing.expectEqual(@as(u32, 1), b2.row);
    try std.testing.expectEqual(@as(u16, 1), b2.col);

    // Test invalid references
    try std.testing.expectError(error.InvalidCellReference, cell.strToRowCol("A"));
    try std.testing.expectError(error.InvalidCellReference, cell.strToRowCol("123"));
}
