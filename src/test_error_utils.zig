const std = @import("std");
const excellent = @import("excellent.zig");
const XlsxError = excellent.XlsxError;

// Test error utils
test "xlsx_error_values" {
    // Test error creation and comparison
    const memory_error = XlsxError.MemoryError;
    try std.testing.expectEqual(XlsxError.MemoryError, memory_error);

    const format_error = XlsxError.FormatError;
    try std.testing.expectEqual(XlsxError.FormatError, format_error);
}
