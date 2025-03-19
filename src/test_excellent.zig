const std = @import("std");
const excellent = @import("excellent.zig");

// Test for column range validation
test "column_range_validation" {
    const testing = std.testing;

    // Simple validation helpers
    const validateOneColon = struct {
        fn check(str: []const u8) !void {
            const parts = std.mem.count(u8, str, ":");
            if (parts > 1) return error.TooManyColons;
        }
    };

    // These should pass
    try validateOneColon.check("A:B");
    try validateOneColon.check(":");
    try validateOneColon.check("AB:XY");

    // This should fail with TooManyColons
    try testing.expectError(error.TooManyColons, validateOneColon.check("A:B:C"));
}
