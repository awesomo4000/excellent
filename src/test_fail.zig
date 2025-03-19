const std = @import("std");

// Deliberately failing test
test "failing_test" {
    // This should fail
    try std.testing.expect(false);
}
