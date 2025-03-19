const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    // Create a new workbook and add a worksheet
    var workbook = try excel.Workbook.create(
        allocator,
        "test_column_range_error.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // This should work fine
    std.debug.print("Testing valid column range 'A:B'...\n", .{});
    worksheet.setColumnOptRange("A:B", 10, null, false) catch |err| {
        std.debug.print("Unexpected error with valid range: {}\n", .{err});
        return;
    };
    std.debug.print("Valid range test passed.\n", .{});

    // This should fail with TooManyColonsInRange
    std.debug.print("\nTesting invalid column range 'A:B:C'...\n", .{});
    worksheet.setColumnOptRange("A:B:C", 10, null, false) catch |err| {
        if (err == error.TooManyColonsInRange) {
            std.debug.print("Correctly caught error: TooManyColonsInRange\n", .{});
        } else {
            std.debug.print("Unexpected error type: {}\n", .{err});
        }
        return;
    };

    std.debug.print("ERROR: Expected to catch an error with invalid range 'A:B:C', but no error occurred.\n", .{});
}
