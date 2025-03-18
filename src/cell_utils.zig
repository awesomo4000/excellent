const std = @import("std");

/// Helper functions for column and cell references
pub const cell = struct {
    /// Convert a column string like "A", "B", "AA" to a zero-indexed column number
    pub fn colToIndex(col_str: []const u8) !u16 {
        if (col_str.len == 0) return error.InvalidColumn;

        var result: u16 = 0;
        for (col_str) |x| {
            const upper_c = std.ascii.toUpper(x);
            if (upper_c < 'A' or upper_c > 'Z') return error.InvalidColumn;

            result = result * 26 + @as(u16, upper_c - 'A' + 1);
        }

        return result - 1; // Convert to 0-indexed
    }

    /// Parse a cell reference like "A1", "B2", "AA10" into row and column indices
    pub fn strToRowCol(cell_ref: []const u8) !struct { row: u32, col: u16 } {
        if (cell_ref.len < 2) return error.InvalidCellReference;

        // Find the boundary between column letters and row number
        var i: usize = 0;
        while (i < cell_ref.len and isAlpha(cell_ref[i])) : (i += 1) {}

        if (i == 0 or i == cell_ref.len) return error.InvalidCellReference;

        const col_str = cell_ref[0..i];
        const row_str = cell_ref[i..];

        // Parse column letters
        const col = try colToIndex(col_str);

        // Parse row number (1-indexed in Excel, convert to 0-indexed)
        const row = try std.fmt.parseInt(u32, row_str, 10);
        if (row < 1) return error.InvalidCellReference;

        return .{ .row = row - 1, .col = col };
    }

    /// Convert a column index to an Excel column string (e.g., 0 -> "A", 25 -> "Z", 26 -> "AA")
    pub fn indexToCol(col_index: u16) ![]u8 {
        if (col_index > 16383) return error.ColumnOutOfRange; // Excel's limit

        var col_num = col_index + 1; // Convert to 1-indexed for the algorithm
        var result = std.ArrayList(u8).init(std.heap.page_allocator);
        defer result.deinit();

        while (col_num > 0) {
            const remainder = @mod(col_num - 1, 26);
            try result.append(@as(u8, @intCast(remainder)) + 'A');
            col_num = @divFloor(col_num - 1, 26);
        }

        // Reverse the string
        var i: usize = 0;
        var j: usize = result.items.len - 1;
        while (i < j) {
            const temp = result.items[i];
            result.items[i] = result.items[j];
            result.items[j] = temp;
            i += 1;
            j -= 1;
        }

        return result.toOwnedSlice();
    }

    // Helper function to check if a character is a letter
    fn isAlpha(c: u8) bool {
        return (c >= 'A' and c <= 'Z') or (c >= 'a' and c <= 'z');
    }
};
