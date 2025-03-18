const std = @import("std");
const Format = @import("format.zig").Format;
const Worksheet = @import("worksheet.zig").Worksheet;
const Workbook = @import("workbook.zig").Workbook;
const c = @import("xlsxwriter");

/// Represents text with associated formatting
pub const StyledText = struct {
    text: []const u8,
    style: ?*Format,

    pub fn init(text: []const u8, style: ?*Format) StyledText {
        return .{
            .text = text,
            .style = style,
        };
    }

    // Implement custom formatter for StyledText
    pub fn format(
        self: StyledText,
        comptime fmt: []const u8,
        options: std.fmt.FormatOptions,
        writer: anytype,
    ) !void {
        _ = fmt;
        _ = options;
        try writer.writeAll(self.text);
    }

    // Convenience constructors for common styles
    pub fn bold(workbook: *Workbook, text: []const u8) !StyledText {
        var style = try workbook.addFormat();
        _ = style.setBold();
        return StyledText.init(text, style);
    }

    pub fn italic(workbook: *Workbook, text: []const u8) !StyledText {
        var style = try workbook.addFormat();
        _ = style.setItalic();
        return StyledText.init(text, style);
    }

    pub fn colored(workbook: *Workbook, text: []const u8, color: u32) !StyledText {
        var style = try workbook.addFormat();
        _ = style.setFontColor(color);
        return StyledText.init(text, style);
    }

    pub fn boldItalic(workbook: *Workbook, text: []const u8) !StyledText {
        var style = try workbook.addFormat();
        _ = style.setBold().setItalic();
        return StyledText.init(text, style);
    }

    pub fn withBgColor(workbook: *Workbook, text: []const u8, color: u32) !StyledText {
        var style = try workbook.addFormat();
        _ = style.setBgColor(color);
        _ = style.setPattern(c.LXW_PATTERN_SOLID);
        _ = style.setFontColor(0x000000);
        _ = style.setBorder(.thin);
        return StyledText.init(text, style);
    }
};

/// A writer that supports writing formatted text to a worksheet
pub const StyledWriter = struct {
    worksheet: *Worksheet,
    current_row: usize,
    current_col: usize,
    default_format: ?*Format,

    pub fn init(worksheet: *Worksheet, start_row: usize, start_col: usize, format: ?*Format) StyledWriter {
        return .{
            .worksheet = worksheet,
            .current_row = start_row,
            .current_col = start_col,
            .default_format = format,
        };
    }

    pub fn print(self: *StyledWriter, comptime fmt: []const u8, args: anytype) !void {
        var buf: [1024]u8 = undefined;
        const text = try std.fmt.bufPrint(&buf, fmt, args);

        // Write to worksheet with appropriate formatting
        try self.worksheet.writeString(self.current_row, self.current_col, text, self.default_format);
        self.current_col += 1;
    }

    /// Special print method for handling StyledText
    pub fn printStyled(self: *StyledWriter, comptime fmt: []const u8, args: anytype) !void {
        const ArgsType = @TypeOf(args);
        const fields = std.meta.fields(ArgsType);
        var fragments = std.ArrayList(c.lxw_rich_string_tuple).init(self.worksheet.workbook.allocator);
        defer fragments.deinit();

        // First, calculate total string length needed
        var total_len: usize = 0;
        var current_pos: usize = 0;
        inline for (fields) |field| {
            const value = @field(args, field.name);
            const pos = std.mem.indexOf(u8, fmt[current_pos..], "{s}") orelse break;
            if (pos > 0) {
                total_len += pos + 1; // +1 for null terminator
            }
            current_pos += pos + 3;
            if (@TypeOf(value) == StyledText) {
                total_len += value.text.len + 1;
            } else {
                total_len += value.len + 1;
            }
        }
        if (current_pos < fmt.len) {
            total_len += fmt.len - current_pos + 1;
        }

        // Allocate a single buffer for all strings and initialize with null terminators
        var buffer = try self.worksheet.workbook.allocator.alloc(u8, total_len);
        defer self.worksheet.workbook.allocator.free(buffer);
        @memset(buffer, 0);

        // Reset position for second pass
        current_pos = 0;
        var buffer_pos: usize = 0;

        // Process the format string and build rich string fragments
        inline for (fields) |field| {
            const value = @field(args, field.name);
            const pos = std.mem.indexOf(u8, fmt[current_pos..], "{s}") orelse break;

            // Add text before placeholder as unformatted fragment
            if (pos > 0) {
                const text = buffer[buffer_pos .. buffer_pos + pos :0];
                @memcpy(text[0..pos], fmt[current_pos .. current_pos + pos]);
                try fragments.append(.{
                    .format = null,
                    .string = text.ptr,
                });
                buffer_pos += pos + 1;
            }
            current_pos += pos + 3; // Skip the {s} placeholder

            // Add the value as a formatted fragment
            if (@TypeOf(value) == StyledText) {
                const text = buffer[buffer_pos .. buffer_pos + value.text.len :0];
                @memcpy(text[0..value.text.len], value.text);
                const format_ptr = if (value.style) |s| s.format else null;
                try fragments.append(.{
                    .format = format_ptr,
                    .string = text.ptr,
                });
                buffer_pos += value.text.len + 1;
            } else {
                const text = buffer[buffer_pos .. buffer_pos + value.len :0];
                @memcpy(text[0..value.len], value);
                try fragments.append(.{
                    .format = null,
                    .string = text.ptr,
                });
                buffer_pos += value.len + 1;
            }
        }

        // Add any remaining text after the last placeholder
        if (current_pos < fmt.len) {
            const remaining_len = fmt.len - current_pos;
            const text = buffer[buffer_pos .. buffer_pos + remaining_len :0];
            @memcpy(text[0..remaining_len], fmt[current_pos..]);
            try fragments.append(.{
                .format = null,
                .string = text.ptr,
            });
        }

        // Create a properly aligned array of pointers to fragments
        var fragment_ptrs = try self.worksheet.workbook.allocator.alloc([*c]c.lxw_rich_string_tuple, fragments.items.len + 1);
        defer self.worksheet.workbook.allocator.free(fragment_ptrs);

        for (fragments.items, 0..) |*fragment, i| {
            fragment_ptrs[i] = @ptrCast(fragment);
        }
        fragment_ptrs[fragments.items.len] = null; // Null terminate the array like in the C examples

        // Write the rich string to the worksheet
        try self.worksheet.writeRichString(
            self.current_row,
            self.current_col,
            fragment_ptrs.ptr,
            self.default_format, // Pass the cell format
        );

        // Move to next column
        self.current_col += 1;
    }

    pub fn withFormat(self: *StyledWriter, format: ?*Format) StyledWriter {
        return .{
            .worksheet = self.worksheet,
            .current_row = self.current_row,
            .current_col = self.current_col,
            .default_format = format,
        };
    }

    pub fn nextRow(self: *StyledWriter) void {
        self.current_row += 1;
        self.current_col = 0;
    }
};
