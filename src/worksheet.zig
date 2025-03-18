const std = @import("std");
const c = @import("xlsxwriter");
const format_mod = @import("format.zig");
const cell_utils = @import("cell_utils.zig");

/// Represents a worksheet within a workbook
pub const Worksheet = struct {
    workbook: *@import("workbook.zig").Workbook,
    worksheet: *c.lxw_worksheet,

    /// Write a string to a cell, optionally with formatting
    pub fn writeString(
        self: *Worksheet,
        row: usize,
        col: usize,
        text: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;

        // Ensure text is null-terminated
        const null_term_text = try self.workbook.allocator.dupeZ(u8, text);
        defer self.workbook.allocator.free(null_term_text);

        const result = c.worksheet_write_string(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            null_term_text.ptr,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a string to a cell using a cell reference (e.g., "A1", "B2")
    pub fn writeStringCell(
        self: *Worksheet,
        cell_ref: []const u8,
        text: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.writeString(pos.row, pos.col, text, format);
    }

    /// Write a number to a cell, optionally with formatting
    pub fn writeNumber(
        self: *Worksheet,
        row: usize,
        col: usize,
        number: f64,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;
        const result = c.worksheet_write_number(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            number,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a number to a cell using a cell reference (e.g., "A1", "B2")
    pub fn writeNumberCell(
        self: *Worksheet,
        cell_ref: []const u8,
        number: f64,
        format: ?*format_mod.Format,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.writeNumber(pos.row, pos.col, number, format);
    }

    /// Write a formula to a cell, optionally with formatting
    pub fn writeFormula(
        self: *Worksheet,
        row: usize,
        col: usize,
        formula: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;

        // Ensure formula is null-terminated
        const null_term_formula = try self.workbook.allocator.dupeZ(u8, formula);
        defer self.workbook.allocator.free(null_term_formula);

        const result = c.worksheet_write_formula(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            null_term_formula.ptr,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a formula to a cell using a cell reference (e.g., "A1", "B2")
    pub fn writeFormulaCell(
        self: *Worksheet,
        cell_ref: []const u8,
        formula: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.writeFormula(pos.row, pos.col, formula, format);
    }

    /// Set the column width
    pub fn setColumnWidth(
        self: *Worksheet,
        first_col: u16,
        last_col: u16,
        width: f64,
    ) void {
        _ = c.worksheet_set_column(
            self.worksheet,
            first_col,
            last_col,
            width,
            null,
        );
    }

    /// Set the column width with formatting
    pub fn setColumnWidthFormat(
        self: *Worksheet,
        first_col: u16,
        last_col: u16,
        width: f64,
        format: *format_mod.Format,
    ) void {
        _ = c.worksheet_set_column(
            self.worksheet,
            first_col,
            last_col,
            width,
            format.format,
        );
    }

    /// Set the row height
    pub fn setRowHeight(
        self: *Worksheet,
        row: u32,
        height: f64,
    ) void {
        _ = c.worksheet_set_row(
            self.worksheet,
            row,
            height,
            null,
        );
    }

    /// Set the row height with formatting
    pub fn setRowHeightFormat(
        self: *Worksheet,
        row: u32,
        height: f64,
        format: *format_mod.Format,
    ) void {
        _ = c.worksheet_set_row(
            self.worksheet,
            row,
            height,
            format.format,
        );
    }

    /// Merge a range of cells
    pub fn mergeRange(
        self: *Worksheet,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        content: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;
        const result = c.worksheet_merge_range(
            self.worksheet,
            first_row,
            first_col,
            last_row,
            last_col,
            content.ptr,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.MergeFailed;
    }

    /// Merge a range of cells using cell references (e.g., "A1:C3")
    pub fn mergeRangeRef(
        self: *Worksheet,
        range: []const u8,
        content: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        // Split the range into start and end cells
        var iter = std.mem.split(u8, range, ":");
        const start = iter.next() orelse return error.InvalidRange;
        const end = iter.next() orelse return error.InvalidRange;

        // Parse start and end cells
        const start_pos = try cell_utils.cell.strToRowCol(start);
        const end_pos = try cell_utils.cell.strToRowCol(end);

        try self.mergeRange(
            start_pos.row,
            start_pos.col,
            end_pos.row,
            end_pos.col,
            content,
            format,
        );
    }

    /// Write a datetime to a cell, optionally with formatting
    pub fn writeDateTime(
        self: *Worksheet,
        row: usize,
        col: usize,
        datetime: *c.lxw_datetime,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;
        const result = c.worksheet_write_datetime(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            datetime,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a datetime to a cell using a cell reference (e.g., "A1", "B2")
    pub fn writeDateTimeCell(
        self: *Worksheet,
        cell_ref: []const u8,
        datetime: *c.lxw_datetime,
        format: ?*format_mod.Format,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.writeDateTime(pos.row, pos.col, datetime, format);
    }
};
