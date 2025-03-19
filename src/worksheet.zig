const std = @import("std");
const c = @import("xlsxwriter");
const format_mod = @import("format.zig");
const cell_utils = @import("cell_utils.zig");
const styled = @import("styled.zig");
const chart = @import("chart.zig");

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
        var iter = std.mem.splitScalar(u8, range, ':');
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

    /// Creates a StyledWriter starting at the specified position
    pub fn writer(self: *Worksheet, start_row: usize, start_col: usize, format: ?*format_mod.Format) styled.StyledWriter {
        return styled.StyledWriter.init(self, start_row, start_col, format);
    }

    /// Write a rich string to a cell, optionally with formatting
    pub fn writeRichString(
        self: *Worksheet,
        row: usize,
        col: usize,
        fragments: [*c][*c]c.lxw_rich_string_tuple,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;
        const result = c.worksheet_write_rich_string(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            fragments,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a rich string to a cell using a cell reference (e.g., "A1", "B2")
    pub fn writeRichStringCell(
        self: *Worksheet,
        cell_ref: []const u8,
        fragments: [*c][*c]c.lxw_rich_string_tuple,
        format: ?*format_mod.Format,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.writeRichString(pos.row, pos.col, fragments, format);
    }

    /// Write an array formula to a range of cells
    pub fn writeArrayFormula(
        self: *Worksheet,
        first_row: usize,
        first_col: usize,
        last_row: usize,
        last_col: usize,
        formula: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;

        // Ensure formula is null-terminated
        const null_term_formula = try self.workbook.allocator.dupeZ(u8, formula);
        defer self.workbook.allocator.free(null_term_formula);

        const result = c.worksheet_write_array_formula(
            self.worksheet,
            @intCast(first_row),
            @intCast(first_col),
            @intCast(last_row),
            @intCast(last_col),
            null_term_formula.ptr,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write an array formula to a range of cells using a cell range string (e.g., "A1:B2")
    pub fn writeArrayFormulaRange(
        self: *Worksheet,
        range: []const u8,
        formula: []const u8,
        format: ?*format_mod.Format,
    ) !void {
        // Split the range into start and end cells
        var iter = std.mem.splitScalar(u8, range, ':');
        const start = iter.next() orelse return error.InvalidRange;
        const end = iter.next() orelse return error.InvalidRange;

        // Parse start and end cells
        const start_pos = try cell_utils.cell.strToRowCol(start);
        const end_pos = try cell_utils.cell.strToRowCol(end);

        try self.writeArrayFormula(
            start_pos.row,
            start_pos.col,
            end_pos.row,
            end_pos.col,
            formula,
            format,
        );
    }

    /// Set the default row properties
    /// height: The height of the row in points
    /// hidden: If true, hide all rows that don't have data
    pub fn setDefaultRow(
        self: *Worksheet,
        height: f64,
        hidden: bool,
    ) void {
        _ = c.worksheet_set_default_row(
            self.worksheet,
            height,
            if (hidden) c.LXW_TRUE else c.LXW_FALSE,
        );
    }

    /// Hide a specific row
    pub fn hideRow(
        self: *Worksheet,
        row: u32,
    ) void {
        var options = c.lxw_row_col_options{
            .hidden = 1,
            .level = 0,
            .collapsed = 0,
        };
        _ = c.worksheet_set_row_opt(
            self.worksheet,
            row,
            15, // Default Excel row height
            null,
            &options,
        );
    }

    /// Set column properties with additional options
    /// first_col, last_col: Column range to format
    /// width: The width of the columns
    /// format: Optional formatting
    /// hidden: If true, hide the columns
    pub fn setColumnOpt(
        self: *Worksheet,
        first_col: u16,
        last_col: u16,
        width: f64,
        format: ?*format_mod.Format,
        hidden: bool,
    ) void {
        var options = c.lxw_row_col_options{
            .hidden = if (hidden) 1 else 0,
            .level = 0,
            .collapsed = 0,
        };

        const format_ptr = if (format) |f| f.format else null;

        _ = c.worksheet_set_column_opt(
            self.worksheet,
            first_col,
            last_col,
            width,
            format_ptr,
            &options,
        );
    }

    /// Set column properties using a column range string like "A:Z"
    /// column_range: Column range in string format (e.g., "A:Z", "C:D")
    /// width: The width of the columns
    /// format: Optional formatting
    /// hidden: If true, hide the columns
    pub fn setColumnOptRange(
        self: *Worksheet,
        column_range: []const u8,
        width: f64,
        format: ?*format_mod.Format,
        hidden: bool,
    ) !void {
        // Split the range into start and end columns
        var iter = std.mem.splitScalar(u8, column_range, ':');
        const start_col_str = iter.next() orelse return error.InvalidRange;
        const end_col_str = iter.next() orelse return error.InvalidRange;

        // Check if there are more parts than expected (i.e., more than one colon)
        if (iter.next() != null) return error.TooManyColonsInRange;

        // Convert column letters to indices
        const first_col = try cell_utils.cell.colToIndex(start_col_str);
        const last_col = try cell_utils.cell.colToIndex(end_col_str);

        self.setColumnOpt(first_col, last_col, width, format, hidden);
    }

    pub fn insertChart(self: *Worksheet, row: usize, col: usize, chart_obj: chart.Chart) !void {
        _ = c.worksheet_insert_chart(self.worksheet, @intCast(row), @intCast(col), chart_obj.inner);
    }

    /// Filter criteria for autofilter
    pub const FilterCriteria = enum {
        equal_to,
        not_equal_to,
        greater_than,
        less_than,
        greater_than_or_equal_to,
        less_than_or_equal_to,
        blanks,
        non_blanks,
    };

    /// Filter rule for autofilter
    pub const FilterRule = struct {
        criteria: FilterCriteria,
        value_string: ?[]const u8 = null,
        value: f64 = 0,
    };

    /// Filter operator for combining two filter rules
    pub const FilterOperator = enum {
        and_op,
        or_op,
    };

    /// Add an autofilter to a range of cells
    pub fn autofilter(
        self: *Worksheet,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) !void {
        const result = c.worksheet_autofilter(
            self.worksheet,
            first_row,
            first_col,
            last_row,
            last_col,
        );
        if (result != c.LXW_NO_ERROR) return error.AutofilterFailed;
    }

    /// Add a filter rule to a column
    pub fn filterColumn(
        self: *Worksheet,
        col: u16,
        rule: FilterRule,
    ) !void {
        var c_rule = c.lxw_filter_rule{
            .criteria = switch (rule.criteria) {
                .equal_to => c.LXW_FILTER_CRITERIA_EQUAL_TO,
                .not_equal_to => c.LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                .greater_than => c.LXW_FILTER_CRITERIA_GREATER_THAN,
                .less_than => c.LXW_FILTER_CRITERIA_LESS_THAN,
                .greater_than_or_equal_to => c.LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
                .less_than_or_equal_to => c.LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO,
                .blanks => c.LXW_FILTER_CRITERIA_BLANKS,
                .non_blanks => c.LXW_FILTER_CRITERIA_NON_BLANKS,
            },
            .value_string = if (rule.value_string) |s| s.ptr else null,
            .value = rule.value,
        };

        const result = c.worksheet_filter_column(
            self.worksheet,
            col,
            &c_rule,
        );
        if (result != c.LXW_NO_ERROR) return error.FilterColumnFailed;
    }

    /// Add two filter rules to a column with a specified operator
    pub fn filterColumn2(
        self: *Worksheet,
        col: u16,
        rule1: FilterRule,
        rule2: FilterRule,
        operator: FilterOperator,
    ) !void {
        var c_rule1 = c.lxw_filter_rule{
            .criteria = switch (rule1.criteria) {
                .equal_to => c.LXW_FILTER_CRITERIA_EQUAL_TO,
                .not_equal_to => c.LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                .greater_than => c.LXW_FILTER_CRITERIA_GREATER_THAN,
                .less_than => c.LXW_FILTER_CRITERIA_LESS_THAN,
                .greater_than_or_equal_to => c.LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
                .less_than_or_equal_to => c.LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO,
                .blanks => c.LXW_FILTER_CRITERIA_BLANKS,
                .non_blanks => c.LXW_FILTER_CRITERIA_NON_BLANKS,
            },
            .value_string = if (rule1.value_string) |s| s.ptr else null,
            .value = rule1.value,
        };

        var c_rule2 = c.lxw_filter_rule{
            .criteria = switch (rule2.criteria) {
                .equal_to => c.LXW_FILTER_CRITERIA_EQUAL_TO,
                .not_equal_to => c.LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
                .greater_than => c.LXW_FILTER_CRITERIA_GREATER_THAN,
                .less_than => c.LXW_FILTER_CRITERIA_LESS_THAN,
                .greater_than_or_equal_to => c.LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
                .less_than_or_equal_to => c.LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO,
                .blanks => c.LXW_FILTER_CRITERIA_BLANKS,
                .non_blanks => c.LXW_FILTER_CRITERIA_NON_BLANKS,
            },
            .value_string = if (rule2.value_string) |s| s.ptr else null,
            .value = rule2.value,
        };

        const result = c.worksheet_filter_column2(
            self.worksheet,
            col,
            &c_rule1,
            &c_rule2,
            switch (operator) {
                .and_op => c.LXW_FILTER_AND,
                .or_op => c.LXW_FILTER_OR,
            },
        );
        if (result != c.LXW_NO_ERROR) return error.FilterColumn2Failed;
    }

    /// Set the background image for the worksheet
    pub fn setBackground(
        self: *Worksheet,
        image_path: []const u8,
    ) !void {
        // Ensure the path is null-terminated
        const null_term_path = try self.workbook.allocator.dupeZ(u8, image_path);
        defer self.workbook.allocator.free(null_term_path);

        const result = c.worksheet_set_background(
            self.worksheet,
            null_term_path.ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.BackgroundFailed;
    }
};
