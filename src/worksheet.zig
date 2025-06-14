const std = @import("std");
const c = @import("xlsxwriter");
const format_mod = @import("format.zig");
const cell_utils = @import("cell_utils.zig");
const styled = @import("styled.zig");
const chart = @import("chart.zig");
const comment = @import("comment.zig");
const cf = @import("conditional_format.zig");
const DateTime = @import("date_time.zig").DateTime;

/// Represents a worksheet within a workbook
pub const Worksheet = struct {
    workbook: *@import("workbook.zig").Workbook,
    worksheet: *c.lxw_worksheet,

    pub fn deinit(self: *Worksheet) void {
        // The worksheet is owned by the workbook, so we don't need to free it
        // But we should clear any resources we've allocated
        self.worksheet = undefined;
        self.workbook = undefined;
    }

    /// Set the zoom level for the worksheet (10-400)
    pub fn setZoom(self: *Worksheet, zoom: u16) void {
        _ = c.worksheet_set_zoom(self.worksheet, zoom);
    }

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
        datetime: DateTime,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;
        var datetime_mut = datetime.toC(); // use to avoid const cast err
        const result = c.worksheet_write_datetime(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            &datetime_mut,
            format_ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a datetime to a cell using a cell reference (e.g., "A1", "B2")
    pub fn writeDateTimeCell(
        self: *Worksheet,
        cell_ref: []const u8,
        datetime: DateTime,
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

    /// Write a unix time to a cell, optionally with formatting
    pub fn writeUnixTime(
        self: *Worksheet,
        row: usize,
        col: usize,
        unix_time: i64,
        format: ?*format_mod.Format,
    ) !void {
        const format_ptr = if (format) |f| f.format else null;
        const result = c.worksheet_write_unixtime(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            unix_time,
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

    /// Insert a chart into the worksheet at the specified position
    pub fn insertChart(
        self: *Worksheet,
        row: usize,
        col: usize,
        chart_obj: *chart.Chart,
    ) !void {
        const result = c.worksheet_insert_chart(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            chart_obj.inner,
        );
        if (result != c.LXW_NO_ERROR) return error.InsertChartFailed;
    }

    /// Chart options for chart positioning
    pub const ChartOptions = struct {
        x_offset: i32 = 0,
        y_offset: i32 = 0,
        x_scale: f32 = 1.0,
        y_scale: f32 = 1.0,
    };

    /// Insert a chart into the worksheet at the specified position with options
    pub fn insertChartOpt(
        self: *Worksheet,
        row: usize,
        col: usize,
        chart_obj: *chart.Chart,
        options: ChartOptions,
    ) !void {
        var c_options = c.lxw_chart_options{
            .x_offset = options.x_offset,
            .y_offset = options.y_offset,
            .x_scale = options.x_scale,
            .y_scale = options.y_scale,
        };

        const result = c.worksheet_insert_chart_opt(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            chart_obj.inner,
            &c_options,
        );
        if (result != c.LXW_NO_ERROR) return error.ChartInsertFailed;
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

    /// Write a comment to a cell
    pub fn writeComment(
        self: *Worksheet,
        row: usize,
        col: usize,
        comment_text: []const u8,
    ) !void {
        // Ensure comment text is null-terminated
        const null_term_comment = try self.workbook.allocator.dupeZ(u8, comment_text);
        defer self.workbook.allocator.free(null_term_comment);

        const result = c.worksheet_write_comment(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            null_term_comment.ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a comment to a cell using a cell reference (e.g., "A1", "B2")
    pub fn writeCommentCell(
        self: *Worksheet,
        cell_ref: []const u8,
        comment_text: []const u8,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.writeComment(pos.row, pos.col, comment_text);
    }

    /// Write a comment to a cell with options
    pub fn writeCommentOpt(
        self: *Worksheet,
        row: usize,
        col: usize,
        comment_text: []const u8,
        options: comment.CommentOptions,
    ) !void {
        // Ensure comment text is null-terminated
        const null_term_comment = try self.workbook.allocator.dupeZ(u8, comment_text);
        defer self.workbook.allocator.free(null_term_comment);

        // Convert options to C library format
        var c_options = try options.toCOptions(self.workbook.allocator);

        // Store author pointer so we can free it later if needed
        const author_ptr = c_options.author;

        const result = c.worksheet_write_comment_opt(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            null_term_comment.ptr,
            &c_options,
        );

        // Free the author string if it was allocated
        if (author_ptr != null) {
            self.workbook.allocator.free(std.mem.sliceTo(author_ptr, 0));
        }

        if (result != c.LXW_NO_ERROR) return error.WriteFailed;
    }

    /// Write a comment to a cell with options using a cell reference (e.g., "A1", "B2")
    pub fn writeCommentOptCell(
        self: *Worksheet,
        cell_ref: []const u8,
        comment_text: []const u8,
        options: comment.CommentOptions,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.writeCommentOpt(pos.row, pos.col, comment_text, options);
    }

    /// Show all comments in the worksheet
    pub fn showComments(self: *Worksheet) void {
        c.worksheet_show_comments(self.worksheet);
    }

    /// Hide all comments on the worksheet
    pub fn hideComments(self: *Worksheet) void {
        _ = c.worksheet_hide_comments(self.worksheet);
    }

    /// Insert an image into the worksheet at the specified position
    pub fn insertImage(
        self: *Worksheet,
        row: usize,
        col: usize,
        filename: []const u8,
    ) !void {
        // Ensure filename is null-terminated
        const null_term_filename = try self.workbook.allocator.dupeZ(u8, filename);
        defer self.workbook.allocator.free(null_term_filename);

        const result = c.worksheet_insert_image(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            null_term_filename.ptr,
        );
        if (result != c.LXW_NO_ERROR) return error.InsertImageFailed;
    }

    /// Button options for inserting buttons
    pub const ButtonOptions = struct {
        caption: []const u8,
        macro: []const u8,
        width: u16 = 80,
        height: u16 = 30,
    };

    /// Insert a button into the worksheet
    pub fn insertButton(
        self: *Worksheet,
        row: usize,
        col: usize,
        options: ButtonOptions,
    ) !void {
        // Ensure strings are null-terminated
        const c_caption = try self.workbook.allocator.dupeZ(u8, options.caption);
        defer self.workbook.allocator.free(c_caption);

        const c_macro = try self.workbook.allocator.dupeZ(u8, options.macro);
        defer self.workbook.allocator.free(c_macro);

        var c_options = c.lxw_button_options{
            .caption = c_caption.ptr,
            .macro = c_macro.ptr,
            .width = options.width,
            .height = options.height,
        };

        const result = c.worksheet_insert_button(
            self.worksheet,
            @intCast(row),
            @intCast(col),
            &c_options,
        );
        if (result != c.LXW_NO_ERROR) return error.InsertButtonFailed;
    }

    /// Insert a button into the worksheet using a cell reference (e.g., "A1", "B2")
    pub fn insertButtonCell(
        self: *Worksheet,
        cell_ref: []const u8,
        options: ButtonOptions,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        try self.insertButton(pos.row, pos.col, options);
    }

    pub fn conditionalFormatRange(
        self: *Worksheet,
        range: []const u8,
        format: cf.ConditionalFormat,
    ) !void {
        const pos = try cell_utils.rangeToPositions(range);
        const conditional_format_ptr = try self.workbook.addConditionalFormat(format);
        const result = c.worksheet_conditional_format_range(
            self.worksheet,
            pos.start_row,
            pos.start_col,
            pos.end_row,
            pos.end_col,
            &conditional_format_ptr.inner,
        );
        if (result != c.LXW_NO_ERROR) return error.ConditionalFormatFailed;
    }

    pub fn conditionalFormatCell(
        self: *Worksheet,
        cell_ref: []const u8,
        format: cf.ConditionalFormat,
    ) !void {
        const pos = try cell_utils.cell.strToRowCol(cell_ref);
        const conditional_format_ptr: *cf.ConditionalFormat =
            try self.workbook.addConditionalFormat(format);

        const result = c.worksheet_conditional_format_cell(
            self.worksheet,
            pos.row,
            pos.col,
            &conditional_format_ptr.inner,
        );
        if (result != c.LXW_NO_ERROR) return error.ConditionalFormatFailed;
    }
};
