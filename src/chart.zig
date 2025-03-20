const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

// Export colors for easy use
pub const Colors = struct {
    pub const BLACK: u32 = xlsxwriter.LXW_COLOR_BLACK;
    pub const BLUE: u32 = xlsxwriter.LXW_COLOR_BLUE;
    pub const BROWN: u32 = xlsxwriter.LXW_COLOR_BROWN;
    pub const CYAN: u32 = xlsxwriter.LXW_COLOR_CYAN;
    pub const GRAY: u32 = xlsxwriter.LXW_COLOR_GRAY;
    pub const GREEN: u32 = xlsxwriter.LXW_COLOR_GREEN;
    pub const LIME: u32 = xlsxwriter.LXW_COLOR_LIME;
    pub const MAGENTA: u32 = xlsxwriter.LXW_COLOR_MAGENTA;
    pub const NAVY: u32 = xlsxwriter.LXW_COLOR_NAVY;
    pub const ORANGE: u32 = xlsxwriter.LXW_COLOR_ORANGE;
    pub const PINK: u32 = xlsxwriter.LXW_COLOR_PINK;
    pub const PURPLE: u32 = xlsxwriter.LXW_COLOR_PURPLE;
    pub const RED: u32 = xlsxwriter.LXW_COLOR_RED;
    pub const SILVER: u32 = xlsxwriter.LXW_COLOR_SILVER;
    pub const WHITE: u32 = xlsxwriter.LXW_COLOR_WHITE;
    pub const YELLOW: u32 = xlsxwriter.LXW_COLOR_YELLOW;
};

pub const ChartType = enum {
    column,
    bar,
    bar_stacked,
    bar_stacked_percent,
    line,
    pie,
    scatter,
    area,
    area_stacked,
    area_stacked_percent,
    radar,
    doughnut,
    column_stacked,
    column_stacked_percent,

    fn toNative(self: ChartType) u8 {
        return switch (self) {
            .column => @intCast(xlsxwriter.LXW_CHART_COLUMN),
            .bar => @intCast(xlsxwriter.LXW_CHART_BAR),
            .bar_stacked => @intCast(xlsxwriter.LXW_CHART_BAR_STACKED),
            .bar_stacked_percent => @intCast(xlsxwriter.LXW_CHART_BAR_STACKED_PERCENT),
            .line => @intCast(xlsxwriter.LXW_CHART_LINE),
            .pie => @intCast(xlsxwriter.LXW_CHART_PIE),
            .scatter => @intCast(xlsxwriter.LXW_CHART_SCATTER),
            .area => @intCast(xlsxwriter.LXW_CHART_AREA),
            .area_stacked => @intCast(xlsxwriter.LXW_CHART_AREA_STACKED),
            .area_stacked_percent => @intCast(xlsxwriter.LXW_CHART_AREA_STACKED_PERCENT),
            .radar => @intCast(xlsxwriter.LXW_CHART_RADAR),
            .doughnut => @intCast(xlsxwriter.LXW_CHART_DOUGHNUT),
            .column_stacked => @intCast(xlsxwriter.LXW_CHART_COLUMN_STACKED),
            .column_stacked_percent => @intCast(xlsxwriter.LXW_CHART_COLUMN_STACKED_PERCENT),
        };
    }
};

pub const ChartFont = struct {
    name: []const u8 = "Arial",
    size: f64 = 10.0,
    bold: bool = false,
    italic: bool = false,
    color: u32 = xlsxwriter.LXW_COLOR_BLACK,
    rotation: i16 = 0,

    fn toNative(self: ChartFont) xlsxwriter.lxw_chart_font {
        return .{
            .name = @ptrCast(self.name),
            .size = self.size,
            .bold = if (self.bold) xlsxwriter.LXW_TRUE else xlsxwriter.LXW_FALSE,
            .italic = if (self.italic) xlsxwriter.LXW_TRUE else xlsxwriter.LXW_FALSE,
            .color = self.color,
            .rotation = self.rotation,
            .underline = 0,
            .charset = 0,
            .pitch_family = 0,
            .baseline = 0,
        };
    }
};

pub const ChartLine = struct {
    color: u32 = xlsxwriter.LXW_COLOR_BLACK,
    width: f32 = 2.25,
    dash_type: u8 = xlsxwriter.LXW_CHART_LINE_DASH_SOLID,

    fn toNative(self: ChartLine) xlsxwriter.lxw_chart_line {
        return .{
            .color = self.color,
            .width = self.width,
            .dash_type = self.dash_type,
            .transparency = 0,
            .none = 0,
        };
    }
};

pub const ChartFill = struct {
    color: u32 = xlsxwriter.LXW_COLOR_BLACK,
    transparency: u8 = 0,

    fn toNative(self: ChartFill) xlsxwriter.lxw_chart_fill {
        return .{
            .color = self.color,
            .transparency = self.transparency,
            .none = 0,
        };
    }
};

// Define data label options for series
pub const DataLabelOptions = struct {
    show_name: bool = false,
    show_category: bool = false,
    show_value: bool = false,
    show_series_name: bool = false,
    show_percent: bool = false,
    show_leader_lines: bool = false,
    show_legend_key: bool = false,
};

// Define data label struct for custom data labels
pub const ChartDataLabel = struct {
    value: ?[]const u8 = null,
    font: ?ChartFont = null,
    line: ?ChartLine = null,
    fill: ?ChartFill = null,
    hide: bool = false,
};

pub const ChartLegendPosition = enum {
    none,
    right,
    left,
    top,
    bottom,

    fn toNative(self: ChartLegendPosition) u8 {
        return switch (self) {
            .none => xlsxwriter.LXW_CHART_LEGEND_NONE,
            .right => xlsxwriter.LXW_CHART_LEGEND_RIGHT,
            .left => xlsxwriter.LXW_CHART_LEGEND_LEFT,
            .top => xlsxwriter.LXW_CHART_LEGEND_TOP,
            .bottom => xlsxwriter.LXW_CHART_LEGEND_BOTTOM,
        };
    }
};

pub const ChartAxis = enum {
    x_axis,
    y_axis,
};

pub const ChartSeries = struct {
    inner: *xlsxwriter.lxw_chart_series,
    allocator: std.mem.Allocator,
    strings: std.ArrayList([]const u8),
    data_labels_enabled: bool = false,

    pub fn setName(self: *ChartSeries, name: []const u8) !void {
        const name_str = try self.allocator.dupeZ(u8, name);
        try self.strings.append(name_str[0..name_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_name(self.inner, name_str);
    }

    pub fn setCategories(self: *ChartSeries, sheet: []const u8, first_row: u32, first_col: u16, last_row: u32, last_col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.strings.append(sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_categories(self.inner, sheet_str, first_row, first_col, last_row, last_col);
    }

    pub fn setValues(self: *ChartSeries, sheet: []const u8, first_row: u32, first_col: u16, last_row: u32, last_col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.strings.append(sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_values(self.inner, sheet_str, first_row, first_col, last_row, last_col);
    }

    pub fn setNameRange(self: *ChartSeries, sheet: []const u8, row: u32, col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.strings.append(sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_name_range(self.inner, sheet_str, row, col);
    }

    /// Enable data labels for this series
    pub fn enableDataLabels(self: *ChartSeries) !void {
        _ = xlsxwriter.chart_series_set_labels(self.inner);
        self.data_labels_enabled = true;
    }

    /// Set data label options (which elements to show)
    pub fn setDataLabelOptions(self: *ChartSeries, options: DataLabelOptions) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        _ = xlsxwriter.chart_series_set_labels_options(
            self.inner,
            if (options.show_name) 1 else 0,
            if (options.show_category) 1 else 0,
            if (options.show_value) 1 else 0,
        );

        // TODO: Add additional options in a separate function if needed
        // Currently the C API only supports the first 3 options
    }

    /// Set font for data labels
    pub fn setDataLabelFont(self: *ChartSeries, font: ChartFont) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        var native_font = font.toNative();
        _ = xlsxwriter.chart_series_set_labels_font(self.inner, &native_font);
    }

    /// Set line/border properties for data labels
    pub fn setDataLabelLine(self: *ChartSeries, line: ChartLine) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        var native_line = line.toNative();
        _ = xlsxwriter.chart_series_set_labels_line(self.inner, &native_line);
    }

    /// Set fill/background properties for data labels
    pub fn setDataLabelFill(self: *ChartSeries, fill: ChartFill) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        var native_fill = fill.toNative();
        _ = xlsxwriter.chart_series_set_labels_fill(self.inner, &native_fill);
    }

    /// Set custom data labels for this series
    pub fn setCustomDataLabels(self: *ChartSeries, labels: []const ?ChartDataLabel) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        var native_labels = try self.allocator.alloc(?*xlsxwriter.lxw_chart_data_label, labels.len + 1);
        defer self.allocator.free(native_labels);

        var native_data_labels = try self.allocator.alloc(xlsxwriter.lxw_chart_data_label, labels.len);
        defer self.allocator.free(native_data_labels);

        // Keep track of allocated strings, fonts, lines, fills to free later
        var value_strings = std.ArrayList([]const u8).init(self.allocator);
        defer value_strings.deinit();

        var native_fonts = std.ArrayList(xlsxwriter.lxw_chart_font).init(self.allocator);
        defer native_fonts.deinit();

        var native_lines = std.ArrayList(xlsxwriter.lxw_chart_line).init(self.allocator);
        defer native_lines.deinit();

        var native_fills = std.ArrayList(xlsxwriter.lxw_chart_fill).init(self.allocator);
        defer native_fills.deinit();

        // Set the null terminator at the end of the array
        native_labels[labels.len] = null;

        // Convert each label
        for (labels, 0..) |maybe_label, i| {
            if (maybe_label) |label| {
                native_data_labels[i] = .{};

                // Set value if provided
                if (label.value) |value| {
                    const value_str = try self.allocator.dupeZ(u8, value);
                    try value_strings.append(value_str[0..value_str.len]); // Store for cleanup
                    native_data_labels[i].value = value_str;
                    try self.strings.append(value_str[0..value_str.len]); // Store for deinit
                }

                // Set font if provided
                if (label.font) |font| {
                    const native_font = font.toNative();
                    try native_fonts.append(native_font);
                    native_data_labels[i].font = &native_fonts.items[native_fonts.items.len - 1];
                }

                // Set line if provided
                if (label.line) |line| {
                    const native_line = line.toNative();
                    try native_lines.append(native_line);
                    native_data_labels[i].line = &native_lines.items[native_lines.items.len - 1];
                }

                // Set fill if provided
                if (label.fill) |fill| {
                    const native_fill = fill.toNative();
                    try native_fills.append(native_fill);
                    native_data_labels[i].fill = &native_fills.items[native_fills.items.len - 1];
                }

                // Set hide state
                native_data_labels[i].hide = if (label.hide) 1 else 0;

                // Add the label to the array
                native_labels[i] = &native_data_labels[i];
            } else {
                native_labels[i] = null;
            }
        }

        _ = xlsxwriter.chart_series_set_labels_custom(self.inner, &native_labels[0]);
    }

    fn deinit(self: *ChartSeries) void {
        for (self.strings.items) |str| {
            var owned_str = str; // Make a mutable copy
            owned_str.len += 1; // Include null terminator
            self.allocator.free(owned_str);
        }
        self.strings.deinit();
    }
};

pub const Chart = struct {
    allocator: std.mem.Allocator,
    series: std.ArrayList(*ChartSeries),
    series_strings: std.ArrayList([]const u8),
    // Inner C object, but not exposed in the public API
    chart_inner: *xlsxwriter.lxw_chart,

    pub fn init(allocator: std.mem.Allocator, workbook: *xlsxwriter.lxw_workbook, chart_type: ChartType) !Chart {
        const inner = xlsxwriter.workbook_add_chart(workbook, chart_type.toNative()) orelse {
            return error.ChartCreationFailed;
        };
        return Chart{
            .allocator = allocator,
            .series = std.ArrayList(*ChartSeries).init(allocator),
            .series_strings = std.ArrayList([]const u8).init(allocator),
            .chart_inner = inner,
        };
    }

    pub fn addSeries(self: *Chart, categories: ?[]const u8, values: ?[]const u8) !*ChartSeries {
        const cat_ptr: ?[*:0]const u8 = if (categories) |c| blk: {
            const str = try self.allocator.dupeZ(u8, c);
            errdefer self.allocator.free(str);
            try self.series_strings.append(str[0..str.len]); // Store without null terminator
            break :blk str;
        } else null;

        const val_ptr: ?[*:0]const u8 = if (values) |v| blk: {
            const str = try self.allocator.dupeZ(u8, v);
            errdefer self.allocator.free(str);
            try self.series_strings.append(str[0..str.len]); // Store without null terminator
            break :blk str;
        } else null;

        const series_inner = xlsxwriter.chart_add_series(self.chart_inner, if (cat_ptr) |c| @ptrCast(c) else null, if (val_ptr) |v| @ptrCast(v) else null);

        const series = try self.allocator.create(ChartSeries);
        series.* = .{
            .inner = series_inner,
            .allocator = self.allocator,
            .strings = std.ArrayList([]const u8).init(self.allocator),
        };

        try self.series.append(series);
        return series;
    }

    pub fn setTitle(self: *Chart, title: []const u8) !void {
        const title_str = try self.allocator.dupeZ(u8, title);
        errdefer self.allocator.free(title_str);
        try self.series_strings.append(title_str[0..title_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_title_set_name(self.chart_inner, title_str);
    }

    pub fn setTitleFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = xlsxwriter.chart_title_set_name_font(self.chart_inner, &native_font);
    }

    pub fn setXAxisName(self: *Chart, name: []const u8) !void {
        try self.setAxisName(.x_axis, name);
    }

    pub fn setYAxisName(self: *Chart, name: []const u8) !void {
        try self.setAxisName(.y_axis, name);
    }

    pub fn setAxisName(self: *Chart, axis: ChartAxis, name: []const u8) !void {
        const name_str = try self.allocator.dupeZ(u8, name);
        try self.series_strings.append(name_str[0..name_str.len]); // Store without null terminator
        const axis_ptr = switch (axis) {
            .x_axis => self.chart_inner.x_axis,
            .y_axis => self.chart_inner.y_axis,
        };
        _ = xlsxwriter.chart_axis_set_name(axis_ptr, name_str);
    }

    pub fn setStyle(self: *Chart, style_id: u8) void {
        _ = xlsxwriter.chart_set_style(self.chart_inner, style_id);
    }

    pub fn setLegendPosition(self: *Chart, position: ChartLegendPosition) void {
        _ = xlsxwriter.chart_legend_set_position(self.chart_inner, position.toNative());
    }

    pub fn deinit(self: *Chart) void {
        // Free all series
        for (self.series.items) |series| {
            series.deinit();
            self.allocator.destroy(series);
        }
        self.series.deinit();

        // Free all strings
        for (self.series_strings.items) |str| {
            var owned_str = str; // Make a mutable copy
            owned_str.len += 1; // Include null terminator
            self.allocator.free(owned_str);
        }
        self.series_strings.deinit();

        // The chart object itself is cleaned up when the workbook is closed
        self.chart_inner = undefined;
    }
};
