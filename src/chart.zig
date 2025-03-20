const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const Colors = @import("colors.zig").Colors;

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
    color: u32 = Colors.black,
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
    color: u32 = Colors.black,
    width: f32 = 0.75,
    dash_type: DashType = .solid,

    pub const DashType = enum(u8) {
        solid = xlsxwriter.LXW_CHART_LINE_DASH_SOLID,
        dot = xlsxwriter.LXW_CHART_LINE_DASH_DOT,
        dash = xlsxwriter.LXW_CHART_LINE_DASH_DASH,
        long_dash = xlsxwriter.LXW_CHART_LINE_DASH_LONG_DASH,
    };

    fn toNative(self: ChartLine) xlsxwriter.lxw_chart_line {
        return .{
            .color = self.color,
            .width = self.width,
            .dash_type = @intFromEnum(self.dash_type),
            .transparency = 0,
            .none = 0,
        };
    }
};

pub const ChartFill = struct {
    color: u32 = Colors.black,
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

pub const ChartPoint = struct {
    fill: ?*ChartFill = null,
    line: ?*ChartLine = null,
    native_fill: ?xlsxwriter.lxw_chart_fill = null,
    native_line: ?xlsxwriter.lxw_chart_line = null,

    fn toNative(self: *ChartPoint) xlsxwriter.lxw_chart_point {
        if (self.fill) |f| {
            self.native_fill = f.toNative();
        }
        if (self.line) |l| {
            self.native_line = l.toNative();
        }
        return .{
            .fill = if (self.native_fill) |*f| f else null,
            .line = if (self.native_line) |*l| l else null,
        };
    }
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

    pub const TrendlineType = enum(u8) {
        linear = xlsxwriter.LXW_CHART_TRENDLINE_TYPE_LINEAR,
        poly = xlsxwriter.LXW_CHART_TRENDLINE_TYPE_POLY,
        log = xlsxwriter.LXW_CHART_TRENDLINE_TYPE_LOG,
        exp = xlsxwriter.LXW_CHART_TRENDLINE_TYPE_EXP,
        power = xlsxwriter.LXW_CHART_TRENDLINE_TYPE_POWER,
        moving_average = xlsxwriter.LXW_CHART_TRENDLINE_TYPE_AVERAGE,
    };

    pub fn setTrendline(self: *ChartSeries, trendline_type: TrendlineType, order: u8) !void {
        _ = xlsxwriter.chart_series_set_trendline(self.inner, @intFromEnum(trendline_type), order);
    }

    pub fn setTrendlineLine(self: *ChartSeries, line: *ChartLine) !void {
        var native_line = line.toNative();
        _ = xlsxwriter.chart_series_set_trendline_line(self.inner, &native_line);
    }

    pub const MarkerType = enum(u8) {
        automatic = xlsxwriter.LXW_CHART_MARKER_AUTOMATIC,
        none = xlsxwriter.LXW_CHART_MARKER_NONE,
        square = xlsxwriter.LXW_CHART_MARKER_SQUARE,
        diamond = xlsxwriter.LXW_CHART_MARKER_DIAMOND,
        triangle = xlsxwriter.LXW_CHART_MARKER_TRIANGLE,
        x = xlsxwriter.LXW_CHART_MARKER_X,
        star = xlsxwriter.LXW_CHART_MARKER_STAR,
        short_dash = xlsxwriter.LXW_CHART_MARKER_SHORT_DASH,
        long_dash = xlsxwriter.LXW_CHART_MARKER_LONG_DASH,
        circle = xlsxwriter.LXW_CHART_MARKER_CIRCLE,
        plus = xlsxwriter.LXW_CHART_MARKER_PLUS,
    };

    pub const ErrorBarType = enum(u8) {
        std_error = xlsxwriter.LXW_CHART_ERROR_BAR_TYPE_STD_ERROR,
        fixed = xlsxwriter.LXW_CHART_ERROR_BAR_TYPE_FIXED,
        percentage = xlsxwriter.LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE,
        std_dev = xlsxwriter.LXW_CHART_ERROR_BAR_TYPE_STD_DEV,
    };

    pub fn setMarkerType(self: *ChartSeries, marker_type: MarkerType) !void {
        _ = xlsxwriter.chart_series_set_marker_type(self.inner, @intFromEnum(marker_type));
    }

    pub fn setLabels(self: *ChartSeries) !void {
        _ = xlsxwriter.chart_series_set_labels(self.inner);
    }

    pub fn setErrorBars(self: *ChartSeries, error_bar_type: ErrorBarType, value: f64) !void {
        _ = xlsxwriter.chart_series_set_error_bars(
            self.inner.y_error_bars,
            @intFromEnum(error_bar_type),
            value,
        );
    }

    pub fn setPoints(self: *ChartSeries, points: []const *ChartPoint) !void {
        var native_points = try self.allocator.alloc(?*xlsxwriter.lxw_chart_point, points.len + 1);
        defer self.allocator.free(native_points);

        var native_point_structs = try self.allocator.alloc(xlsxwriter.lxw_chart_point, points.len);
        defer self.allocator.free(native_point_structs);

        // Convert each point to native format
        for (points, 0..) |point, i| {
            native_point_structs[i] = point.toNative();
            native_points[i] = &native_point_structs[i];
        }
        native_points[points.len] = null;

        _ = xlsxwriter.chart_series_set_points(self.inner, @ptrCast(native_points.ptr));
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
    inner: *xlsxwriter.lxw_chart,
    allocator: std.mem.Allocator,
    series: std.ArrayList(*ChartSeries),
    series_strings: std.ArrayList([]const u8),

    pub fn init(allocator: std.mem.Allocator, workbook: *xlsxwriter.lxw_workbook, chart_type: ChartType) !Chart {
        const inner = xlsxwriter.workbook_add_chart(workbook, chart_type.toNative()) orelse {
            return error.ChartCreationFailed;
        };
        return Chart{
            .inner = inner,
            .allocator = allocator,
            .series = std.ArrayList(*ChartSeries).init(allocator),
            .series_strings = std.ArrayList([]const u8).init(allocator),
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

        const series_inner = xlsxwriter.chart_add_series(self.inner, if (cat_ptr) |c| @ptrCast(c) else null, if (val_ptr) |v| @ptrCast(v) else null);

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
        _ = xlsxwriter.chart_title_set_name(self.inner, title_str);
    }

    pub fn setTitleFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = xlsxwriter.chart_title_set_name_font(self.inner, &native_font);
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
            .x_axis => self.inner.x_axis,
            .y_axis => self.inner.y_axis,
        };
        _ = xlsxwriter.chart_axis_set_name(axis_ptr, name_str);
    }

    pub fn setStyle(self: *Chart, style_id: u8) void {
        _ = xlsxwriter.chart_set_style(self.inner, style_id);
    }

    pub fn setLegendPosition(self: *Chart, position: ChartLegendPosition) void {
        _ = xlsxwriter.chart_legend_set_position(self.inner, position.toNative());
    }

    pub fn setTable(self: *Chart) void {
        _ = xlsxwriter.chart_set_table(self.inner);
    }

    pub fn setTableGrid(self: *Chart, horizontal: bool, vertical: bool, outline: bool, legend_keys: bool) void {
        _ = xlsxwriter.chart_set_table_grid(
            self.inner,
            if (horizontal) xlsxwriter.LXW_TRUE else xlsxwriter.LXW_FALSE,
            if (vertical) xlsxwriter.LXW_TRUE else xlsxwriter.LXW_FALSE,
            if (outline) xlsxwriter.LXW_TRUE else xlsxwriter.LXW_FALSE,
            if (legend_keys) xlsxwriter.LXW_TRUE else xlsxwriter.LXW_FALSE,
        );
    }

    pub fn setHighLowLines(self: *Chart, line: ?*ChartLine) !void {
        if (line) |l| {
            var native_line = l.toNative();
            _ = xlsxwriter.chart_set_high_low_lines(self.inner, &native_line);
        } else {
            _ = xlsxwriter.chart_set_high_low_lines(self.inner, null);
        }
    }

    pub fn setDropLines(self: *Chart, line: ?*ChartLine) !void {
        if (line) |l| {
            var native_line = l.toNative();
            _ = xlsxwriter.chart_set_drop_lines(self.inner, &native_line);
        } else {
            _ = xlsxwriter.chart_set_drop_lines(self.inner, null);
        }
    }

    pub fn setUpDownBars(self: *Chart) !void {
        _ = xlsxwriter.chart_set_up_down_bars(self.inner);
    }

    pub fn setUpDownBarsFormat(self: *Chart, up_line: *ChartLine, up_fill: *ChartFill, down_line: *ChartLine, down_fill: *ChartFill) !void {
        var native_up_line = up_line.toNative();
        var native_up_fill = up_fill.toNative();
        var native_down_line = down_line.toNative();
        var native_down_fill = down_fill.toNative();
        _ = xlsxwriter.chart_set_up_down_bars_format(self.inner, &native_up_line, &native_up_fill, &native_down_line, &native_down_fill);
    }

    pub fn setRotation(self: *Chart, rotation: u16) void {
        _ = xlsxwriter.chart_set_rotation(self.inner, rotation);
    }

    pub fn setHoleSize(self: *Chart, size: u8) void {
        _ = xlsxwriter.chart_set_hole_size(self.inner, size);
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
        self.inner = undefined;
    }
};
