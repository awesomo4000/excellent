const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const c = xlsxwriter.c;
const Colors = @import("colors.zig").Colors;

pub const ChartType = enum {
    column,
    bar,
    bar_stacked,
    bar_stacked_percent,
    line,
    line_stacked,
    line_stacked_percent,
    pie,
    scatter,
    scatter_straight,
    scatter_straight_with_markers,
    scatter_smooth,
    scatter_smooth_with_markers,
    area,
    area_stacked,
    area_stacked_percent,
    radar,
    radar_with_markers,
    radar_filled,
    doughnut,
    column_stacked,
    column_stacked_percent,

    fn toNative(self: ChartType) u8 {
        return switch (self) {
            .column => @intCast(c.LXW_CHART_COLUMN),
            .bar => @intCast(c.LXW_CHART_BAR),
            .bar_stacked => @intCast(c.LXW_CHART_BAR_STACKED),
            .bar_stacked_percent => @intCast(
                c.LXW_CHART_BAR_STACKED_PERCENT,
            ),
            .line => @intCast(c.LXW_CHART_LINE),
            .line_stacked => @intCast(c.LXW_CHART_LINE_STACKED),
            .line_stacked_percent => @intCast(
                c.LXW_CHART_LINE_STACKED_PERCENT,
            ),
            .pie => @intCast(c.LXW_CHART_PIE),
            .area => @intCast(c.LXW_CHART_AREA),
            .area_stacked => @intCast(c.LXW_CHART_AREA_STACKED),
            .area_stacked_percent => @intCast(
                c.LXW_CHART_AREA_STACKED_PERCENT,
            ),
            .radar => @intCast(c.LXW_CHART_RADAR),
            .radar_with_markers => @intCast(
                c.LXW_CHART_RADAR_WITH_MARKERS,
            ),
            .radar_filled => @intCast(c.LXW_CHART_RADAR_FILLED),
            .doughnut => @intCast(c.LXW_CHART_DOUGHNUT),
            .column_stacked => @intCast(c.LXW_CHART_COLUMN_STACKED),
            .column_stacked_percent => @intCast(
                c.LXW_CHART_COLUMN_STACKED_PERCENT,
            ),
            .scatter => @intCast(c.LXW_CHART_SCATTER),
            .scatter_straight => @intCast(
                c.LXW_CHART_SCATTER_STRAIGHT,
            ),
            .scatter_straight_with_markers => @intCast(
                c.LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,
            ),
            .scatter_smooth => @intCast(c.LXW_CHART_SCATTER_SMOOTH),
            .scatter_smooth_with_markers => @intCast(
                c.LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,
            ),
        };
    }
};

pub const ChartFont = struct {
    name: []const u8 = "Calibri",
    size: f64 = 10.0,
    bold: bool = false,
    italic: bool = false,
    color: u32 = Colors.black,
    rotation: i16 = 0,
    underline: bool = false,
    charset: u8 = 0,
    pitchFamily: u8 = 0,
    baseline: i8 = 0,

    fn toNative(self: ChartFont) c.lxw_chart_font {
        return .{
            .name = @ptrCast(self.name),
            .size = self.size,
            .bold = if (self.bold) c.LXW_TRUE else c.LXW_FALSE,
            .italic = if (self.italic) c.LXW_TRUE else c.LXW_FALSE,
            .color = self.color,
            .rotation = self.rotation,
            .underline = if (self.underline) c.LXW_TRUE else c.LXW_FALSE,
            .charset = self.charset,
            .pitch_family = self.pitchFamily,
            .baseline = self.baseline,
        };
    }
};

pub const ChartLine = struct {
    color: u32 = Colors.black,
    width: f32 = 0.75,
    dash_type: DashType = .solid,

    pub const DashType = enum(u8) {
        solid = c.LXW_CHART_LINE_DASH_SOLID,
        dot = c.LXW_CHART_LINE_DASH_DOT,
        dash = c.LXW_CHART_LINE_DASH_DASH,
        long_dash = c.LXW_CHART_LINE_DASH_LONG_DASH,
    };

    fn toNative(self: ChartLine) c.lxw_chart_line {
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

    fn toNative(self: ChartFill) c.lxw_chart_fill {
        return .{
            .color = self.color,
            .transparency = self.transparency,
            .none = 0,
        };
    }
};

pub const ChartPattern = struct {
    pattern_type: PatternType,
    fg_color: u32 = Colors.black,
    bg_color: u32 = Colors.white,

    pub const PatternType = enum(u8) {
        none = c.LXW_CHART_PATTERN_NONE,
        percent_5 = c.LXW_CHART_PATTERN_PERCENT_5,
        percent_10 = c.LXW_CHART_PATTERN_PERCENT_10,
        percent_20 = c.LXW_CHART_PATTERN_PERCENT_20,
        percent_25 = c.LXW_CHART_PATTERN_PERCENT_25,
        percent_30 = c.LXW_CHART_PATTERN_PERCENT_30,
        percent_40 = c.LXW_CHART_PATTERN_PERCENT_40,
        percent_50 = c.LXW_CHART_PATTERN_PERCENT_50,
        percent_60 = c.LXW_CHART_PATTERN_PERCENT_60,
        percent_70 = c.LXW_CHART_PATTERN_PERCENT_70,
        percent_75 = c.LXW_CHART_PATTERN_PERCENT_75,
        percent_80 = c.LXW_CHART_PATTERN_PERCENT_80,
        percent_90 = c.LXW_CHART_PATTERN_PERCENT_90,
        light_downward_diagonal = c.LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL,
        light_upward_diagonal = c.LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL,
        dark_downward_diagonal = c.LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL,
        dark_upward_diagonal = c.LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL,
        wide_downward_diagonal = c.LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL,
        wide_upward_diagonal = c.LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL,
        light_vertical = c.LXW_CHART_PATTERN_LIGHT_VERTICAL,
        light_horizontal = c.LXW_CHART_PATTERN_LIGHT_HORIZONTAL,
        narrow_vertical = c.LXW_CHART_PATTERN_NARROW_VERTICAL,
        narrow_horizontal = c.LXW_CHART_PATTERN_NARROW_HORIZONTAL,
        dark_vertical = c.LXW_CHART_PATTERN_DARK_VERTICAL,
        dark_horizontal = c.LXW_CHART_PATTERN_DARK_HORIZONTAL,
        dashed_downward_diagonal = c.LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL,
        dashed_upward_diagonal = c.LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL,
        dashed_horizontal = c.LXW_CHART_PATTERN_DASHED_HORIZONTAL,
        dashed_vertical = c.LXW_CHART_PATTERN_DASHED_VERTICAL,
        small_confetti = c.LXW_CHART_PATTERN_SMALL_CONFETTI,
        large_confetti = c.LXW_CHART_PATTERN_LARGE_CONFETTI,
        zigzag = c.LXW_CHART_PATTERN_ZIGZAG,
        wave = c.LXW_CHART_PATTERN_WAVE,
        diagonal_brick = c.LXW_CHART_PATTERN_DIAGONAL_BRICK,
        horizontal_brick = c.LXW_CHART_PATTERN_HORIZONTAL_BRICK,
        weave = c.LXW_CHART_PATTERN_WEAVE,
        plaid = c.LXW_CHART_PATTERN_PLAID,
        divot = c.LXW_CHART_PATTERN_DIVOT,
        dotted_grid = c.LXW_CHART_PATTERN_DOTTED_GRID,
        dotted_diamond = c.LXW_CHART_PATTERN_DOTTED_DIAMOND,
        shingle = c.LXW_CHART_PATTERN_SHINGLE,
        trellis = c.LXW_CHART_PATTERN_TRELLIS,
        sphere = c.LXW_CHART_PATTERN_SPHERE,
        small_grid = c.LXW_CHART_PATTERN_SMALL_GRID,
        large_grid = c.LXW_CHART_PATTERN_LARGE_GRID,
        small_check = c.LXW_CHART_PATTERN_SMALL_CHECK,
        large_check = c.LXW_CHART_PATTERN_LARGE_CHECK,
        outlined_diamond = c.LXW_CHART_PATTERN_OUTLINED_DIAMOND,
        solid_diamond = c.LXW_CHART_PATTERN_SOLID_DIAMOND,
    };

    fn toNative(self: ChartPattern) c.lxw_chart_pattern {
        return .{
            .type = @intFromEnum(self.pattern_type),
            .fg_color = self.fg_color,
            .bg_color = self.bg_color,
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
            .none => c.LXW_CHART_LEGEND_NONE,
            .right => c.LXW_CHART_LEGEND_RIGHT,
            .left => c.LXW_CHART_LEGEND_LEFT,
            .top => c.LXW_CHART_LEGEND_TOP,
            .bottom => c.LXW_CHART_LEGEND_BOTTOM,
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
    native_fill: ?c.lxw_chart_fill = null,
    native_line: ?c.lxw_chart_line = null,

    fn toNative(self: *ChartPoint) c.lxw_chart_point {
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
    inner: *c.lxw_chart_series,
    allocator: std.mem.Allocator,
    strings: std.ArrayList([]const u8),
    data_labels_enabled: bool = false,

    pub fn setName(self: *ChartSeries, name: []const u8) !void {
        const name_str = try self.allocator.dupeZ(u8, name);
        try self.strings.append(self.allocator, name_str[0..name_str.len]); // Store without null terminator
        _ = c.chart_series_set_name(self.inner, name_str);
    }

    pub fn setCategories(self: *ChartSeries, sheet: []const u8, first_row: u32, first_col: u16, last_row: u32, last_col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.strings.append(self.allocator, sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = c.chart_series_set_categories(self.inner, sheet_str, first_row, first_col, last_row, last_col);
    }

    pub fn setValues(self: *ChartSeries, sheet: []const u8, first_row: u32, first_col: u16, last_row: u32, last_col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.strings.append(self.allocator, sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = c.chart_series_set_values(self.inner, sheet_str, first_row, first_col, last_row, last_col);
    }

    pub fn setNameRange(self: *ChartSeries, sheet: []const u8, row: u32, col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.strings.append(self.allocator, sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = c.chart_series_set_name_range(self.inner, sheet_str, row, col);
    }

    /// Enable data labels for this series
    pub fn enableDataLabels(self: *ChartSeries) !void {
        _ = c.chart_series_set_labels(self.inner);
        self.data_labels_enabled = true;
    }

    /// Set data label options (which elements to show)
    pub fn setDataLabelOptions(self: *ChartSeries, options: DataLabelOptions) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        _ = c.chart_series_set_labels_options(
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
        _ = c.chart_series_set_labels_font(self.inner, &native_font);
    }

    /// Set line/border properties for data labels
    pub fn setDataLabelLine(self: *ChartSeries, line: ChartLine) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        var native_line = line.toNative();
        _ = c.chart_series_set_labels_line(self.inner, &native_line);
    }

    /// Set fill/background properties for data labels
    pub fn setDataLabelFill(self: *ChartSeries, fill: ChartFill) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        var native_fill = fill.toNative();
        _ = c.chart_series_set_labels_fill(self.inner, &native_fill);
    }

    /// Set custom data labels for this series
    pub fn setCustomDataLabels(self: *ChartSeries, labels: []const ?ChartDataLabel) !void {
        if (!self.data_labels_enabled) {
            return error.DataLabelsNotEnabled;
        }

        var native_labels = try self.allocator.alloc(?*c.lxw_chart_data_label, labels.len + 1);
        defer self.allocator.free(native_labels);

        var native_data_labels = try self.allocator.alloc(c.lxw_chart_data_label, labels.len);
        defer self.allocator.free(native_data_labels);

        // Keep track of allocated strings, fonts, lines, fills to free later
        var value_strings: std.ArrayList([]const u8) = .{};
        defer value_strings.deinit(self.allocator);

        var native_fonts: std.ArrayList(c.lxw_chart_font) = .{};
        defer native_fonts.deinit(self.allocator);

        var native_lines: std.ArrayList(c.lxw_chart_line) = .{};
        defer native_lines.deinit(self.allocator);

        var native_fills: std.ArrayList(c.lxw_chart_fill) = .{};
        defer native_fills.deinit(self.allocator);

        // Set the null terminator at the end of the array
        native_labels[labels.len] = null;

        // Convert each label
        for (labels, 0..) |maybe_label, i| {
            if (maybe_label) |label| {
                native_data_labels[i] = .{};

                // Set value if provided
                if (label.value) |value| {
                    const value_str = try self.allocator.dupeZ(u8, value);
                    try value_strings.append(self.allocator, value_str[0..value_str.len]); // Store for cleanup
                    native_data_labels[i].value = value_str;
                    try self.strings.append(self.allocator, value_str[0..value_str.len]); // Store for deinit
                }

                // Set font if provided
                if (label.font) |font| {
                    const native_font = font.toNative();
                    try native_fonts.append(self.allocator, native_font);
                    native_data_labels[i].font = &native_fonts.items[native_fonts.items.len - 1];
                }

                // Set line if provided
                if (label.line) |line| {
                    const native_line = line.toNative();
                    try native_lines.append(self.allocator, native_line);
                    native_data_labels[i].line = &native_lines.items[native_lines.items.len - 1];
                }

                // Set fill if provided
                if (label.fill) |fill| {
                    const native_fill = fill.toNative();
                    try native_fills.append(self.allocator, native_fill);
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

        _ = c.chart_series_set_labels_custom(self.inner, &native_labels[0]);
    }

    pub const TrendlineType = enum(u8) {
        linear = c.LXW_CHART_TRENDLINE_TYPE_LINEAR,
        poly = c.LXW_CHART_TRENDLINE_TYPE_POLY,
        log = c.LXW_CHART_TRENDLINE_TYPE_LOG,
        exp = c.LXW_CHART_TRENDLINE_TYPE_EXP,
        power = c.LXW_CHART_TRENDLINE_TYPE_POWER,
        moving_average = c.LXW_CHART_TRENDLINE_TYPE_AVERAGE,
    };

    pub fn setTrendline(self: *ChartSeries, trendline_type: TrendlineType, order: u8) !void {
        _ = c.chart_series_set_trendline(self.inner, @intFromEnum(trendline_type), order);
    }

    pub fn setTrendlineLine(self: *ChartSeries, line: *ChartLine) !void {
        var native_line = line.toNative();
        _ = c.chart_series_set_trendline_line(self.inner, &native_line);
    }

    pub const MarkerType = enum(u8) {
        automatic = c.LXW_CHART_MARKER_AUTOMATIC,
        none = c.LXW_CHART_MARKER_NONE,
        square = c.LXW_CHART_MARKER_SQUARE,
        diamond = c.LXW_CHART_MARKER_DIAMOND,
        triangle = c.LXW_CHART_MARKER_TRIANGLE,
        x = c.LXW_CHART_MARKER_X,
        star = c.LXW_CHART_MARKER_STAR,
        short_dash = c.LXW_CHART_MARKER_SHORT_DASH,
        long_dash = c.LXW_CHART_MARKER_LONG_DASH,
        circle = c.LXW_CHART_MARKER_CIRCLE,
        plus = c.LXW_CHART_MARKER_PLUS,
    };

    pub const ErrorBarType = enum(u8) {
        std_error = c.LXW_CHART_ERROR_BAR_TYPE_STD_ERROR,
        fixed = c.LXW_CHART_ERROR_BAR_TYPE_FIXED,
        percentage = c.LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE,
        std_dev = c.LXW_CHART_ERROR_BAR_TYPE_STD_DEV,
    };

    pub fn setMarkerType(self: *ChartSeries, marker_type: MarkerType) !void {
        _ = c.chart_series_set_marker_type(self.inner, @intFromEnum(marker_type));
    }

    pub fn setLabels(self: *ChartSeries) !void {
        _ = c.chart_series_set_labels(self.inner);
    }

    pub fn setErrorBars(self: *ChartSeries, error_bar_type: ErrorBarType, value: f64) !void {
        _ = c.chart_series_set_error_bars(
            self.inner.y_error_bars,
            @intFromEnum(error_bar_type),
            value,
        );
    }

    pub fn setPoints(self: *ChartSeries, points: []const *ChartPoint) !void {
        var native_points = try self.allocator.alloc(?*c.lxw_chart_point, points.len + 1);
        defer self.allocator.free(native_points);

        var native_point_structs = try self.allocator.alloc(c.lxw_chart_point, points.len);
        defer self.allocator.free(native_point_structs);

        // Convert each point to native format
        for (points, 0..) |point, i| {
            native_point_structs[i] = point.toNative();
            native_points[i] = &native_point_structs[i];
        }
        native_points[points.len] = null;

        _ = c.chart_series_set_points(self.inner, @ptrCast(native_points.ptr));
    }

    pub fn setPattern(self: *ChartSeries, pattern: ChartPattern) !void {
        var native_pattern = pattern.toNative();
        _ = c.chart_series_set_pattern(self.inner, &native_pattern);
    }

    pub fn setLine(self: *ChartSeries, line: ChartLine) !void {
        var native_line = line.toNative();
        _ = c.chart_series_set_line(self.inner, &native_line);
    }

    fn deinit(self: *ChartSeries) void {
        for (self.strings.items) |str| {
            var owned_str = str; // Make a mutable copy
            owned_str.len += 1; // Include null terminator
            self.allocator.free(owned_str);
        }
        self.strings.deinit(self.allocator);
    }
};

pub const Chart = struct {
    inner: *c.lxw_chart,
    allocator: std.mem.Allocator,
    series: std.ArrayList(*ChartSeries),
    series_strings: std.ArrayList([]const u8),

    pub fn init(allocator: std.mem.Allocator, workbook: *c.lxw_workbook, chart_type: ChartType) !Chart {
        const inner = c.workbook_add_chart(workbook, chart_type.toNative()) orelse {
            return error.ChartCreationFailed;
        };
        return Chart{
            .inner = inner,
            .allocator = allocator,
            .series = .{},
            .series_strings = .{},
        };
    }

    pub fn addSeries(self: *Chart, categories: ?[]const u8, values: ?[]const u8) !*ChartSeries {
        const cat_ptr: ?[*:0]const u8 = if (categories) |cat| blk: {
            const str = try self.allocator.dupeZ(u8, cat);
            errdefer self.allocator.free(str);
            try self.series_strings.append(self.allocator, str[0..str.len]); // Store without null terminator
            break :blk str;
        } else null;

        const val_ptr: ?[*:0]const u8 = if (values) |v| blk: {
            const str = try self.allocator.dupeZ(u8, v);
            errdefer self.allocator.free(str);
            try self.series_strings.append(self.allocator, str[0..str.len]); // Store without null terminator
            break :blk str;
        } else null;

        const series_inner = c.chart_add_series(self.inner, if (cat_ptr) |cp| @ptrCast(cp) else null, if (val_ptr) |vp| @ptrCast(vp) else null);

        const series = try self.allocator.create(ChartSeries);
        series.* = .{
            .inner = series_inner,
            .allocator = self.allocator,
            .strings = .{},
        };

        try self.series.append(self.allocator, series);
        return series;
    }

    pub fn setSeriesGap(self: *Chart, gap: u16) void {
        _ = c.chart_set_series_gap(self.inner, gap);
    }

    pub fn setTitle(self: *Chart, title: []const u8) !void {
        const title_str = try self.allocator.dupeZ(u8, title);
        errdefer self.allocator.free(title_str);
        try self.series_strings.append(self.allocator, title_str[0..title_str.len]); // Store without null terminator
        _ = c.chart_title_set_name(self.inner, title_str);
    }

    pub fn setTitleFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = c.chart_title_set_name_font(self.inner, &native_font);
    }

    pub fn setXAxisName(self: *Chart, name: []const u8) !void {
        try self.setAxisName(.x_axis, name);
    }

    pub fn setYAxisName(self: *Chart, name: []const u8) !void {
        try self.setAxisName(.y_axis, name);
    }

    pub fn setAxisName(self: *Chart, axis: ChartAxis, name: []const u8) !void {
        const name_str = try self.allocator.dupeZ(u8, name);
        try self.series_strings.append(self.allocator, name_str[0..name_str.len]); // Store without null terminator
        const axis_ptr = switch (axis) {
            .x_axis => self.inner.x_axis,
            .y_axis => self.inner.y_axis,
        };
        _ = c.chart_axis_set_name(axis_ptr, name_str);
    }

    pub fn setAxisNameFont(self: *Chart, axis: ChartAxis, font: ChartFont) void {
        const axis_ptr = switch (axis) {
            .x_axis => self.inner.x_axis,
            .y_axis => self.inner.y_axis,
        };
        var native_font = font.toNative();
        _ = c.chart_axis_set_name_font(axis_ptr, &native_font);
    }

    pub fn setXAxisNameFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = c.chart_axis_set_name_font(self.inner.x_axis, &native_font);
    }

    pub fn setYAxisNameFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = c.chart_axis_set_name_font(self.inner.y_axis, &native_font);
    }

    pub fn setAxisNumFont(self: *Chart, axis: ChartAxis, font: ChartFont) void {
        const axis_ptr = switch (axis) {
            .x_axis => self.inner.x_axis,
            .y_axis => self.inner.y_axis,
        };
        var native_font = font.toNative();
        _ = c.chart_axis_set_num_font(axis_ptr, &native_font);
    }

    pub fn setXAxisNumFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = c.chart_axis_set_num_font(self.inner.x_axis, &native_font);
    }

    pub fn setYAxisNumFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = c.chart_axis_set_num_font(self.inner.y_axis, &native_font);
    }

    pub fn setStyle(self: *Chart, style_id: u8) void {
        _ = c.chart_set_style(self.inner, style_id);
    }

    pub fn setLegendPosition(self: *Chart, position: ChartLegendPosition) void {
        _ = c.chart_legend_set_position(self.inner, position.toNative());
    }

    pub fn setLegendFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = c.chart_legend_set_font(self.inner, &native_font);
    }

    pub fn setTable(self: *Chart) void {
        _ = c.chart_set_table(self.inner);
    }

    pub fn setTableGrid(self: *Chart, horizontal: bool, vertical: bool, outline: bool, legend_keys: bool) void {
        _ = c.chart_set_table_grid(
            self.inner,
            if (horizontal) c.LXW_TRUE else c.LXW_FALSE,
            if (vertical) c.LXW_TRUE else c.LXW_FALSE,
            if (outline) c.LXW_TRUE else c.LXW_FALSE,
            if (legend_keys) c.LXW_TRUE else c.LXW_FALSE,
        );
    }

    pub fn setHighLowLines(self: *Chart, line: ?*ChartLine) !void {
        if (line) |l| {
            var native_line = l.toNative();
            _ = c.chart_set_high_low_lines(self.inner, &native_line);
        } else {
            _ = c.chart_set_high_low_lines(self.inner, null);
        }
    }

    pub fn setDropLines(self: *Chart, line: ?*ChartLine) !void {
        if (line) |l| {
            var native_line = l.toNative();
            _ = c.chart_set_drop_lines(self.inner, &native_line);
        } else {
            _ = c.chart_set_drop_lines(self.inner, null);
        }
    }

    pub fn setUpDownBars(self: *Chart) !void {
        _ = c.chart_set_up_down_bars(self.inner);
    }

    pub fn setUpDownBarsFormat(self: *Chart, up_line: *ChartLine, up_fill: *ChartFill, down_line: *ChartLine, down_fill: *ChartFill) !void {
        var native_up_line = up_line.toNative();
        var native_up_fill = up_fill.toNative();
        var native_down_line = down_line.toNative();
        var native_down_fill = down_fill.toNative();
        _ = c.chart_set_up_down_bars_format(self.inner, &native_up_line, &native_up_fill, &native_down_line, &native_down_fill);
    }

    pub fn setRotation(self: *Chart, rotation: u16) void {
        _ = c.chart_set_rotation(self.inner, rotation);
    }

    pub fn setHoleSize(self: *Chart, size: u8) void {
        _ = c.chart_set_hole_size(self.inner, size);
    }

    pub fn deinit(self: *Chart) void {
        // Free all series
        for (self.series.items) |series| {
            series.deinit();
            self.allocator.destroy(series);
        }
        self.series.deinit(self.allocator);

        // Free all strings
        for (self.series_strings.items) |str| {
            var owned_str = str; // Make a mutable copy
            owned_str.len += 1; // Include null terminator
            self.allocator.free(owned_str);
        }
        self.series_strings.deinit(self.allocator);

        // The chart object itself is cleaned up when the workbook is closed
        self.inner = undefined;
    }
};
