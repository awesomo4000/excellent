const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

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
    name: []const u8,
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

pub const Chart = struct {
    inner: *xlsxwriter.lxw_chart,
    allocator: std.mem.Allocator,
    series: std.ArrayList(*xlsxwriter.lxw_chart_series),
    series_strings: std.ArrayList([]const u8),

    pub fn init(allocator: std.mem.Allocator, workbook: *xlsxwriter.lxw_workbook, chart_type: ChartType) !Chart {
        const inner = xlsxwriter.workbook_add_chart(workbook, chart_type.toNative()) orelse {
            return error.ChartCreationFailed;
        };
        return Chart{
            .inner = inner,
            .allocator = allocator,
            .series = std.ArrayList(*xlsxwriter.lxw_chart_series).init(allocator),
            .series_strings = std.ArrayList([]const u8).init(allocator),
        };
    }

    pub fn addSeries(self: *Chart, categories: ?[]const u8, values: ?[]const u8) !void {
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

        const series = xlsxwriter.chart_add_series(self.inner, if (cat_ptr) |c| @ptrCast(c) else null, if (val_ptr) |v| @ptrCast(v) else null);
        try self.series.append(series);
    }

    pub fn setSeriesName(self: *Chart, index: u8, name: []const u8) !void {
        const name_str = try self.allocator.dupeZ(u8, name);
        try self.series_strings.append(name_str[0..name_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_name(self.series.items[index], name_str);
    }

    pub fn setSeriesCategories(self: *Chart, index: u8, sheet: []const u8, first_row: u32, first_col: u16, last_row: u32, last_col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.series_strings.append(sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_categories(self.series.items[index], sheet_str, first_row, first_col, last_row, last_col);
    }

    pub fn setSeriesValues(self: *Chart, index: u8, sheet: []const u8, first_row: u32, first_col: u16, last_row: u32, last_col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.series_strings.append(sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_values(self.series.items[index], sheet_str, first_row, first_col, last_row, last_col);
    }

    pub fn setSeriesNameRange(self: *Chart, index: u8, sheet: []const u8, row: u32, col: u16) !void {
        const sheet_str = try self.allocator.dupeZ(u8, sheet);
        try self.series_strings.append(sheet_str[0..sheet_str.len]); // Store without null terminator
        _ = xlsxwriter.chart_series_set_name_range(self.series.items[index], sheet_str, row, col);
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

    pub fn setStyle(self: *Chart, style_id: u8) void {
        _ = xlsxwriter.chart_set_style(self.inner, style_id);
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

    pub fn setLegendPosition(self: *Chart, position: ChartLegendPosition) void {
        _ = xlsxwriter.chart_legend_set_position(self.inner, position.toNative());
    }

    pub fn deinit(self: *Chart) void {
        for (self.series_strings.items) |str| {
            var owned_str = str; // Make a mutable copy
            owned_str.len += 1; // Include null terminator
            self.allocator.free(owned_str);
        }
        self.series_strings.deinit();
        self.series.deinit();
    }
};
