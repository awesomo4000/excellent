const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub const ChartType = enum {
    column,
    bar,
    line,
    pie,
    scatter,
    area,
    radar,
    doughnut,

    fn toNative(self: ChartType) u8 {
        return switch (self) {
            .column => @intCast(xlsxwriter.LXW_CHART_COLUMN),
            .bar => @intCast(xlsxwriter.LXW_CHART_BAR),
            .line => @intCast(xlsxwriter.LXW_CHART_LINE),
            .pie => @intCast(xlsxwriter.LXW_CHART_PIE),
            .scatter => @intCast(xlsxwriter.LXW_CHART_SCATTER),
            .area => @intCast(xlsxwriter.LXW_CHART_AREA),
            .radar => @intCast(xlsxwriter.LXW_CHART_RADAR),
            .doughnut => @intCast(xlsxwriter.LXW_CHART_DOUGHNUT),
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

pub const Chart = struct {
    inner: ?*xlsxwriter.lxw_chart,
    allocator: std.mem.Allocator,

    pub fn init(allocator: std.mem.Allocator, workbook: *xlsxwriter.lxw_workbook, chart_type: ChartType) !Chart {
        const inner = xlsxwriter.workbook_add_chart(workbook, chart_type.toNative()) orelse {
            return error.ChartCreationFailed;
        };
        return Chart{
            .inner = inner,
            .allocator = allocator,
        };
    }

    pub fn addSeries(self: *Chart, categories: ?[]const u8, values: []const u8) !void {
        const cat_ptr: ?[*:0]const u8 = if (categories) |c| (try self.allocator.dupeZ(u8, c)).ptr else null;
        const val_ptr = (try self.allocator.dupeZ(u8, values)).ptr;
        defer if (cat_ptr) |c| self.allocator.free(@as([*:0]u8, @constCast(c))[0..categories.?.len :0]);
        defer self.allocator.free(val_ptr[0..values.len :0]);

        _ = xlsxwriter.chart_add_series(self.inner, if (cat_ptr) |c| @ptrCast(c) else null, @ptrCast(val_ptr));
    }

    pub fn setTitle(self: *Chart, title: []const u8) !void {
        const title_str = try self.allocator.dupeZ(u8, title);
        defer self.allocator.free(title_str);
        _ = xlsxwriter.chart_title_set_name(self.inner, title_str);
    }

    pub fn setTitleFont(self: *Chart, font: ChartFont) void {
        var native_font = font.toNative();
        _ = xlsxwriter.chart_title_set_name_font(self.inner, &native_font);
    }

    pub fn setStyle(self: *Chart, style_id: u8) void {
        _ = xlsxwriter.chart_set_style(self.inner, style_id);
    }

    pub fn setLegendPosition(self: *Chart, position: ChartLegendPosition) void {
        _ = xlsxwriter.chart_legend_set_position(self.inner, position.toNative());
    }
};
