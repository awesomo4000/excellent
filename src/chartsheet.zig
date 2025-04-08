const std = @import("std");
const c = @import("xlsxwriter");
const chart_mod = @import("chart.zig");

/// Represents a chartsheet within a workbook
pub const Chartsheet = struct {
    workbook: *@import("workbook.zig").Workbook,
    chartsheet: *c.lxw_chartsheet,

    pub fn deinit(self: *Chartsheet) void {
        // The chartsheet is owned by the workbook, so we don't need to free it
        // But we should clear any resources we've allocated
        self.chartsheet = undefined;
        self.workbook = undefined;
    }

    /// Set the chart for this chartsheet
    pub fn setChart(self: *Chartsheet, chart: *chart_mod.Chart) !void {
        const result = c.chartsheet_set_chart(self.chartsheet, chart.inner);
        if (result != c.LXW_NO_ERROR) return error.SetChartFailed;
    }

    /// Make this chartsheet the active sheet
    pub fn activate(self: *Chartsheet) !void {
        c.chartsheet_activate(self.chartsheet);
    }

    /// Set the zoom level for the chartsheet (10-400)
    pub fn setZoom(self: *Chartsheet, zoom: u16) void {
        c.chartsheet_set_zoom(self.chartsheet, zoom);
    }
};
