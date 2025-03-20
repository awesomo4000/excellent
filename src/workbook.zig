const std = @import("std");
const c = @import("xlsxwriter");
const error_utils = @import("error_utils.zig");
const worksheet_mod = @import("worksheet.zig");
const format_mod = @import("format.zig");
const chart_mod = @import("chart.zig");

pub const Workbook = struct {
    allocator: std.mem.Allocator,
    filename: []const u8,
    workbook: *c.lxw_workbook,
    isOpen: bool = false,
    formats: std.ArrayList(*format_mod.Format) = std.ArrayList(*format_mod.Format).init(std.heap.page_allocator),
    charts: std.ArrayList(*chart_mod.Chart) = std.ArrayList(*chart_mod.Chart).init(std.heap.page_allocator),

    pub fn create(
        allocator: std.mem.Allocator,
        filename: []const u8,
    ) !*Workbook {
        const c_workbook = c.workbook_new(filename.ptr);
        if (c_workbook == null) return error.WorkbookCreationFailed;

        const workbook = try allocator.create(Workbook);

        workbook.* = .{
            .allocator = allocator,
            .filename = filename,
            .workbook = c_workbook,
            .isOpen = true,
            .formats = std.ArrayList(*format_mod.Format).init(allocator),
            .charts = std.ArrayList(*chart_mod.Chart).init(allocator),
        };

        return workbook;
    }

    pub fn close(self: *Workbook) !void {
        if (!self.isOpen) return;
        const error_code = c.workbook_close(self.workbook);
        self.isOpen = false;
        if (error_code != c.LXW_NO_ERROR) return error_utils.translateErrorCode(error_code);
        self.isOpen = false;
    }

    pub fn deinit(self: *Workbook) void {
        if (self.isOpen) {
            _ = c.workbook_close(self.workbook);
            self.isOpen = false;
        }
        // deinit the formats
        for (self.formats.items) |format| {
            format.deinit();
            self.allocator.destroy(format);
        }
        self.formats.deinit();
        // deinit the charts
        for (self.charts.items) |chart| {
            chart.deinit();
            self.allocator.destroy(chart);
        }
        self.charts.deinit();
        self.allocator.destroy(self);
    }

    pub fn addFormat(self: *Workbook) !*format_mod.Format {
        if (!self.isOpen) return error_utils.XlsxError.GenericError;

        const c_format = c.workbook_add_format(self.workbook) orelse {
            return error_utils.XlsxError.FormatError;
        };

        const format = try self.allocator.create(format_mod.Format);
        format.* = .{
            .format = c_format,
            .allocator = self.allocator,
        };
        try self.formats.append(format);
        return format;
    }

    pub fn addChart(self: *Workbook, chart_type: chart_mod.ChartType) !*chart_mod.Chart {
        if (!self.isOpen) return error_utils.XlsxError.GenericError;

        const chart = try self.allocator.create(chart_mod.Chart);
        chart.* = try chart_mod.Chart.init(self.allocator, self.workbook, chart_type);
        try self.charts.append(chart);
        return chart;
    }

    pub fn addWorksheet(self: *Workbook, name: ?[]const u8) !worksheet_mod.Worksheet {
        const name_ptr = if (name) |n| n.ptr else null;
        const worksheet = c.workbook_add_worksheet(self.workbook, name_ptr);
        if (worksheet == null) return error.WorksheetCreationFailed;

        return worksheet_mod.Worksheet{
            .workbook = self,
            .worksheet = worksheet,
        };
    }
};
