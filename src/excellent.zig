const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

pub const Workbook = struct {
    allocator: std.mem.Allocator,
    filename: []const u8,
    workbook: *xlsxwriter.lxw_workbook,

    pub fn init(filename: []const u8) !Workbook {
        const workbook = xlsxwriter.workbook_new(filename.ptr);
        if (workbook == null) return error.WorkbookCreationFailed;

        return Workbook{
            .allocator = std.heap.page_allocator,
            .filename = filename,
            .workbook = workbook,
        };
    }

    pub fn deinit(self: *Workbook) void {
        _ = xlsxwriter.workbook_close(self.workbook);
    }

    pub fn addWorksheet(self: *Workbook, name: ?[]const u8) !Worksheet {
        const name_ptr = if (name) |n| n.ptr else null;
        const worksheet = xlsxwriter.workbook_add_worksheet(self.workbook, name_ptr);
        if (worksheet == null) return error.WorksheetCreationFailed;

        return Worksheet{
            .workbook = self,
            .worksheet = worksheet,
        };
    }
};

pub const Worksheet = struct {
    workbook: *Workbook,
    worksheet: *xlsxwriter.lxw_worksheet,

    pub fn writeString(self: *Worksheet, row: usize, col: usize, text: []const u8) !void {
        const result = xlsxwriter.worksheet_write_string(self.worksheet, @intCast(row), @intCast(col), text.ptr, null);
        if (result != xlsxwriter.LXW_NO_ERROR) return error.WriteFailed;
    }

    pub fn writeNumber(self: *Worksheet, row: usize, col: usize, number: f64) !void {
        const result = xlsxwriter.worksheet_write_number(self.worksheet, @intCast(row), @intCast(col), number, null);
        if (result != xlsxwriter.LXW_NO_ERROR) return error.WriteFailed;
    }
};
