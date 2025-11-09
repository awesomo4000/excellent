//
// An example of writing cell comments to a worksheet using libxlsxwriter.
//
// Each of the worksheets demonstrates different features of cell comments.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const excel = @import("excellent");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(allocator, "comments2.xlsx");
    defer {
        _ = workbook.close() catch {};
        workbook.deinit();
    }

    // Create a text wrap format
    var text_wrap = try workbook.addFormat();
    _ = text_wrap.setTextWrap();
    _ = text_wrap.setAlign(.vertical_top);

    // Create worksheets directly rather than storing pointers to them
    var ws1 = try workbook.addWorksheet("Sheet1");
    var ws2 = try workbook.addWorksheet("Sheet2");
    var ws3 = try workbook.addWorksheet("Sheet3");
    var ws4 = try workbook.addWorksheet("Sheet4");
    var ws5 = try workbook.addWorksheet("Sheet5");
    var ws6 = try workbook.addWorksheet("Sheet6");
    var ws7 = try workbook.addWorksheet("Sheet7");
    var ws8 = try workbook.addWorksheet("Sheet8");

    // Example 1: Simple cell comment without formatting
    _ = ws1.setColumnWidth(2, 2, 25);
    _ = ws1.setRowHeight(2, 50);

    try ws1.writeString(2, 2, "Hold the mouse over this cell to see the comment.", text_wrap);
    try ws1.writeComment(2, 2, "This is a comment.");

    // Example 2: Visible and hidden comments
    _ = ws2.setColumnWidth(2, 2, 25);
    _ = ws2.setRowHeight(2, 50);

    try ws2.writeString(2, 2, "This cell comment is visible.", text_wrap);
    try ws2.writeCommentOpt(2, 2, "Hello.", .{
        .visible = true,
        .x_scale = 1.0,
        .y_scale = 1.0,
    });

    try ws2.writeString(5, 2, "This cell comment isn't visible until you pass the mouse over it (the default).", text_wrap);
    try ws2.writeComment(5, 2, "Hello.");

    // Example 3: Worksheet level comment visibility
    _ = ws3.setColumnWidth(2, 2, 25);
    _ = ws3.setRowHeight(2, 50);
    _ = ws3.setRowHeight(5, 50);
    _ = ws3.setRowHeight(8, 50);

    _ = ws3.showComments();

    try ws3.writeString(2, 2, "This cell comment is visible, explicitly.", text_wrap);
    try ws3.writeCommentOpt(2, 2, "Hello", .{
        .visible = true,
        .x_scale = 1.0,
        .y_scale = 1.0,
    });

    try ws3.writeString(5, 2, "This cell comment is also visible because we used worksheet_show_comments().", text_wrap);
    try ws3.writeComment(5, 2, "Hello");

    try ws3.writeString(8, 2, "However, we can still override it locally.", text_wrap);
    try ws3.writeCommentOpt(8, 2, "Hello", .{
        .visible = false,
        .force_hidden = true,
        .color = 0x0,
        .author = null,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .width = 0,
        .height = 0,
    });

    // Example 4: Comment box dimensions
    _ = ws4.setColumnWidth(2, 2, 25);
    _ = ws4.setRowHeight(2, 50);
    _ = ws4.setRowHeight(5, 50);
    _ = ws4.setRowHeight(8, 50);
    _ = ws4.setRowHeight(15, 50);
    _ = ws4.setRowHeight(18, 50);

    _ = ws4.showComments();

    try ws4.writeString(2, 2, "This cell comment is default size.", text_wrap);
    try ws4.writeComment(2, 2, "Hello");

    try ws4.writeString(5, 2, "This cell comment is twice as wide.", text_wrap);
    try ws4.writeCommentOpt(5, 2, "Hello", .{
        .visible = false,
        .x_scale = 2.0,
        .y_scale = 1.0,
    });

    try ws4.writeString(8, 2, "This cell comment is twice as high.", text_wrap);
    try ws4.writeCommentOpt(8, 2, "Hello", .{
        .visible = false,
        .x_scale = 1.0,
        .y_scale = 2.0,
    });

    try ws4.writeString(15, 2, "This cell comment is scaled in both directions.", text_wrap);
    try ws4.writeCommentOpt(15, 2, "Hello", .{
        .visible = false,
        .x_scale = 1.2,
        .y_scale = 0.5,
    });

    try ws4.writeString(18, 2, "This cell comment has width and height specified in pixels.", text_wrap);
    try ws4.writeCommentOpt(18, 2, "Hello", .{
        .visible = false,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .width = 200,
        .height = 50,
    });

    // Example 5: Comment positioning
    _ = ws5.setColumnWidth(2, 2, 25);
    _ = ws5.setRowHeight(2, 50);
    _ = ws5.setRowHeight(5, 50);
    _ = ws5.setRowHeight(8, 50);

    _ = ws5.showComments();

    try ws5.writeString(2, 2, "This cell comment is in the default position.", text_wrap);
    try ws5.writeComment(2, 2, "Hello");

    try ws5.writeString(5, 2, "This cell comment has been moved to another cell.", text_wrap);
    try ws5.writeCommentOpt(5, 2, "Hello", .{
        .visible = false,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .start_row = 3,
        .start_col = 4,
    });

    try ws5.writeString(8, 2, "This cell comment has been shifted within its default cell.", text_wrap);
    try ws5.writeCommentOpt(8, 2, "Hello", .{
        .visible = false,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .x_offset = 30,
        .y_offset = 12,
    });

    // Example 6: Comment colors
    _ = ws6.setColumnWidth(2, 2, 25);
    _ = ws6.setRowHeight(2, 50);
    _ = ws6.setRowHeight(5, 50);
    _ = ws6.setRowHeight(8, 50);

    _ = ws6.showComments();

    try ws6.writeString(2, 2, "This cell comment has a different color.", text_wrap);
    try ws6.writeCommentOpt(2, 2, "Hello", .{
        .color = 0x008000, // Green
    });

    try ws6.writeString(5, 2, "This cell comment has the default color.", text_wrap);
    try ws6.writeComment(5, 2, "Hello");

    try ws6.writeString(8, 2, "This cell comment has a different color.", text_wrap);
    try ws6.writeCommentOpt(8, 2, "Hello", .{
        .color = 0xFF6600,
    });

    // Example 7: Comment authors
    _ = ws7.setColumnWidth(2, 2, 25);
    _ = ws7.setRowHeight(2, 50);
    _ = ws7.setRowHeight(5, 60);

    try ws7.writeString(2, 2, "Move the mouse over this cell and you will see 'Cell C3 commented by' (blank) in the status bar at the bottom.", text_wrap);
    try ws7.writeComment(2, 2, "Hello");

    try ws7.writeString(5, 2, "Move the mouse over this cell and you will see 'Cell C6 commented by libxlsxwriter' in the status bar at the bottom.", text_wrap);
    try ws7.writeCommentOpt(5, 2, "Hello", .{
        .author = "libxlsxwriter",
    });

    // Example 8: Row height and comment box size relationship
    _ = ws8.setColumnWidth(2, 2, 25);

    // Set explicit row height
    _ = ws8.setRowHeight(2, 80);
    _ = ws8.showComments();

    try ws8.writeString(2, 2, "The height of this row has been adjusted explicitly using worksheet_set_row(). The size of the comment box is adjusted accordingly by libxlsxwriter", text_wrap);
    try ws8.writeComment(2, 2, "Hello");

    // Row with text wrap
    try ws8.writeString(5, 2, "The height of this row has been adjusted by Excel when the file is opened due to the text wrap property being set. Unfortunately this means that the height of the row is unknown to libxlsxwriter at run time and thus the comment box is stretched as well.\n\nUse worksheet_set_row() to specify the row height explicitly to avoid this problem.", text_wrap);
    try ws8.writeComment(5, 2, "Hello");
}
