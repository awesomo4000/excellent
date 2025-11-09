const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const c = xlsxwriter.c;

/// Options for cell comments
pub const CommentOptions = struct {
    /// Whether the comment is visible by default
    visible: bool = false,
    /// Whether to explicitly force hiding the comment (overrides worksheet_show_comments)
    force_hidden: bool = false,
    /// Scale factor for comment box width
    x_scale: f64 = 1.0,
    /// Scale factor for comment box height
    y_scale: f64 = 1.0,
    /// Color of the comment box (RGB format)
    color: u32 = 0x0,
    /// Starting row for the comment box
    start_row: u32 = 0,
    /// Starting column for the comment box
    start_col: u16 = 0,
    /// X offset for the comment box
    x_offset: i32 = 0,
    /// Y offset for the comment box
    y_offset: i32 = 0,
    /// Author of the comment
    author: ?[]const u8 = null,
    /// Width of the comment box in pixels
    width: u16 = 0,
    /// Height of the comment box in pixels
    height: u16 = 0,

    /// Convert CommentOptions to C library's lxw_comment_options
    pub fn toCOptions(self: CommentOptions, allocator: std.mem.Allocator) !c.lxw_comment_options {
        var options = c.lxw_comment_options{
            .visible = if (self.visible)
                c.LXW_COMMENT_DISPLAY_VISIBLE
            else if (self.force_hidden)
                c.LXW_COMMENT_DISPLAY_HIDDEN
            else
                c.LXW_COMMENT_DISPLAY_DEFAULT,
            .x_scale = self.x_scale,
            .y_scale = self.y_scale,
            .color = self.color,
            .start_row = self.start_row,
            .start_col = self.start_col,
            .x_offset = self.x_offset,
            .y_offset = self.y_offset,
            .width = self.width,
            .height = self.height,
            .author = null,
        };

        if (self.author) |author| {
            const author_z = try allocator.dupeZ(u8, author);
            options.author = author_z.ptr;
            return options;
        }

        return options;
    }

    /// Free resources associated with the comment options
    pub fn deinit(_: *CommentOptions, _: std.mem.Allocator) void {
        // Nothing to free since we don't own any allocated memory in this struct
    }
};
