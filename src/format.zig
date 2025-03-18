const std = @import("std");
const c = @import("xlsxwriter");
const error_utils = @import("error_utils.zig");

/// Alignment options for cells
pub const Alignment = enum(u8) {
    None = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterAcross = 6,
    Distributed = 7,
    VerticalTop = 8,
    VerticalBottom = 9,
    VerticalCenter = 10,
    VerticalJustify = 11,
    VerticalDistributed = 12,
};

/// Border styles for cells
pub const BorderStyle = enum(u8) {
    None = 0,
    Thin = 1,
    Medium = 2,
    Dashed = 3,
    Dotted = 4,
    Thick = 5,
    Double = 6,
    Hair = 7,
    MediumDashed = 8,
    DashDot = 9,
    MediumDashDot = 10,
    DashDotDot = 11,
    MediumDashDotDot = 12,
    SlantDashDot = 13,
};

/// Represents a cell format with a fluent builder API.
/// All methods return the Format to allow for method chaining.
pub const Format = struct {
    format: *c.lxw_format,
    allocator: std.mem.Allocator,

    pub fn deinit(self: *Format) void {
        self.allocator.destroy(self);
    }

    /// Set font properties
    pub fn setFontName(self: *Format, font_name: []const u8) !*Format {
        const c_font_name = try self.allocator.dupeZ(u8, font_name);
        defer self.allocator.free(c_font_name);

        _ = c.format_set_font_name(self.format, c_font_name.ptr);
        return self;
    }

    pub fn setFontSize(self: *Format, size: f64) *Format {
        _ = c.format_set_font_size(self.format, size);
        return self;
    }

    pub fn setBold(self: *Format) *Format {
        _ = c.format_set_bold(self.format);
        return self;
    }

    pub fn setItalic(self: *Format) *Format {
        _ = c.format_set_italic(self.format);
        return self;
    }

    pub fn setUnderline(self: *Format, underline_type: u8) *Format {
        _ = c.format_set_underline(self.format, underline_type);
        return self;
    }

    pub fn setNumFormat(self: *Format, num_format: []const u8) !*Format {
        const c_num_format = try self.allocator.dupeZ(u8, num_format);
        defer self.allocator.free(c_num_format);

        _ = c.format_set_num_format(self.format, c_num_format.ptr);
        return self;
    }

    /// Set cell alignment
    pub fn setAlign(self: *Format, alignment: Alignment) *Format {
        _ = c.format_set_align(self.format, @intFromEnum(alignment));
        return self;
    }

    /// Set cell border
    pub fn setBorder(self: *Format, border_style: BorderStyle) *Format {
        _ = c.format_set_border(self.format, @intFromEnum(border_style));
        return self;
    }

    /// Set background color
    pub fn setBgColor(self: *Format, color: u32) *Format {
        _ = c.format_set_bg_color(self.format, color);
        return self;
    }

    /// Set font color
    pub fn setFontColor(self: *Format, color: u32) *Format {
        _ = c.format_set_font_color(self.format, color);
        return self;
    }

    /// Helper method to set left, right, top, and bottom borders to the same style
    pub fn setBorders(self: *Format, border_style: BorderStyle) *Format {
        const style = @intFromEnum(border_style);
        _ = c.format_set_bottom(self.format, style);
        _ = c.format_set_top(self.format, style);
        _ = c.format_set_left(self.format, style);
        _ = c.format_set_right(self.format, style);
        return self;
    }

    /// Set individual border styles
    pub fn setBottomBorder(self: *Format, border_style: BorderStyle) *Format {
        _ = c.format_set_bottom(self.format, @intFromEnum(border_style));
        return self;
    }

    pub fn setTopBorder(self: *Format, border_style: BorderStyle) *Format {
        _ = c.format_set_top(self.format, @intFromEnum(border_style));
        return self;
    }

    pub fn setLeftBorder(self: *Format, border_style: BorderStyle) *Format {
        _ = c.format_set_left(self.format, @intFromEnum(border_style));
        return self;
    }

    pub fn setRightBorder(self: *Format, border_style: BorderStyle) *Format {
        _ = c.format_set_right(self.format, @intFromEnum(border_style));
        return self;
    }
};
