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
    none = 0,
    thin = 1,
    medium = 2,
    dashed = 3,
    dotted = 4,
    thick = 5,
    double = 6,
    hair = 7,
    medium_dashed = 8,
    dash_dot = 9,
    medium_dash_dot = 10,
    dash_dot_dot = 11,
    medium_dash_dot_dot = 12,
    slant_dash_dot = 13,
};

/// Diagonal border types
pub const DiagonalType = enum(u8) {
    none = 0,
    up = 1,
    down = 2,
    up_down = 3,
};

/// Represents a cell format with a fluent builder API.
/// All methods return the Format to allow for method chaining.
pub const Format = struct {
    format: *c.lxw_format,
    allocator: std.mem.Allocator,

    pub fn deinit(self: *Format) void {
        // The format is owned by the workbook, so we don't need to free it
        // But we should clear any resources we've allocated
        self.format = undefined;
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

    /// Set cell pattern
    pub fn setPattern(self: *Format, pattern: u8) *Format {
        _ = c.format_set_pattern(self.format, pattern);
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

    /// Set cell border
    pub fn setBorder(self: *Format, border_style: BorderStyle) *Format {
        _ = c.format_set_border(self.format, @intFromEnum(border_style));
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

    /// Set diagonal border type
    pub fn setDiagonalType(self: *Format, diag_type: DiagonalType) *Format {
        _ = c.format_set_diag_type(self.format, @intFromEnum(diag_type));
        return self;
    }

    /// Set diagonal border style
    pub fn setDiagonalBorder(self: *Format, border_style: BorderStyle) *Format {
        _ = c.format_set_diag_border(self.format, @intFromEnum(border_style));
        return self;
    }

    /// Set diagonal border color
    pub fn setDiagonalColor(self: *Format, color: u32) *Format {
        _ = c.format_set_diag_color(self.format, color);
        return self;
    }

    /// Set text wrapping for a cell
    pub fn setTextWrap(self: *Format) *Format {
        _ = c.format_set_text_wrap(self.format);
        return self;
    }
};
