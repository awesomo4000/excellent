const std = @import("std");
const xlsxwriter = @import("xlsxwriter");

// Re-export all module components
pub const cell = @import("cell_utils.zig").cell;
pub const XlsxError = @import("error_utils.zig").XlsxError;
pub const Format = @import("format.zig").Format;
pub const Alignment = @import("format.zig").Alignment;
pub const BorderStyle = @import("format.zig").BorderStyle;
pub const Workbook = @import("workbook.zig").Workbook;
pub const Worksheet = @import("worksheet.zig").Worksheet;
pub const StyledText = @import("styled.zig").StyledText;
pub const StyledWriter = @import("styled.zig").StyledWriter;
pub const Styled = @import("styled.zig").Styled;
pub const Chart = @import("chart.zig").Chart;
pub const ChartType = @import("chart.zig").ChartType;
pub const ChartFont = @import("chart.zig").ChartFont;
pub const DiagonalType = @import("format.zig").DiagonalType;
pub const CommentOptions = @import("comment.zig").CommentOptions;
pub const TmpFile = @import("mktmp.zig").TmpFile;
/// Excel DateTime representation
pub const DateTime = struct {
    year: u16 = 0,
    month: u8 = 0,
    day: u8 = 0,
    hour: u8 = 0,
    min: u8 = 0,
    sec: f64 = 0,

    /// Convert to lxw_datetime for internal use
    pub fn toLxwDateTime(self: DateTime) @import("xlsxwriter").lxw_datetime {
        return @import("xlsxwriter").lxw_datetime{
            .year = self.year,
            .month = self.month,
            .day = self.day,
            .hour = self.hour,
            .min = self.min,
            .sec = self.sec,
        };
    }
};

// Helper function to convert a double timestamp to lxw_datetime
// Place in appropriate module later if functionality expands
pub fn dateTimeToLxwDateTime(_: f64) @import("xlsxwriter").lxw_datetime {
    // This is a simplified implementation
    // In a real application, you would convert from a standard datetime format
    return @import("xlsxwriter").lxw_datetime{
        .year = 2023,
        .month = 1,
        .day = 1,
        .hour = 0,
        .min = 0,
        .sec = 0,
    };
}

// Include all test files in the test build
comptime {
    _ = @import("test_worksheet.zig");
    _ = @import("test_excellent.zig");
    _ = @import("test_format.zig");
    _ = @import("test_workbook.zig");
    _ = @import("test_styled.zig");
    _ = @import("test_cell_utils.zig");
    _ = @import("test_error_utils.zig");
    _ = @import("test_fail.zig");
    _ = @import("test_chart.zig");
}

test {
    std.testing.refAllDecls(@This());
}
