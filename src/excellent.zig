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
// pub const ValidationType = @import("data_validation.zig").ValidationType;
// pub const ValidationCriteria = @import("data_validation.zig").ValidationCriteria;
// pub const ValidationErrorType = @import("data_validation.zig").ValidationErrorType;
pub const DataValidation = @import("data_validation.zig");
pub const StyledText = @import("styled.zig").StyledText;
pub const StyledWriter = @import("styled.zig").StyledWriter;
pub const Chart = @import("chart.zig").Chart;
pub const ChartType = @import("chart.zig").ChartType;
pub const ChartFont = @import("chart.zig").ChartFont;
pub const ChartLine = @import("chart.zig").ChartLine;
pub const ChartFill = @import("chart.zig").ChartFill;
pub const ChartPattern = @import("chart.zig").ChartPattern;
pub const Colors = @import("colors.zig").Colors;
pub const ChartSeries = @import("chart.zig").ChartSeries;
pub const ChartDataLabel = @import("chart.zig").ChartDataLabel;
pub const DataLabelOptions = @import("chart.zig").DataLabelOptions;
pub const Date = @import("date_time.zig").Date;
pub const Time = @import("date_time.zig").Time;
pub const DateTime = @import("date_time.zig").DateTime;
pub const DiagonalType = @import("format.zig").DiagonalType;
pub const CommentOptions = @import("comment.zig").CommentOptions;
pub const TmpFile = @import("mktmp.zig").TmpFile;
pub const chart = @import("chart.zig");
pub const Chartsheet = @import("chartsheet.zig").Chartsheet;
pub const ConditionalFormat = @import("conditional_format.zig").ConditionalFormat;
pub const cf = @import("conditional_format.zig");

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
