const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const c = xlsxwriter.c;

/// Error type specific to the Excellent library
pub const XlsxError = error{
    MemoryError,
    FileCreateError,
    FileWriteError,
    InvalidParameter,
    SheetError,
    FormatError,
    StringError,
    ImageError,
    ChartError,
    GenericError,
};

/// Helper function to translate C error codes to Zig errors
pub fn translateErrorCode(error_code: c_uint) XlsxError {
    return switch (error_code) {
        1 => XlsxError.MemoryError,
        2 => XlsxError.FileCreateError,
        3 => XlsxError.FileWriteError,
        4 => XlsxError.InvalidParameter,
        5 => XlsxError.SheetError,
        6 => XlsxError.FormatError,
        7 => XlsxError.StringError,
        8 => XlsxError.ImageError,
        9 => XlsxError.ChartError,
        else => XlsxError.GenericError,
    };
}
