pub const xlsxError = @import("errors.zig");
pub const c = @cImport({
    @cDefine("struct_headname", "");
    @cInclude("xlsxwriter.h");
});
