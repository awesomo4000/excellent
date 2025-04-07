const xlsxwriter = @import("xlsxwriter");

// Export colors for easy use
pub const Colors = struct {
    pub const black: u32 = xlsxwriter.LXW_COLOR_BLACK;
    pub const blue: u32 = xlsxwriter.LXW_COLOR_BLUE;
    pub const brown: u32 = xlsxwriter.LXW_COLOR_BROWN;
    pub const conifer: u32 = 0x92d050;
    pub const atlantis: u32 = 0x92d050;
    pub const deepskyblue: u32 = 0x00b0f0;
    pub const cyan: u32 = xlsxwriter.LXW_COLOR_CYAN;
    pub const gray: u32 = xlsxwriter.LXW_COLOR_GRAY;
    pub const green: u32 = xlsxwriter.LXW_COLOR_GREEN;
    pub const jade: u32 = 0x00b050;
    pub const lime: u32 = xlsxwriter.LXW_COLOR_LIME;
    pub const magenta: u32 = xlsxwriter.LXW_COLOR_MAGENTA;
    pub const navy: u32 = xlsxwriter.LXW_COLOR_NAVY;
    pub const orange: u32 = xlsxwriter.LXW_COLOR_ORANGE;
    pub const pink: u32 = xlsxwriter.LXW_COLOR_PINK;
    pub const royalpurple: u32 = 0x7030a0;
    pub const purple: u32 = xlsxwriter.LXW_COLOR_PURPLE;
    pub const red: u32 = xlsxwriter.LXW_COLOR_RED;
    pub const silver: u32 = xlsxwriter.LXW_COLOR_SILVER;
    pub const white: u32 = xlsxwriter.LXW_COLOR_WHITE;
    pub const yellow: u32 = xlsxwriter.LXW_COLOR_YELLOW;
    pub const aztecgold: u32 = 0xc68c53;
    pub const saddlebrown: u32 = 0x804000;
    pub const traditionalchocolate: u32 = 0x804000;
    pub const harlockscape: u32 = 0xb30000;
    pub const artfulred: u32 = 0xb30000;
    pub const pompelmo: u32 = 0xff6666;
};
