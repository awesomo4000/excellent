const xlswriter = @import("xlsxwriter");

pub const DateTime = struct {
    year: u16 = 0,
    month: u8 = 0,
    day: u8 = 0,
    hour: u8 = 0,
    min: u8 = 0,
    sec: f64 = 0,

    pub fn toC(self: DateTime) xlswriter.lxw_datetime {
        return xlswriter.lxw_datetime{
            .year = self.year,
            .month = self.month,
            .day = self.day,
            .hour = self.hour,
            .min = self.min,
            .sec = self.sec,
        };
    }
};
