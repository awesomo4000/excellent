const xlsxwriter = @import("xlsxwriter");
const c = xlsxwriter.c;

pub const Date = struct {
    year: u16 = 0,
    month: u8 = 0,
    day: u8 = 0,

    pub fn toDateTime(self: Date) DateTime {
        return DateTime{
            .year = self.year,
            .month = self.month,
            .day = self.day,
            .hour = 0,
            .minute = 0,
            .second = 0,
        };
    }
};

pub const Time = struct {
    hour: u8 = 0,
    minute: u8 = 0,
    second: f64 = 0,

    pub fn toDateTime(self: Time) DateTime {
        return DateTime{
            .year = 0,
            .month = 0,
            .day = 0,
            .hour = self.hour,
            .minute = self.minute,
            .second = self.second,
        };
    }
};

pub const DateTime = struct {
    year: u16 = 0,
    month: u8 = 0,
    day: u8 = 0,
    hour: u8 = 0,
    minute: u8 = 0,
    second: f64 = 0,

    pub fn toC(self: DateTime) c.lxw_datetime {
        return c.lxw_datetime{
            .year = self.year,
            .month = self.month,
            .day = self.day,
            .hour = self.hour,
            .min = self.minute,
            .sec = self.second,
        };
    }
};
