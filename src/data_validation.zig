const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
const c = xlsxwriter.c;
const DateTime = @import("date_time.zig").DateTime;
const DateT = @import("date_time.zig").Date;
const TimeT = @import("date_time.zig").Time;

/// Data validation types supported by Excel
pub const ValidationType = enum {
    integer,
    integer_formula,
    decimal,
    list,
    list_formula,
    date,
    time,
    length,
    custom_formula,
};

/// Data validation criteria for comparisons
pub const Criteria = enum {
    between,
    not_between,
    equal_to,
    not_equal_to,
    greater_than,
    less_than,
    greater_than_or_equal_to,
    less_than_or_equal_to,
};

/// Error message types for data validation
pub const ErrorType = enum {
    stop,
    warning,
    information,
};

pub const Options = struct {
    ignore_blank: bool = true,
    dropdown: bool = true,
    input_title: ?[]const u8 = null,
    input_message: ?[]const u8 = null,
    error_title: ?[]const u8 = null,
    error_message: ?[]const u8 = null,
    error_type: ErrorType = .stop,
};

pub const Integer = struct {
    pub fn gt(value: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .greater_than,
            .value_number = @as(f64, @floatFromInt(value)),
            .options = options,
        };
    }

    pub fn gte(value: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .greater_than_or_equal_to,
            .value_number = @as(f64, @floatFromInt(value)),
            .options = options,
        };
    }

    pub fn lt(value: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .less_than,
            .value_number = @as(f64, @floatFromInt(value)),
            .options = options,
        };
    }

    pub fn lte(value: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .less_than_or_equal_to,
            .value_number = @as(f64, @floatFromInt(value)),
            .options = options,
        };
    }

    pub fn between(minimum: i64, maximum: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .between,
            .minimum_number = @as(f64, @floatFromInt(minimum)),
            .maximum_number = @as(f64, @floatFromInt(maximum)),
            .options = options,
        };
    }

    pub fn not_between(minimum: i64, maximum: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .not_between,
            .minimum_number = @as(f64, @floatFromInt(minimum)),
            .maximum_number = @as(f64, @floatFromInt(maximum)),
            .options = options,
        };
    }

    pub fn eql(value: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .equal_to,
            .value_number = @as(f64, @floatFromInt(value)),
            .options = options,
        };
    }

    pub fn not_eql(value: i64, options: Options) Validation {
        return Validation{
            .validate = .integer,
            .criteria = .not_equal_to,
            .value_number = @as(f64, @floatFromInt(value)),
            .options = options,
        };
    }
};

pub const IntegerFormula = struct {
    pub fn between(
        minimum: []const u8,
        maximum: []const u8,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .between,
            .minimum_formula = minimum,
            .maximum_formula = maximum,
            .options = options,
        };
    }

    pub fn not_between(
        minimum: []const u8,
        maximum: []const u8,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .not_between,
            .minimum_formula = minimum,
            .maximum_formula = maximum,
            .options = options,
        };
    }

    pub fn gt(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .greater_than,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn gte(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .greater_than_or_equal_to,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn lt(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .less_than,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn lte(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .less_than_or_equal_to,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn eql(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .equal_to,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn not_eql(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .integer_formula,
            .criteria = .not_equal_to,
            .value_formula = value,
            .options = options,
        };
    }
};

pub const Decimal = struct {
    pub fn between(
        minimum: f64,
        maximum: f64,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .between,
            .minimum_number = minimum,
            .maximum_number = maximum,
            .options = options,
        };
    }

    pub fn not_between(
        minimum: f64,
        maximum: f64,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .not_between,
            .minimum_number = minimum,
            .maximum_number = maximum,
            .options = options,
        };
    }

    pub fn gt(value: f64, options: Options) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .greater_than,
            .value_number = value,
            .options = options,
        };
    }

    pub fn gte(value: f64, options: Options) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .greater_than_or_equal_to,
            .value_number = value,
            .options = options,
        };
    }

    pub fn lte(value: f64, options: Options) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .less_than_or_equal_to,
            .value_number = value,
            .options = options,
        };
    }

    pub fn lt(value: f64, options: Options) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .less_than,
            .value_number = value,
            .options = options,
        };
    }

    pub fn eql(value: f64, options: Options) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .equal_to,
            .value_number = value,
            .options = options,
        };
    }

    pub fn not_eql(value: f64, options: Options) Validation {
        return Validation{
            .validate = .decimal,
            .criteria = .not_equal_to,
            .value_number = value,
            .options = options,
        };
    }
};

pub const DecimalFormula = struct {
    pub fn gt(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .greater_than,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn gte(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .greater_than_or_equal_to,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn lt(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .less_than,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn lte(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .less_than_or_equal_to,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn eql(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .equal_to,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn not_eql(value: []const u8, options: Options) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .not_equal_to,
            .value_formula = value,
            .options = options,
        };
    }

    pub fn not_between(
        minimum: []const u8,
        maximum: []const u8,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .not_between,
            .minimum_formula = minimum,
            .maximum_formula = maximum,
            .options = options,
        };
    }

    pub fn between(
        minimum: []const u8,
        maximum: []const u8,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .decimal_formula,
            .criteria = .between,
            .minimum_formula = minimum,
            .maximum_formula = maximum,
            .options = options,
        };
    }
};

pub fn List(value_list: []const []const u8, options: Options) Validation {
    return Validation{
        .validate = .list,
        .value_list = value_list,
        .options = options,
    };
}

pub fn ListFormula(value_formula: []const u8, options: Options) Validation {
    return Validation{
        .validate = .list_formula,
        .value_formula = value_formula,
        .options = options,
    };
}

pub const Date = struct {
    pub fn between(
        minimum: DateT,
        maximum: DateT,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .date,
            .criteria = .between,
            .minimum_datetime = minimum.toDateTime(),
            .maximum_datetime = maximum.toDateTime(),
            .options = options,
        };
    }

    pub fn not_between(
        minimum: DateT,
        maximum: DateT,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .date,
            .criteria = .not_between,
            .minimum_datetime = minimum.toDateTime(),
            .maximum_datetime = maximum.toDateTime(),
            .options = options,
        };
    }

    pub fn gt(value: DateT, options: Options) Validation {
        return Validation{
            .validate = .date,
            .criteria = .greater_than,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn gte(value: DateT, options: Options) Validation {
        return Validation{
            .validate = .date,
            .criteria = .greater_than_or_equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn lt(value: DateT, options: Options) Validation {
        return Validation{
            .validate = .date,
            .criteria = .less_than,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn lte(value: DateT, options: Options) Validation {
        return Validation{
            .validate = .date,
            .criteria = .less_than_or_equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn eql(value: DateT, options: Options) Validation {
        return Validation{
            .validate = .date,
            .criteria = .equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn not_eql(value: DateT, options: Options) Validation {
        return Validation{
            .validate = .date,
            .criteria = .not_equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }
};

pub const Time = struct {
    pub fn between(
        minimum: TimeT,
        maximum: TimeT,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .time,
            .criteria = .between,
            .minimum_datetime = minimum.toDateTime(),
            .maximum_datetime = maximum.toDateTime(),
            .options = options,
        };
    }

    pub fn not_between(
        minimum: TimeT,
        maximum: TimeT,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .time,
            .criteria = .not_between,
            .minimum_datetime = minimum.toDateTime(),
            .maximum_datetime = maximum.toDateTime(),
            .options = options,
        };
    }

    pub fn gt(value: TimeT, options: Options) Validation {
        return Validation{
            .validate = .time,
            .criteria = .greater_than,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn gte(value: TimeT, options: Options) Validation {
        return Validation{
            .validate = .time,
            .criteria = .greater_than_or_equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn lt(value: TimeT, options: Options) Validation {
        return Validation{
            .validate = .time,
            .criteria = .less_than,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn lte(value: TimeT, options: Options) Validation {
        return Validation{
            .validate = .time,
            .criteria = .less_than_or_equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn eql(value: TimeT, options: Options) Validation {
        return Validation{
            .validate = .time,
            .criteria = .equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }

    pub fn not_eql(value: TimeT, options: Options) Validation {
        return Validation{
            .validate = .time,
            .criteria = .not_equal_to,
            .value_datetime = value.toDateTime(),
            .options = options,
        };
    }
};

pub const Length = struct {
    pub fn between(
        minimum: u32,
        maximum: u32,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .length,
            .criteria = .between,
            .minimum_length = minimum,
            .maximum_length = maximum,
            .options = options,
        };
    }

    pub fn not_between(
        minimum: u32,
        maximum: u32,
        options: Options,
    ) Validation {
        return Validation{
            .validate = .length,
            .criteria = .not_between,
            .minimum_length = minimum,
            .maximum_length = maximum,
            .options = options,
        };
    }

    pub fn gt(value: u32, options: Options) Validation {
        return Validation{
            .validate = .length,
            .criteria = .greater_than,
            .value_length = value,
            .options = options,
        };
    }

    pub fn gte(value: u32, options: Options) Validation {
        return Validation{
            .validate = .length,
            .criteria = .greater_than_or_equal_to,
            .value_length = value,
            .options = options,
        };
    }

    pub fn lt(value: u32, options: Options) Validation {
        return Validation{
            .validate = .length,
            .criteria = .less_than,
            .value_length = value,
            .options = options,
        };
    }

    pub fn lte(value: u32, options: Options) Validation {
        return Validation{
            .validate = .length,
            .criteria = .less_than_or_equal_to,
            .value_length = value,
            .options = options,
        };
    }

    pub fn eql(value: u32, options: Options) Validation {
        return Validation{
            .validate = .length,
            .criteria = .equal_to,
            .value_length = value,
            .options = options,
        };
    }

    pub fn not_eql(value: u32, options: Options) Validation {
        return Validation{
            .validate = .length,
            .criteria = .not_equal_to,
            .value_length = value,
            .options = options,
        };
    }
};

pub fn CustomFormula(
    formula: []const u8,
    options: Options,
) Validation {
    return Validation{
        .validate = .custom_formula,
        .value_formula = formula,
        .options = options,
    };
}

/// Data validation configuration
pub const Validation = struct {
    validate: ValidationType,
    criteria: ?Criteria = null,
    // Numeric values
    minimum_number: ?f64 = null,
    maximum_number: ?f64 = null,
    value_number: ?f64 = null,

    // Formula values
    minimum_formula: ?[]const u8 = null,
    maximum_formula: ?[]const u8 = null,
    value_formula: ?[]const u8 = null,

    // List values
    value_list: ?[]const []const u8 = null,

    // Date/time values
    minimum_datetime: ?DateTime = null,
    maximum_datetime: ?DateTime = null,

    // Length values (stored as numbers in C API)
    minimum_length: ?u32 = null,
    maximum_length: ?u32 = null,
    value_length: ?u32 = null,

    // Messages
    options: Options,

    /// Convert DataValidation to C library format
    pub fn toC(
        self: Validation,
        allocator: std.mem.Allocator,
    ) !*c.lxw_data_validation {
        var data_validation = try allocator.create(c.lxw_data_validation);

        // Initialize with zeros
        @memset(
            @as([*]u8, @ptrCast(data_validation))[0..@sizeOf(c.lxw_data_validation)],
            0,
        );

        // Set validation type
        data_validation.validate = switch (self.validate) {
            .integer => c.LXW_VALIDATION_TYPE_INTEGER,
            .integer_formula => c.LXW_VALIDATION_TYPE_INTEGER_FORMULA,
            .decimal => c.LXW_VALIDATION_TYPE_DECIMAL,
            .list => c.LXW_VALIDATION_TYPE_LIST,
            .list_formula => c.LXW_VALIDATION_TYPE_LIST_FORMULA,
            .date => c.LXW_VALIDATION_TYPE_DATE,
            .time => c.LXW_VALIDATION_TYPE_TIME,
            .length => c.LXW_VALIDATION_TYPE_LENGTH,
            .custom_formula => c.LXW_VALIDATION_TYPE_CUSTOM_FORMULA,
        };

        // Set criteria if provided
        if (self.criteria) |criteria| {
            data_validation.criteria = switch (criteria) {
                .between => c.LXW_VALIDATION_CRITERIA_BETWEEN,
                .not_between => c.LXW_VALIDATION_CRITERIA_NOT_BETWEEN,
                .equal_to => c.LXW_VALIDATION_CRITERIA_EQUAL_TO,
                .not_equal_to => c.LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO,
                .greater_than => c.LXW_VALIDATION_CRITERIA_GREATER_THAN,
                .less_than => c.LXW_VALIDATION_CRITERIA_LESS_THAN,
                .greater_than_or_equal_to => c.LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
                .less_than_or_equal_to => c.LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO,
            };
        }

        // Set numeric values
        if (self.minimum_number) |min| data_validation.minimum_number = min;
        if (self.maximum_number) |max| data_validation.maximum_number = max;
        if (self.value_number) |val| data_validation.value_number = val;

        // Set formulas (need null-terminated strings)
        if (self.minimum_formula) |formula| {
            const null_term = try allocator.dupeZ(u8, formula);
            data_validation.minimum_formula = null_term.ptr;
        }
        if (self.maximum_formula) |formula| {
            const null_term = try allocator.dupeZ(u8, formula);
            data_validation.maximum_formula = null_term.ptr;
        }
        if (self.value_formula) |formula| {
            const null_term = try allocator.dupeZ(u8, formula);
            data_validation.value_formula = null_term.ptr;
        }

        // Set list values
        if (self.value_list) |list_values| {
            var c_list = try allocator.alloc(
                [*c]const u8,
                list_values.len + 1,
            );
            for (list_values, 0..) |item, i| {
                const null_term = try allocator.dupeZ(u8, item);
                c_list[i] = null_term.ptr;
            }
            c_list[list_values.len] = null; // Null-terminate the array
            data_validation.value_list = @ptrCast(c_list.ptr);
        }

        // Set list formula
        if (self.value_formula) |formula| {
            const null_term = try allocator.dupeZ(u8, formula);
            data_validation.value_formula = null_term.ptr;
        }

        // Set datetime values
        if (self.minimum_datetime) |dt| {
            data_validation.minimum_datetime = c.lxw_datetime{
                .year = dt.year,
                .month = dt.month,
                .day = dt.day,
                .hour = dt.hour,
                .min = dt.minute,
                .sec = dt.second,
            };
        }
        if (self.maximum_datetime) |dt| {
            data_validation.maximum_datetime = c.lxw_datetime{
                .year = dt.year,
                .month = dt.month,
                .day = dt.day,
                .hour = dt.hour,
                .min = dt.minute,
                .sec = dt.second,
            };
        }

        // Set length values (C API stores these as numbers)
        if (self.minimum_length) |min| data_validation.minimum_number = @floatFromInt(min);
        if (self.maximum_length) |max| data_validation.maximum_number = @floatFromInt(max);
        if (self.value_length) |val| data_validation.value_number = @floatFromInt(val);

        // Set messages
        if (self.options.input_title) |title| {
            const null_term = try allocator.dupeZ(u8, title);
            data_validation.input_title = null_term.ptr;
        }
        if (self.options.input_message) |message| {
            const null_term = try allocator.dupeZ(u8, message);
            data_validation.input_message = null_term.ptr;
        }
        if (self.options.error_title) |title| {
            const null_term = try allocator.dupeZ(u8, title);
            data_validation.error_title = null_term.ptr;
        }
        if (self.options.error_message) |message| {
            const null_term = try allocator.dupeZ(u8, message);
            data_validation.error_message = null_term.ptr;
        }

        // Set error type
        data_validation.error_type = switch (self.options.error_type) {
            .stop => c.LXW_VALIDATION_ERROR_TYPE_STOP,
            .warning => c.LXW_VALIDATION_ERROR_TYPE_WARNING,
            .information => c.LXW_VALIDATION_ERROR_TYPE_INFORMATION,
        };

        // Set options
        data_validation.ignore_blank = if (self.options.ignore_blank) 1 else 0;
        data_validation.dropdown = if (self.options.dropdown) 1 else 0;

        return data_validation;
    }
};
