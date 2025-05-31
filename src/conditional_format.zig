const std = @import("std");
const Format = @import("format.zig").Format;
const c = @import("xlsxwriter");

// const CellCriteria = enum {
//     equal_to,
//     not_equal_to,
//     greater_than,
//     less_than,
//     greater_than_or_equal_to,
//     less_than_or_equal_to,
// };

// const TimePeriodCriteria = enum {
//     yesterday,
//     today,
//     tomorrow,
//     last_week,
//     this_week,
// };

// const TextCriteria = enum {
//     contains,
//     not_contains,
//     starts_with,
//     ends_with,
// };

// const AverageCriteria = enum {
//     above,
//     below,
//     above_or_equal,
//     below_or_equal,
// };

pub const ConditionalFormat = struct {
    inner: c.lxw_conditional_format,
    pub fn blank() ConditionalFormat {
        return ConditionalFormat{
            .inner = std.mem.zeroes(c.lxw_conditional_format),
        };
    }

    pub fn setStopIfTrue(
        self: *ConditionalFormat,
        stop_if_true: bool,
    ) void {
        self.inner.stop_if_true =
            if (stop_if_true) c.LXW_TRUE else c.LXW_FALSE;
    }

    pub fn setFormat(
        self: *ConditionalFormat,
        format: ?*Format,
    ) void {
        self.inner.format = format.?.format;
    }
    pub fn deinit(self: *ConditionalFormat) void {
        self.inner.format = null;
    }
};

pub fn cellEqual(
    value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_EQUAL_TO;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn cellNotEqual(
    value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn cellGreaterThan(
    value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria =
        c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn cellGreaterThanOrEqualTo(
    value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria =
        c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn cellLessThan(
    value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_LESS_THAN;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn cellLessThanOrEqualTo(
    value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria =
        c.LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn cellBetween(
    min_value: f64,
    max_value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_BETWEEN;
    f.inner.min_value = min_value;
    f.inner.max_value = max_value;
    f.setFormat(format);
    return f;
}

pub fn cellNotBetween(
    min_value: f64,
    max_value: f64,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_CELL;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN;
    f.inner.min_value = min_value;
    f.inner.max_value = max_value;
    f.setFormat(format);
    return f;
}

pub fn timePeriodYesterday(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_YESTERDAY;
    f.setFormat(format);
    return f;
}

pub fn timePeriodToday(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_TODAY;
    f.setFormat(format);
    return f;
}

pub fn timePeriodTomorrow(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_TOMORROW;
    f.setFormat(format);
    return f;
}

pub fn timePeriodLastWeek(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_LAST_WEEK;
    f.setFormat(format);
    return f;
}

pub fn timePeriodThisWeek(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_THIS_WEEK;
    f.setFormat(format);
    return f;
}

pub fn timePeriodNextWeek(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_NEXT_WEEK;
    f.setFormat(format);
    return f;
}

pub fn timePeriodLastMonth(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_LAST_MONTH;
    f.setFormat(format);
    return f;
}

pub fn textContains(
    text: []const u8,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TEXT;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_CONTAINS;
    f.inner.value = text;
    f.setFormat(format);
    return f;
}

pub fn textNotContaining(
    text: []const u8,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TEXT;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_NOT_CONTAINS;
    f.inner.value = text;
    f.setFormat(format);
    return f;
}

pub fn textStartsWith(
    text: []const u8,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TEXT;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_STARTS_WITH;
    f.inner.value = text;
    f.setFormat(format);
    return f;
}

pub fn textEndsWith(
    text: []const u8,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TEXT;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_ENDS_WITH;
    f.inner.value = text;
    f.setFormat(format);
    return f;
}

pub fn averageAbove(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE;
    f.setFormat(format);
    return f;
}

pub fn averageBelow(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW;
    f.setFormat(format);
    return f;
}

pub fn averageAboveOrEqual(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria =
        c.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL_TO;
    f.setFormat(format);
    return f;
}

pub fn averageBelowOrEqual(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria =
        c.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL_TO;
    f.setFormat(format);
    return f;
}

pub fn average_1_StdDevAbove(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_1_STD_DEV_ABOVE;
    f.setFormat(format);
    return f;
}

pub fn average_1_StdDevBelow(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_1_STD_DEV_BELOW;
    f.setFormat(format);
    return f;
}

pub fn average_2_StdDevAbove(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_2_STD_DEV_ABOVE;
    f.setFormat(format);
    return f;
}

pub fn average_2_StdDevBelow(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_2_STD_DEV_BELOW;
    f.setFormat(format);
    return f;
}

pub fn average_3_StdDevAbove(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_3_STD_DEV_ABOVE;
    f.setFormat(format);
    return f;
}

pub fn average_3_StdDevBelow(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_AVERAGE;
    f.inner.criteria = c.LXW_CONDITIONAL_CRITERIA_3_STD_DEV_BELOW;
    f.setFormat(format);
    return f;
}

pub fn duplicate(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_DUPLICATE;
    f.setFormat(format);
    return f;
}

pub fn unique(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_UNIQUE;
    f.setFormat(format);
    return f;
}

pub fn top(value: f64, format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_TOP;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn bottom(value: f64, format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_BOTTOM;
    f.inner.value = value;
    f.setFormat(format);
    return f;
}

pub fn blank(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_BLANKS;
    f.setFormat(format);
    return f;
}

pub fn notBlank(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_NOT_BLANKS;
    f.setFormat(format);
    return f;
}

pub fn errors(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_ERRORS;
    f.setFormat(format);
    return f;
}

pub fn notErrors(format: ?*Format) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_NOT_ERRORS;
    f.setFormat(format);
    return f;
}

pub fn formula(
    formula_str: []const u8,
    format: ?*Format,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_FORMULA;
    f.inner.formula = formula_str;
    f.setFormat(format);
    return f;
}

pub const TwoColorScaleOpts = struct {
    min_color: u32,
    max_color: u32,
};

pub fn twoColorScale(opts: ?TwoColorScaleOpts) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_2_COLOR_SCALE;
    if (opts) |o| {
        f.inner.min_color = o.min_color;
        f.inner.max_color = o.max_color;
    }
    return f;
}

pub const ThreeColorScaleOpts = struct {
    min_color: u32,
    mid_color: u32,
    max_color: u32,
};

pub fn threeColorScale(opts: ?ThreeColorScaleOpts) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_3_COLOR_SCALE;
    if (opts) |o| {
        f.inner.min_color = o.min_color;
        f.inner.mid_color = o.mid_color;
        f.inner.max_color = o.max_color;
    }
    return f;
}

pub const BarAxisPosition = enum {
    automatic,
    midpoint,
    none,

    pub fn optToC(
        self: ?BarAxisPosition,
    ) c.lxw_conditional_bar_axis_position {
        if (self) |b| return b.toC();
        return c.LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC;
    }

    pub fn toC(self: BarAxisPosition) u8 {
        return switch (self) {
            .automatic => c.LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC,
            .midpoint => c.LXW_CONDITIONAL_BAR_AXIS_MIDPOINT,
            .none => c.LXW_CONDITIONAL_BAR_AXIS_NONE,
        };
    }
};

pub const BarDirection = enum {
    context,
    right_to_left,
    left_to_right,

    pub fn optToC(
        self: ?BarDirection,
    ) c.lxw_conditional_format_bar_direction {
        if (self) |b| return b.toC();
        return c.LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT;
    }

    pub fn toC(self: BarDirection) u8 {
        return switch (self) {
            .context => c.LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT,
            .right_to_left => c.LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT,
            .left_to_right => c.LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT,
        };
    }
};

pub const RuleType = enum {
    none,
    minimum,
    maximum,
    number,
    percent,
    percentile,
    formula,
    auto_min,
    auto_max,

    pub fn toC(self: RuleType) u8 {
        return switch (self) {
            .none => c.LXW_CONDITIONAL_RULE_TYPE_NONE,
            .minimum => c.LXW_CONDITIONAL_RULE_TYPE_MINIMUM,
            .maximum => c.LXW_CONDITIONAL_RULE_TYPE_MAXIMUM,
            .number => c.LXW_CONDITIONAL_RULE_TYPE_NUMBER,
            .percent => c.LXW_CONDITIONAL_RULE_TYPE_PERCENT,
            .percentile => c.LXW_CONDITIONAL_RULE_TYPE_PERCENTILE,
            .formula => c.LXW_CONDITIONAL_RULE_TYPE_FORMULA,
            .auto_min => c.LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN,
            .auto_max => c.LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX,
        };
    }
};

pub const DataBarOpts = struct {
    axis_color: ?u32 = null,
    axis_position: ?BarAxisPosition = null,
    border_color: ?u32 = null,
    color: ?u32 = null,
    direction: ?BarDirection = null,
    negative_border_color: ?u32 = null,
    negative_border_color_same: ?bool = null,
    negative_color: ?u32 = null,
    negative_color_same: ?bool = null,
    no_border: ?bool = null,
    bar_only: ?bool = null,
    solid: ?bool = null,
    excel_2010_style: ?bool = null,
    max_rule_type: ?RuleType = null,
    max_value: ?f64 = null,
    max_value_string: ?[]const u8 = null,
    min_rule_type: ?RuleType = null,
    min_value: ?f64 = null,
    min_value_string: ?[]const u8 = null,
};

pub fn optBoolToC(value: ?bool) u8 {
    if (value) |v| return if (v) c.LXW_TRUE else c.LXW_FALSE;
    return c.LXW_FALSE;
}

pub fn dataBar(opts: ?DataBarOpts) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_DATA_BAR;
    if (opts == null) return f;
    const o = opts.?;
    f.inner.bar_only = optBoolToC(o.bar_only);
    f.inner.bar_color = o.color orelse c.LXW_COLOR_UNSET;
    f.inner.bar_border_color = o.border_color orelse c.LXW_COLOR_UNSET;
    f.inner.bar_no_border = optBoolToC(o.no_border);
    f.inner.bar_solid = optBoolToC(o.solid);
    if (o.direction != null) {
        f.inner.bar_direction = o.direction.?.toC();
    }
    if (o.axis_position != null) {
        f.inner.bar_axis_position = o.axis_position.?.toC();
    }
    f.inner.bar_axis_color = o.axis_color orelse c.LXW_COLOR_UNSET;
    f.inner.bar_negative_color =
        o.negative_color orelse c.LXW_COLOR_UNSET;
    f.inner.bar_negative_border_color =
        o.negative_border_color orelse c.LXW_COLOR_UNSET;
    f.inner.bar_negative_color_same =
        optBoolToC(o.negative_color_same);
    f.inner.bar_negative_border_color_same =
        optBoolToC(o.negative_border_color_same);
    f.inner.data_bar_2010 = optBoolToC(o.excel_2010_style);
    if (o.max_rule_type != null) {
        f.inner.max_rule_type = o.max_rule_type.?.toC();
    }
    if (o.max_value != null) {
        f.inner.max_value = o.max_value.?;
    }
    const max_val_str_ptr = if (o.max_value_string) |s| s.ptr else null;
    if (max_val_str_ptr != null) {
        f.inner.max_value_string = max_val_str_ptr;
    }
    if (o.min_rule_type != null) {
        f.inner.min_rule_type = o.min_rule_type.?.toC();
    }
    if (o.min_value != null) {
        f.inner.min_value = o.min_value.?;
    }
    const min_val_str_ptr = if (o.min_value_string) |s| s.ptr else null;
    if (min_val_str_ptr != null) {
        f.inner.min_value_string = min_val_str_ptr;
    }
    return f;
}

pub const IconSetStyle = enum {
    three_arrows_colored,
    three_arrows_gray,
    three_flags,
    three_traffic_lights_unrimmed,
    three_traffic_lights_rimmed,
    three_signs,
    three_symbols_circled,
    three_symbols_uncircled,
    four_arrows_colored,
    four_arrows_gray,
    four_red_to_black,
    four_ratings,
    four_traffic_lights,
    five_arrows_colored,
    five_arrows_gray,
    five_ratings,
    five_quarters,

    fn toC(self: IconSetStyle) u8 {
        return switch (self) {
            .three_arrows_colored => c.LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED,
            .three_arrows_gray => c.LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY,
            .three_flags => c.LXW_CONDITIONAL_ICONS_3_FLAGS,
            .three_traffic_lights_unrimmed => c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED,
            .three_traffic_lights_rimmed => c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED,
            .three_signs => c.LXW_CONDITIONAL_ICONS_3_SIGNS,
            .three_symbols_circled => c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED,
            .three_symbols_uncircled => c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED,
            .four_arrows_colored => c.LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED,
            .four_arrows_gray => c.LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY,
            .four_red_to_black => c.LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK,
            .four_ratings => c.LXW_CONDITIONAL_ICONS_4_RATINGS,
            .four_traffic_lights => c.LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS,
            .five_arrows_colored => c.LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED,
            .five_arrows_gray => c.LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY,
            .five_ratings => c.LXW_CONDITIONAL_ICONS_5_RATINGS,
            .five_quarters => c.LXW_CONDITIONAL_ICONS_5_QUARTERS,
        };
    }
};

pub const IconSetOpts = struct {
    icon_style: ?IconSetStyle = null,
    icons_only: ?bool = null,
    reverse_icons: ?bool = null,
};

pub fn iconSet(
    format: ?*Format,
    opts: ?IconSetOpts,
) ConditionalFormat {
    var f = ConditionalFormat.blank();
    f.inner.type = c.LXW_CONDITIONAL_TYPE_ICON_SETS;
    if (format != null) f.setFormat(format);
    if (opts == null) return f;
    const o = opts.?;
    if (o.icon_style != null) {
        f.inner.icon_style = o.icon_style.?.toC();
    }
    f.inner.icons_only = optBoolToC(o.icons_only);
    f.inner.reverse_icons = optBoolToC(o.reverse_icons);

    return f;
}

// Apply conditional formatting to a range of cells
// pub fn conditionalFormat(
//     self: *Worksheet,
//     range: []const u8,
//     options: ConditionalFormatOptions,
// ) !void {
//     // Split the range into start and end cells
//     var iter = std.mem.splitScalar(u8, range, ':');
//     const start = iter.next() orelse return error.InvalidRange;
//     const end = iter.next() orelse return error.InvalidRange;

//     // Parse start and end cells
//     const start_pos = try cell_utils.cell.strToRowCol(start);
//     const end_pos = try cell_utils.cell.strToRowCol(end);

//     // Create a conditional format object
//     var conditional_format: c.lxw_conditional_format = std.mem.zeroes(c.lxw_conditional_format);

//     // Set the type
//     conditional_format.type = switch (options.type) {
//         .cell => c.LXW_CONDITIONAL_TYPE_CELL,
//         .duplicate => c.LXW_CONDITIONAL_TYPE_DUPLICATE,
//         .unique => c.LXW_CONDITIONAL_TYPE_UNIQUE,
//         .top => c.LXW_CONDITIONAL_TYPE_TOP,
//         .bottom => c.LXW_CONDITIONAL_TYPE_BOTTOM,
//         .average => c.LXW_CONDITIONAL_TYPE_AVERAGE,
//         .two_color_scale => c.LXW_CONDITIONAL_2_COLOR_SCALE,
//         .three_color_scale => c.LXW_CONDITIONAL_3_COLOR_SCALE,
//         .data_bar => c.LXW_CONDITIONAL_DATA_BAR,
//         .icon_sets => c.LXW_CONDITIONAL_TYPE_ICON_SETS,
//     };

//     // Set the criteria if provided
//     if (options.criteria) |criteria| {
//         conditional_format.criteria = switch (criteria) {
//             .between => c.LXW_CONDITIONAL_CRITERIA_BETWEEN,
//             .not_between => c.LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN,
//             .equal_to => c.LXW_CONDITIONAL_CRITERIA_EQUAL_TO,
//             .not_equal_to => c.LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO,
//             .greater_than => c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN,
//             .less_than => c.LXW_CONDITIONAL_CRITERIA_LESS_THAN,
//             .greater_than_or_equal_to => c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
//             .less_than_or_equal_to => c.LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO,
//             .average_above => c.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE,
//             .average_below => c.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW,
//         };
//     }

//     // Set the value if provided
//     if (options.value) |value| {
//         conditional_format.value = value;
//     }

//     // Set min/max values if provided
//     if (options.min_value) |min_value| {
//         conditional_format.min_value = min_value;
//     }
//     if (options.max_value) |max_value| {
//         conditional_format.max_value = max_value;
//     }

//     // Set the format if provided
//     if (options.format) |format| {
//         conditional_format.format = format.format;
//     }

//     // Set data bar options if this is a data bar
//     if (options.type == .data_bar) {
//         conditional_format.bar_only = if (options.bar_only) c.LXW_TRUE else c.LXW_FALSE;
//         if (options.bar_color) |color| {
//             conditional_format.bar_color = color;
//         }
//         conditional_format.bar_solid = if (options.bar_solid) c.LXW_TRUE else c.LXW_FALSE;
//         conditional_format.bar_direction = switch (options.bar_direction) {
//             .left_to_right => c.LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT,
//             .right_to_left => c.LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT,
//         };
//         conditional_format.data_bar_2010 = if (options.data_bar_2010) c.LXW_TRUE else c.LXW_FALSE;
//         conditional_format.bar_negative_color_same = if (options.bar_negative_color_same) c.LXW_TRUE else c.LXW_FALSE;
//         conditional_format.bar_negative_border_color_same = if (options.bar_negative_border_color_same) c.LXW_TRUE else c.LXW_FALSE;
//     }

//     // Set icon set options if this is an icon set
//     if (options.type == .icon_sets and options.icon_style) |style| {
//         conditional_format.icon_style = switch (style) {
//             .three_traffic_lights_rimmed => c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED,
//             .three_traffic_lights_unrimmed => c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED,
//             .three_arrows_colored => c.LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED,
//             .three_arrows_gray => c.LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY,
//             .three_flags => c.LXW_CONDITIONAL_ICONS_3_FLAGS,
//             .three_signs => c.LXW_CONDITIONAL_ICONS_3_SIGNS,
//             .three_symbols_circled => c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED,
//             .three_symbols_uncircled => c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED,
//             .three_stars => c.LXW_CONDITIONAL_ICONS_3_STARS,
//             .three_triangles => c.LXW_CONDITIONAL_ICONS_3_TRIANGLES,
//             .four_traffic_lights => c.LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS,
//             .four_arrows_colored => c.LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED,
//             .four_arrows_gray => c.LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY,
//             .four_red_to_black => c.LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK,
//             .four_ratings => c.LXW_CONDITIONAL_ICONS_4_RATINGS,
//             .four_traffic_lights_rimmed => c.LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS_RIMMED,
//             .five_arrows_colored => c.LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED,
//             .five_arrows_gray => c.LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY,
//             .five_ratings => c.LXW_CONDITIONAL_ICONS_5_RATINGS,
//             .five_quarters => c.LXW_CONDITIONAL_ICONS_5_QUARTERS,
//         };
//         conditional_format.icons_only = if (options.icons_only) c.LXW_TRUE else c.LXW_FALSE;
//         conditional_format.reverse_icons = if (options.reverse_icons) c.LXW_TRUE else c.LXW_FALSE;
//     }

//     // Apply the conditional format
//     const result = c.worksheet_conditional_format_range(
//         self.worksheet,
//         @intCast(start_pos.row),
//         @intCast(start_pos.col),
//         @intCast(end_pos.row),
//         @intCast(end_pos.col),
//         &conditional_format,
//     );
//     if (result != c.LXW_NO_ERROR) return error.WriteFailed;
// }
