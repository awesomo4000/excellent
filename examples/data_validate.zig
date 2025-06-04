//
// Examples of how to add data validation and dropdown lists using the
// excellent library.
//
// Data validation is a feature of Excel which allows you to restrict the data
// that a user enters in a cell and to display help and warning messages. It
// also allows you to restrict input to values in a dropdown list.
//

const std = @import("std");
const excel = @import("excellent");
const Valid = excel.DataValidation;
const DateTime = excel.DateTime;

// Write some data to the worksheet.
fn writeWorksheetData(
    worksheet: *excel.Worksheet,
    format: *excel.Format,
) !void {
    try worksheet.writeString(
        0,
        0,
        "Some examples of data validation in libxlsxwriter",
        format,
    );
    try worksheet.writeString(0, 1, "Enter values in this column", format);
    try worksheet.writeString(0, 3, "Sample Data", format);

    try worksheet.writeString(2, 3, "Integers", null);
    try worksheet.writeNumber(2, 4, 1, null);
    try worksheet.writeNumber(2, 5, 10, null);

    try worksheet.writeString(3, 3, "List data", null);
    try worksheet.writeString(3, 4, "open", null);
    try worksheet.writeString(3, 5, "high", null);
    try worksheet.writeString(3, 6, "close", null);

    try worksheet.writeString(4, 3, "Formula", null);
    try worksheet.writeFormula(4, 4, "=AND(F5=50,G5=60)", null);
    try worksheet.writeNumber(4, 5, 50, null);
    try worksheet.writeNumber(4, 6, 60, null);
}

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer if (gpa.deinit() == .leak) {
        std.debug.panic("leaks detected", .{});
    };
    const allocator = gpa.allocator();

    var workbook = try excel.Workbook.create(
        allocator,
        "data_validate.xlsx",
    );
    defer workbook.deinit();

    var worksheet = try workbook.addWorksheet(null);

    // Add a format to use to highlight the header cells.
    var format = try workbook.addFormat();
    _ = format.setBorder(.thin);
    _ = format.setFgColor(0xC6EFCE);
    _ = format.setBold();
    _ = format.setTextWrap();
    _ = format.setAlign(.vertical_center);
    _ = format.setIndent(1);

    // Write some data for the validations.
    try writeWorksheetData(&worksheet, format);

    // Set up layout of the worksheet.
    worksheet.setColumnWidth(0, 0, 55);
    worksheet.setColumnWidth(1, 1, 15);
    worksheet.setColumnWidth(3, 3, 15);
    worksheet.setRowHeight(0, 36);

    //
    // Example 1. Limiting input to an integer in a fixed range.
    //
    try worksheet.writeString(
        2,
        0,
        "Enter an integer between 1 and 10",
        null,
    );

    var validator = Valid.Integer.between(1, 10, .{});
    try worksheet.dataValidationCell(2, 1, validator);

    //
    // Example 2. Limiting input to an integer outside a fixed range.
    //
    const text =
        \\Enter an integer that is not between 1 and 10
        \\(using cell references)
    ;

    try worksheet.writeString(
        4,
        0,
        text,
        null,
    );

    validator = Valid.IntegerFormula.not_between(
        "=E3",
        "=F3",
        .{},
    );
    try worksheet.dataValidationCellRef("B5", validator);

    //
    // Example 3. Limiting input to an integer greater than a fixed value.
    //
    try worksheet.writeString(
        6,
        0,
        "Enter an integer greater than 0",
        null,
    );
    // DataValidation.integer.greater_than()
    validator = Valid.Integer.greater_than(0, .{});
    try worksheet.dataValidationCell(6, 1, validator);

    //
    // Example 4. Limiting input to an integer less than a fixed value.
    //
    try worksheet.writeString(
        8,
        0,
        "Enter an integer less than 10",
        null,
    );

    validator = Valid.Integer.lt(10, .{});
    try worksheet.dataValidationCell(8, 1, validator);

    //
    // Example 5. Limiting input to a decimal in a fixed range.
    //
    try worksheet.writeString(
        10,
        0,
        "Enter a decimal between 0.1 and 0.5",
        null,
    );

    validator = Valid.Decimal.between(0.1, 0.5, .{});
    try worksheet.dataValidationCell(10, 1, validator);

    //
    // Example 6. Limiting input to a value in a dropdown list.
    //
    try worksheet.writeString(
        12,
        0,
        "Select a value from a dropdown list",
        null,
    );

    const list_values = [_][]const u8{ "open", "high", "close" };
    validator = Valid.list(&list_values, .{});
    try worksheet.dataValidationCell(12, 1, validator);

    //
    // Example 7. Limiting input to a value in a dropdown list.
    //
    try worksheet.writeString(
        14,
        0,
        "Select a value from a dropdown list (using a cell range)",
        null,
    );

    validator = Valid.listFormula("=$E$4:$G$4", .{});
    try worksheet.dataValidationCellRef("B14", validator);

    //
    // Example 8. Limiting input to a date in a fixed range.
    //
    try worksheet.writeString(
        16,
        0,
        "Enter a date between 1/1/2024 and 12/12/2024",
        null,
    );

    const datetime1 = DateTime{
        .year = 2024,
        .month = 1,
        .day = 1,
        .hour = 0,
        .minute = 0,
        .second = 0,
    };
    const datetime2 = DateTime{
        .year = 2024,
        .month = 12,
        .day = 12,
        .hour = 0,
        .minute = 0,
        .second = 0,
    };

    validator = Valid.Date.between(datetime1, datetime2, .{});
    try worksheet.dataValidationCell(16, 1, validator);

    //
    // Example 9. Limiting input to a time in a fixed range.
    //
    try worksheet.writeString(
        18,
        0,
        "Enter a time between 6:00 and 12:00",
        null,
    );

    const datetime3 = DateTime{
        .year = 0,
        .month = 0,
        .day = 0,
        .hour = 6,
        .minute = 0,
        .second = 0,
    };
    const datetime4 = DateTime{
        .year = 0,
        .month = 0,
        .day = 0,
        .hour = 12,
        .minute = 0,
        .second = 0,
    };

    validator = Valid.Time.between(datetime3, datetime4, .{});
    try worksheet.dataValidationCell(18, 1, validator);

    //
    // Example 10. Limiting input to a string greater than a fixed length.
    //
    try worksheet.writeString(
        20,
        0,
        "Enter a string longer than 3 characters",
        null,
    );

    validator = Valid.Length.greater_than(3, .{});
    try worksheet.dataValidationCell(20, 1, validator);

    //
    // Example 11. Limiting input based on a formula.
    //
    try worksheet.writeString(
        22,
        0,
        "Enter a value if the following is true \"=AND(F5=50,G5=60)\"",
        null,
    );

    validator = Valid.formula("=AND(F5=50,G5=60)", .{});
    try worksheet.dataValidationCell(22, 1, validator);

    //
    // Example 12. Displaying and modifying data validation messages.
    //
    try worksheet.writeString(
        24,
        0,
        "Displays a message when you select the cell",
        null,
    );

    validator = Valid.Integer.between(1, 100, .{
        .input_title = "Enter an integer:",
        .input_message = "between 1 and 100",
        .error_title = "Input value is not valid!",
        .error_message = "It should be an integer between 1 and 100",
        .error_type = .information,
    });
    try worksheet.validationCell(24, 1, validator);

    //
    // Example 13. Displaying and modifying data validation messages.
    //
    try worksheet.writeString(
        26,
        0,
        "Display a custom error message when integer isn't between 1 and 100",
        null,
    );

    validator = Valid.Integer.between(1, 100, .{
        .input_title = "Enter an integer:",
        .input_message = "between 1 and 100",
        .error_title = "Input value is not valid!",
        .error_message = "It should be an integer between 1 and 100",
    });
    try worksheet.dataValidationCell(26, 1, validator);

    //
    // Example 14. Displaying and modifying data validation messages.
    //
    try worksheet.writeString(
        28,
        0,
        "Display a custom info message when integer isn't between 1 and 100",
        null,
    );

    validator = Valid.Integer.between(1, 100, .{
        .input_title = "Enter an integer:",
        .input_message = "between 1 and 100",
        .error_title = "Input value is not valid!",
        .error_message = "It should be an integer between 1 and 100",
        .error_type = .information,
    });
    try worksheet.dataValidationCellRef("B29", validator);

    try workbook.close();
}
