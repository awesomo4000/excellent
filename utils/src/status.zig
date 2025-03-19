// A status program that shows the progress of the project

// print to stdout all the time instead of log or debug

const std = @import("std");
const fs = std.fs;
const stdout = std.io.getStdOut();
const stdoutWriter = stdout.writer();

const passed_autocheck_file = "autochecked";
const Example = struct {
    name: []const u8,
    implemented: bool,
    verified: bool,
    autocheck_passed: bool,
    have_ref_xlsx: bool,
};

const OutputFormat = enum {
    normal,
    short,
};

fn printWrappedNames(
    writer: anytype,
    names: []const []const u8,
    prefix: []const u8,
) !void {
    const labels = [_][]const u8{
        "haveRefXlsx: ",
        "haveZig: ",
        "autoChecked: ",
        "verified: ",
    };

    // Find the longest label and colon position
    var max_label_len: usize = 0;
    var max_colon_pos: usize = 0;
    for (labels) |label| {
        max_label_len = @max(max_label_len, label.len);
        const colon_pos = std.mem.indexOfScalar(u8, label, ':') orelse label.len;
        max_colon_pos = @max(max_colon_pos, colon_pos);
    }

    // Calculate prefix padding to right-align the colon
    const colon_pos = std.mem.indexOfScalar(u8, prefix, ':') orelse prefix.len;
    const prefix_padding = max_colon_pos - colon_pos;

    // Calculate total prefix length with padding
    const padded_prefix_len = prefix.len + prefix_padding;

    // Print prefix padding first
    for (0..prefix_padding) |_| {
        try writer.print(" ", .{});
    }

    // Print the prefix
    try writer.print("{s}", .{prefix});

    var line_len: usize = padded_prefix_len;

    for (names, 0..) |name, i| {
        if (line_len + name.len > 70) {
            try writer.print("\n", .{});
            for (0..max_label_len) |_| {
                try writer.print(" ", .{});
            }
            line_len = max_label_len;
        }
        try writer.print("{s}", .{name});
        if (i < names.len - 1) {
            try writer.print(" ", .{});
            line_len += 1;
        }
        line_len += name.len;
    }
    try writer.print("\n", .{});
}

fn setupExamples(
    allocator: std.mem.Allocator,
) !std.ArrayList(Example) {
    const examples_dir = "testing/zig-c-binding-examples";
    const examples_out_dir = "examples";

    var examples = std.ArrayList(Example).init(allocator);

    // Get all example files
    var dir = try fs.cwd().openDir(examples_dir, .{ .iterate = true });
    defer dir.close();

    var dir_iterator = dir.iterate();
    while (try dir_iterator.next()) |entry| {
        if (entry.kind == .file and std.mem.endsWith(u8, entry.name, ".zig")) {
            const name =
                try allocator.dupe(u8, entry.name[0 .. entry.name.len - 4]);
            const implemented = blk: {
                const out_path = try std.fmt.allocPrint(
                    allocator,
                    "{s}/{s}.zig",
                    .{ examples_out_dir, name },
                );
                defer allocator.free(out_path);
                break :blk fs.cwd().access(out_path, .{}) != error.FileNotFound;
            };
            const verified = blk: {
                const result_path = try std.fmt.allocPrint(
                    allocator,
                    "testing/results/{s}/verified",
                    .{name},
                );
                defer allocator.free(result_path);
                break :blk fs.cwd().access(result_path, .{}) != error.FileNotFound;
            };
            const autocheck_passed = blk: {
                const autocheck_path = try std.fmt.allocPrint(
                    allocator,
                    "testing/results/{s}/{s}",
                    .{ name, passed_autocheck_file },
                );
                defer allocator.free(autocheck_path);
                break :blk fs.cwd().access(autocheck_path, .{}) != error.FileNotFound;
            };
            const have_ref_xlsx = blk: {
                const ref_path = try std.fmt.allocPrint(
                    allocator,
                    "testing/reference-xls/{s}.xlsx",
                    .{name},
                );
                defer allocator.free(ref_path);
                break :blk fs.cwd().access(ref_path, .{}) != error.FileNotFound;
            };
            try examples.append(.{
                .name = name,
                .implemented = implemented,
                .verified = verified,
                .autocheck_passed = autocheck_passed,
                .have_ref_xlsx = have_ref_xlsx,
            });
        }
    }

    // Sort examples alphabetically
    const lessThan = struct {
        fn lessThan(_: void, a: Example, b: Example) bool {
            return std.mem.lessThan(u8, a.name, b.name);
        }
    }.lessThan;
    std.mem.sort(Example, examples.items, {}, lessThan);

    return examples;
}

fn findMaxNameLength(
    examples: std.ArrayList(Example),
) usize {
    var max_len: usize = 0;
    for (examples.items) |example| {
        max_len = @max(max_len, example.name.len);
    }
    return max_len;
}

fn printShortOutput(
    writer: anytype,
    examples: std.ArrayList(Example),
    allocator: std.mem.Allocator,
) !void {
    var impl_names = std.ArrayList([]const u8).init(allocator);
    defer impl_names.deinit();
    var verified_names = std.ArrayList([]const u8).init(allocator);
    defer verified_names.deinit();
    var autocheck_passed_names = std.ArrayList([]const u8).init(allocator);
    defer autocheck_passed_names.deinit();
    var have_ref_xlsx_names = std.ArrayList([]const u8).init(allocator);
    defer have_ref_xlsx_names.deinit();

    for (examples.items) |example| {
        if (example.implemented) {
            try impl_names.append(example.name);
            if (example.verified) {
                try verified_names.append(example.name);
            }
            if (example.autocheck_passed) {
                try autocheck_passed_names.append(example.name);
            }
        }
        if (example.have_ref_xlsx) {
            try have_ref_xlsx_names.append(example.name);
        }
    }

    try printWrappedNames(
        writer,
        have_ref_xlsx_names.items,
        "haveRefXlsx: ",
    );
    try writer.print(
        "({d}/{d})\n\n",
        .{ have_ref_xlsx_names.items.len, examples.items.len },
    );
    try printWrappedNames(
        writer,
        impl_names.items,
        "haveZig: ",
    );
    try writer.print(
        "({d}/{d})\n\n",
        .{ impl_names.items.len, examples.items.len },
    );
    try printWrappedNames(
        writer,
        autocheck_passed_names.items,
        "autoChecked: ",
    );
    try writer.print(
        "({d}/{d})\n\n",
        .{ autocheck_passed_names.items.len, examples.items.len },
    );
    try printWrappedNames(
        writer,
        verified_names.items,
        "verified: ",
    );
    try writer.print(
        "({d}/{d})\n",
        .{ verified_names.items.len, examples.items.len },
    );
}

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    var args = try std.process.argsWithAllocator(allocator);
    defer args.deinit();
    _ = args.skip(); // skip program name
    const example_arg = args.next();

    var examples = try setupExamples(allocator);
    defer {
        for (examples.items) |example| {
            allocator.free(example.name);
        }
        examples.deinit();
    }

    // If specific example requested, show only that
    if (example_arg) |arg| {
        if (std.mem.eql(u8, arg, "--short")) {
            try printShortOutput(stdoutWriter, examples, allocator);
            return;
        }

        const name = arg[0 .. std.mem.indexOf(
            u8,
            arg,
            ".zig",
        ) orelse arg.len];
        for (examples.items) |example| {
            if (std.mem.eql(u8, example.name, name)) {
                try stdoutWriter.print(
                    "{s: >30}  {s}  {s}  {s}  {s}\n",
                    .{
                        example.name,
                        if (example.implemented) "[.zig]" else "",
                        if (example.autocheck_passed) "[autochecked]" else "",
                        if (example.verified) "[verified]" else "",
                        if (example.have_ref_xlsx) "[ref]" else "",
                    },
                );
                return;
            }
        }
        try stdoutWriter.print(
            "err: Example '{s}' not found\n",
            .{name},
        );
        try stdoutWriter.print(
            "Usage: status [--short|example_name]\n",
            .{},
        );
        return;
    }

    // Show all examples
    var implemented: usize = 0;
    var verified: usize = 0;
    var autocheck_passed: usize = 0;
    var have_ref_xlsx: usize = 0;
    for (examples.items) |example| {
        if (example.implemented) implemented += 1;
        if (example.verified) verified += 1;
        if (example.autocheck_passed) autocheck_passed += 1;
        if (example.have_ref_xlsx) have_ref_xlsx += 1;
        try stdoutWriter.print(
            "{s: >30}  {s}  {s}  {s}  {s}\n",
            .{
                example.name,
                if (example.implemented) "zig" else "",
                if (example.verified) "verf" else "",
                if (example.autocheck_passed) "autochecked" else "",
                if (example.have_ref_xlsx) "ref" else "",
            },
        );
    }
    try stdoutWriter.print(
        "\nProgress: {d}/{d} examples started ({d:.1}%), {d}/{d} autochecked ({d:.1}%), {d}/{d} verified ({d:.1}%), {d}/{d} have ref ({d:.1}%)\n",
        .{
            implemented,
            examples.items.len,
            @as(f64, @floatFromInt(implemented)) /
                @as(f64, @floatFromInt(examples.items.len)) * 100.0,
            autocheck_passed,
            examples.items.len,
            @as(f64, @floatFromInt(autocheck_passed)) /
                @as(f64, @floatFromInt(examples.items.len)) * 100.0,
            verified,
            examples.items.len,
            @as(f64, @floatFromInt(verified)) /
                @as(f64, @floatFromInt(examples.items.len)) * 100.0,
            have_ref_xlsx,
            examples.items.len,
            @as(f64, @floatFromInt(have_ref_xlsx)) /
                @as(f64, @floatFromInt(examples.items.len)) * 100.0,
        },
    );
}
