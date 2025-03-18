// A status program that shows the progress of the project

// print to stdout all the time instead of log or debug

const std = @import("std");
const fs = std.fs;
const stdout = std.io.getStdOut();
const stdoutWriter = stdout.writer();

const Example = struct {
    name: []const u8,
    implemented: bool,
    verified: bool,
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
    var line_len: usize = prefix.len;
    try writer.print("{s}", .{prefix});

    for (names, 0..) |name, i| {
        if (line_len + name.len > 60) {
            try writer.print("\n", .{});
            for (0..prefix.len) |_| {
                try writer.print(" ", .{});
            }
            line_len = prefix.len;
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
            try examples.append(.{
                .name = name,
                .implemented = implemented,
                .verified = verified,
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
            var unstarted_names = std.ArrayList([]const u8).init(allocator);
            defer unstarted_names.deinit();
            var impl_names = std.ArrayList([]const u8).init(allocator);
            defer impl_names.deinit();
            var verified_names = std.ArrayList([]const u8).init(allocator);
            defer verified_names.deinit();

            for (examples.items) |example| {
                if (!example.implemented) {
                    try unstarted_names.append(example.name);
                } else {
                    try impl_names.append(example.name);
                    if (example.verified) {
                        try verified_names.append(example.name);
                    }
                }
            }

            try printWrappedNames(
                stdoutWriter,
                unstarted_names.items,
                "Unstarted: ",
            );
            try stdoutWriter.print(
                "({d}/{d})\n\n",
                .{ unstarted_names.items.len, examples.items.len },
            );
            try printWrappedNames(
                stdoutWriter,
                impl_names.items,
                "Implemented: ",
            );
            try stdoutWriter.print(
                "({d}/{d})\n\n",
                .{ impl_names.items.len, examples.items.len },
            );
            try printWrappedNames(
                stdoutWriter,
                verified_names.items,
                "Verified: ",
            );
            try stdoutWriter.print(
                "({d}/{d})\n",
                .{ verified_names.items.len, examples.items.len },
            );
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
                    "{s: >30}  impl={s}  verified={s}\n",
                    .{
                        example.name,
                        if (example.implemented) "✓" else "✗",
                        if (example.verified) "✓" else "✗",
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
    for (examples.items) |example| {
        if (example.implemented) implemented += 1;
        if (example.verified) verified += 1;
        try stdoutWriter.print(
            "{s: >30}  impl={s}  verified={s}\n",
            .{
                example.name,
                if (example.implemented) "✓" else "✗",
                if (example.verified) "✓" else "✗",
            },
        );
    }
    try stdoutWriter.print(
        "\nProgress: {d}/{d} examples implemented ({d:.1}%), {d}/{d} verified ({d:.1}%)\n",
        .{
            implemented,
            examples.items.len,
            @as(f64, @floatFromInt(implemented)) /
                @as(f64, @floatFromInt(examples.items.len)) * 100.0,
            verified,
            examples.items.len,
            @as(f64, @floatFromInt(verified)) /
                @as(f64, @floatFromInt(examples.items.len)) * 100.0,
        },
    );
}
