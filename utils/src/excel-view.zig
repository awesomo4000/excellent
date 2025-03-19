const std = @import("std");
const fs = std.fs;
const process = std.process;
const os = std.os;
const Allocator = std.mem.Allocator;
const ArrayList = std.ArrayList;

pub fn main() !void {
    // Get the allocator
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    // Get command line arguments
    var args = try process.argsWithAllocator(allocator);
    defer args.deinit();

    // Get program name for usage message
    const program_name = args.next() orelse "excel-position";

    // Get the two required arguments
    const reference_file = args.next() orelse {
        std.debug.print(
            "Usage: {s} <reference_excel_file> <generated_excel_file>\n",
            .{program_name},
        );
        process.exit(1);
    };

    const generated_file = args.next() orelse {
        std.debug.print(
            "Usage: {s} <reference_excel_file> <generated_excel_file>\n",
            .{program_name},
        );
        process.exit(1);
    };

    // Verify files exist
    const reference_path = try fs.realpathAlloc(
        allocator,
        reference_file,
    );
    defer allocator.free(reference_path);

    const generated_path = try fs.realpathAlloc(
        allocator,
        generated_file,
    );
    defer allocator.free(generated_path);

    // Check if files exist
    _ = try fs.accessAbsolute(reference_path, .{});
    _ = try fs.accessAbsolute(generated_path, .{});

    // Create temporary file
    var temp_dir = try fs.openDirAbsolute("/tmp", .{});
    defer temp_dir.close();

    // Create a unique temporary file name using timestamp
    const timestamp = std.time.timestamp();
    const temp_name = try std.fmt.allocPrint(
        allocator,
        "excel_{d}.xlsx",
        .{timestamp},
    );
    defer allocator.free(temp_name);

    const temp_path = try std.fmt.allocPrint(
        allocator,
        "/tmp/{s}",
        .{temp_name},
    );
    defer allocator.free(temp_path);

    // Create the temporary file
    const temp_file = try fs.createFileAbsolute(
        temp_path,
        .{},
    );
    temp_file.close();

    // Copy generated file to temp location
    try fs.copyFileAbsolute(
        generated_path,
        temp_path,
        .{},
    );

    // Create AppleScript command
    const apple_script = try std.fmt.allocPrint(
        allocator,
        \\set displaySize to do shell script "system_profiler SPDisplaysDataType | grep Resolution | head -1"
        \\set screenWidth to word 2 of displaySize
        \\set screenHeight to word 4 of displaySize
        \\set gap to 0
        \\set halfGap to gap / 2
        \\tell application "Microsoft Excel"
        \\    activate
        \\    set windowWidth to screenWidth / 2
        \\    set windowHeight to screenHeight * 7 / 8
        \\    open POSIX file "{s}"
        \\    set bounds of window 1 to {{0, 0, windowWidth-halfGap, windowHeight}}
        \\    open POSIX file "{s}"
        \\    set bounds of window 1 to {{windowWidth + halfGap, 0, windowWidth * 2 + halfGap, windowHeight}}
        \\end tell
    ,
        .{ reference_path, temp_path },
    );
    defer allocator.free(apple_script);

    // Execute AppleScript
    var child = std.process.Child.init(
        &[_][]const u8{ "osascript", "-e", apple_script },
        allocator,
    );
    child.stdin_behavior = .Ignore;
    child.stdout_behavior = .Ignore;
    child.stderr_behavior = .Ignore;

    try child.spawn();
    const term = try child.wait();

    switch (term) {
        .Exited => |code| {
            if (code != 0) {
                std.debug.print("Error executing AppleScript\n", .{});
                process.exit(1);
            }
        },
        else => {
            std.debug.print("Error executing AppleScript\n", .{});
            process.exit(1);
        },
    }

    // Clean up temporary file
    try fs.deleteFileAbsolute(temp_path);
}
