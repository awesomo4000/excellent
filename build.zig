const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    // Add a clean step that uses std.fs operations
    const clean_step = b.step("clean", "Clean up.");
    clean_step.dependOn(&b.addRemoveDirTree(b.path("zig-out")).step);
    clean_step.dependOn(&b.addRemoveDirTree(b.path(".zig-cache")).step);

    const example_option = b.option(
        []const u8,
        "example",
        "Specify which example to run",
    );

    const xlsxwriter_dep = b.dependency("xlsxwriter", .{
        .target = target,
        .optimize = optimize,
    });

    const lib_mod = b.createModule(.{
        .root_source_file = b.path("src/excellent.zig"),
        .target = target,
        .optimize = optimize,
    });

    lib_mod.addImport("xlsxwriter", xlsxwriter_dep.module("xlsxwriter"));

    const exe_mod = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });

    exe_mod.addImport("excellent", lib_mod);

    const lib = b.addLibrary(.{
        .linkage = .static,
        .name = "excellent",
        .root_module = lib_mod,
    });

    b.installArtifact(lib);

    const run_step = b.step(
        "run",
        "Run example (-Dexample=example_name)",
    );

    const lib_unit_tests = b.addTest(.{
        .root_module = lib_mod,
    });

    const run_lib_unit_tests = b.addRunArtifact(lib_unit_tests);

    const test_step = b.step("test", "Run unit tests");
    test_step.dependOn(&run_lib_unit_tests.step);

    // A status program that shows the progress of the project
    const status_mod = b.createModule(.{
        .root_source_file = b.path("testing/status.zig"),
        .target = target,
        .optimize = optimize,
    });

    const status_exe = b.addExecutable(.{
        .name = "status",
        .root_module = status_mod,
    });

    b.installArtifact(status_exe);

    const run_status_cmd = b.addRunArtifact(status_exe);

    run_status_cmd.step.dependOn(b.getInstallStep());

    if (b.args) |args| {
        run_status_cmd.addArgs(args);
    }

    const status_step = b.step("status", "Run the status program");
    status_step.dependOn(&run_status_cmd.step);

    // Make status the default run command
    run_step.dependOn(&run_status_cmd.step);

    // Create executables for each example
    const examples_dir = "examples";
    const examples_step = b.step("examples", "Build all examples");

    // If a specific example is requested, only build that one
    if (example_option) |example| {
        const example_path = b.fmt("{s}/{s}.zig", .{ examples_dir, example });
        std.debug.print("Building example: {s}\n", .{example_path});
        const example_mod = b.createModule(.{
            .root_source_file = b.path(example_path),
            .target = target,
            .optimize = optimize,
        });

        example_mod.addImport("excellent", lib_mod);

        const example_exe = b.addExecutable(.{
            .name = example,
            .root_module = example_mod,
        });

        const run_example = b.addRunArtifact(example_exe);
        b.installArtifact(example_exe);
        examples_step.dependOn(&example_exe.step);
        run_step.dependOn(&run_example.step);
    } else {
        // Otherwise build all examples
        var dir = std.fs.cwd().openDir(examples_dir, .{ .iterate = true }) catch unreachable;
        defer dir.close();

        var it = dir.iterate();
        while (it.next() catch unreachable) |entry| {
            if (entry.kind != .file or !std.mem.endsWith(u8, entry.name, ".zig")) continue;

            const example_name = entry.name[0 .. entry.name.len - 4];
            const example_mod = b.createModule(.{
                .root_source_file = b.path(b.fmt("{s}/{s}", .{ examples_dir, entry.name })),
                .target = target,
                .optimize = optimize,
            });

            example_mod.addImport("excellent", lib_mod);

            const example_exe = b.addExecutable(.{
                .name = example_name,
                .root_module = example_mod,
            });

            b.installArtifact(example_exe);
            examples_step.dependOn(&example_exe.step);
            examples_step.dependOn(b.getInstallStep());
        }
    }
}
