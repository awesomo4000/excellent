const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});
    const lib_mod = b.createModule(.{
        .root_source_file = b.path("src/root.zig"),
        .target = target,
        .optimize = optimize,
    });
    const exe_mod = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });

    exe_mod.addImport("excellent_lib", lib_mod);

    const lib = b.addLibrary(.{
        .linkage = .static,
        .name = "excellent",
        .root_module = lib_mod,
    });

    b.installArtifact(lib);

    const exe = b.addExecutable(.{
        .name = "excellent",
        .root_module = exe_mod,
    });

    b.installArtifact(exe);

    const run_cmd = b.addRunArtifact(exe);

    run_cmd.step.dependOn(b.getInstallStep());

    if (b.args) |args| {
        run_cmd.addArgs(args);
    }

    const run_step = b.step("run", "Run the app");

    run_step.dependOn(&run_cmd.step);

    const lib_unit_tests = b.addTest(.{
        .root_module = lib_mod,
    });

    const run_lib_unit_tests = b.addRunArtifact(lib_unit_tests);

    const exe_unit_tests = b.addTest(.{
        .root_module = exe_mod,
    });

    const run_exe_unit_tests = b.addRunArtifact(exe_unit_tests);

    const test_step = b.step("test", "Run unit tests");

    test_step.dependOn(&run_lib_unit_tests.step);

    test_step.dependOn(&run_exe_unit_tests.step);

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
}
