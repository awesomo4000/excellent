const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    // Add a dependency on the xlsxwriter library
    const xlsxwriter_dep =
        b.dependency("xlsxwriter", .{
            .target = target,
            .optimize = optimize,
        });

    // Add a module for the excellent library
    const lib_mod = b.addModule("excellent", .{
        .root_source_file = b.path("src/excellent.zig"),
        .target = target,
        .optimize = optimize,
    });

    lib_mod.addImport("xlsxwriter", xlsxwriter_dep.module(
        "xlsxwriter",
    ));

    const mktmp_mod = b.addModule("mktmp", .{
        .root_source_file = b.path("src/mktmp.zig"),
        .target = target,
        .optimize = optimize,
    });
    lib_mod.addImport("mktmp", mktmp_mod);

    // Add a clean step that uses std.fs operations
    const clean_step = b.step("clean", "Clean up.");
    clean_step.dependOn(
        &b.addRemoveDirTree(b.path("zig-out")).step,
    );
    clean_step.dependOn(
        &b.addRemoveDirTree(b.path(".zig-cache")).step,
    );

    // Add custom clean step for Excel files
    // in the root directory
    const clean_xlsx_step =
        b.addSystemCommand(&[_][]const u8{
            "sh", "-c", "rm -f *.xlsx *.xlsm",
        });
    clean_step.dependOn(&clean_xlsx_step.step);

    const example_option = b.option(
        []const u8,
        "example",
        "Specify which example to run",
    );

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
        .strip = false,
        .name = "excellent_test",
    });

    // Install the test binary
    const install_test = b.addInstallArtifact(
        lib_unit_tests,
        .{},
    );

    const run_lib_unit_tests = b.addRunArtifact(
        lib_unit_tests,
    );

    const test_step = b.step(
        "test",
        "Run unit tests",
    );
    test_step.dependOn(&run_lib_unit_tests.step);

    // Add coverage step using kcov
    const coverage_dir = b.path("coverage");
    const coverage_cmd = b.addSystemCommand(&[_][]const u8{
        "sh",
        "-c",
        b.fmt(
            \\mkdir -p {s} && \
            \\kcov --clean --include-pattern=src/ {s} {s}
        , .{
            coverage_dir.getPath(b),
            coverage_dir.getPath(b),
            b.pathJoin(
                &.{ b.install_prefix, "bin", "excellent_test" },
            ),
        }),
    });

    coverage_cmd.step.dependOn(&install_test.step);

    const coverage_step = b.step(
        "coverage",
        "Run tests with kcov coverage analysis",
    );
    coverage_step.dependOn(&coverage_cmd.step);

    // A status program that shows the progress of the project
    const status_mod = b.createModule(.{
        .root_source_file = b.path("utils/src/status.zig"),
        .target = target,
        .optimize = optimize,
    });

    const status_exe = b.addExecutable(.{
        .name = "status",
        .root_module = status_mod,
    });

    // Install status to utils
    const status_install = b.addInstallArtifact(
        status_exe,
        .{
            .dest_sub_path = "../../utils/status",
        },
    );
    b.getInstallStep().dependOn(&status_install.step);

    const run_status_cmd = b.addRunArtifact(status_exe);

    run_status_cmd.step.dependOn(b.getInstallStep());

    if (b.args) |args| {
        run_status_cmd.addArgs(args);
    }

    const status_step = b.step(
        "status",
        "Run the status program",
    );
    status_step.dependOn(&run_status_cmd.step);

    // Make status the default run command
    run_step.dependOn(&run_status_cmd.step);

    // Add excel-position utility
    const excel_view_mod = b.createModule(.{
        .root_source_file = b.path("utils/src/excel-view.zig"),
        .target = target,
        .optimize = optimize,
    });

    const excel_view_exe = b.addExecutable(.{
        .name = "excel-view",
        .root_module = excel_view_mod,
    });

    // Install excel-view to utils
    const excel_view_install = b.addInstallArtifact(
        excel_view_exe,
        .{
            .dest_sub_path = "../../utils/excel-view",
        },
    );
    b.getInstallStep().dependOn(&excel_view_install.step);

    // Create a step for building and installing utilities
    const utils_step = b.step(
        "utils",
        "Build and install utility programs",
    );
    utils_step.dependOn(&status_install.step);
    utils_step.dependOn(&excel_view_install.step);

    // Create executables for each example
    const examples_dir = "examples";
    const examples_step = b.step(
        "examples",
        "Build all examples",
    );

    // Make examples step depend on the install step
    examples_step.dependOn(b.getInstallStep());

    // If a specific example is requested, only build that one
    if (example_option) |example| {
        const example_path = b.fmt(
            "{s}/{s}.zig",
            .{ examples_dir, example },
        );
        std.debug.print(
            "Building example: {s}\n",
            .{example_path},
        );
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

        b.installArtifact(example_exe);
        const run_example = b.addRunArtifact(example_exe);
        run_step.dependOn(&run_example.step);
        return;
    }

    // Build all examples if the directory exists
    var has_examples = false;
    if (std.fs.cwd().access(examples_dir, .{})) {
        has_examples = true;
    } else |err| {
        if (err != error.FileNotFound) {
            std.debug.print("[excellent:WARN]", .{});
            std.debug.print(
                "Err checking examples directory: {}\n",
                .{err},
            );
        }
    }

    if (has_examples) {
        if (std.fs.cwd().openDir(
            examples_dir,
            .{ .iterate = true },
        )) |dir| {
            var mutable_dir = dir;
            defer mutable_dir.close();

            var it = mutable_dir.iterate();
            while (it.next() catch unreachable) |entry| {
                if (entry.kind != .file or !std.mem.endsWith(
                    u8,
                    entry.name,
                    ".zig",
                )) continue;

                const example_name =
                    entry.name[0 .. entry.name.len - 4];
                const example_mod_all = b.createModule(.{
                    .root_source_file = b.path(b.fmt(
                        "{s}/{s}",
                        .{ examples_dir, entry.name },
                    )),
                    .target = target,
                    .optimize = optimize,
                });

                example_mod_all.addImport("excellent", lib_mod);

                const example_exe = b.addExecutable(.{
                    .name = example_name,
                    .root_module = example_mod_all,
                });

                // Install the example executable
                const install_example =
                    b.addInstallArtifact(example_exe, .{});
                examples_step.dependOn(&install_example.step);
            }
        } else |err| {
            if (err != error.FileNotFound) {
                std.debug.print("[excellent:WARN]", .{});
                std.debug.print(
                    "Err opening examples directory: {}\n",
                    .{err},
                );
            }
        }
    }
}
