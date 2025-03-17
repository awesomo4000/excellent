# Welcome, new-hire.

> **Note**: This document is a guide for the AI assistant to understand the project structure and workflow. When starting fresh, this document will help the assistant provide consistent and accurate assistance.

You are writing Zig code. The code produced will provide a high-level, user-friendly ergonomic, and idiomatic API for the production of Excel spreadsheet (.xlsx files).
The high-level API is a wrapper around a lower-level zig binding to a C library, libxlsxwriter. The example programs for this binding will be referred to for conversion into the high-level API.

## Overview
Your task is to create a high-level API that makes Excel file generation simple and intuitive in Zig. You'll do this by:
1. Studying the low-level bindings in `testing/zig-c-binding-examples/`
2. Creating corresponding high-level examples in `examples/`
3. Verifying your work against reference files

You will be running programs from the project root which is the current directory. 

Directories in the project are:

**src/**  : High-level interface wrapper code

**examples/** : Being populated with examples using the new highlevel wrappers being developed. Each example from testing/zig-c-binding-examples/ should be represented by a corresponding example in this directory.

**testing/**  : Files and programs useful for testing and verification

**testing/reference-xls/*.xlsx** : Reference xlsx files that are to be compared to the output generated from the high-level wrappers being developed

**testing/test-output-xls** : Output directory to place verified outputs from examples/*.zig after testing them to make sure they produce correct output

**testing/zig-c-binding-examples**: binding-style zig examples that will be examined for understanding the lower level calls to use when designing the wrapper interface to the bindings. These will not be modified.

**zig-out/bin/** : Output executables from the zig build process (example programs) will be here. When executed, they produce a spreadsheet. This will be compared against a reference spreadsheet in **testing/reference-xls/*.xlsx** .

**testing/status.zig** : A program that will show the current status of progress on creating the examples corresponding to the examples in zig-c-binding-examples.

## Workflow

### Development Cycle
Follow this cycle:

1. Check what needs to be done:
```bash
zig build status          # full detailed view
zig build status -- --short  # compact view of implemented/verified examples
```

2. Check status of a specific example:
```bash
zig build status -- hello
```

3. Verify your work:
```bash
zig build run verify -- hello
```

## Coding Standards

- Use comma (,) after the last parameter in function definitions, struct
  definitions, function calls, etc. so the zig formatter will wrap lines
  and keep the line width less than 80 characters

- Refactor common functionality so functions are short