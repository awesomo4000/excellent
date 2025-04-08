# Enhancement Suggestion: Additional Radar Chart Types

## Current Status
Currently, the high-level API only supports the basic `.radar` chart type, but the C library supports three radar chart types:
1. `LXW_CHART_RADAR` - Basic radar chart
2. `LXW_CHART_RADAR_WITH_MARKERS` - Radar chart with markers
3. `LXW_CHART_RADAR_FILLED` - Filled radar chart

## Proposed Changes
The following changes would enhance the library's chart capabilities:

1. Add new chart types to the `ChartType` enum in `src/chart.zig`:
```zig
pub const ChartType = enum {
    column,
    bar,
    // ... existing types ...
    radar,
    radar_with_markers,  // <-- Add this
    radar_filled,       // <-- Add this
    doughnut,
    // ... other existing types ...
};
```

2. Update the `toNative` function to handle the new chart types:
```zig
fn toNative(self: ChartType) u8 {
    return switch (self) {
        // ... existing mappings ...
        .radar => @intCast(xlsxwriter.LXW_CHART_RADAR),
        .radar_with_markers => @intCast(xlsxwriter.LXW_CHART_RADAR_WITH_MARKERS),
        .radar_filled => @intCast(xlsxwriter.LXW_CHART_RADAR_FILLED),
        .doughnut => @intCast(xlsxwriter.LXW_CHART_DOUGHNUT),
        // ... other existing mappings ...
    };
}
```

3. Then, the `chart_radar.zig` example could be updated to use the specific chart types:
```zig
// Chart 1: Create a simple radar chart
var chart1 = try workbook.addChart(.radar);

// Chart 2: Create a radar chart with markers
var chart2 = try workbook.addChart(.radar_with_markers);

// Chart 3: Create a filled radar chart
var chart3 = try workbook.addChart(.radar_filled);
```

## Benefits
- Provides access to all the radar chart types supported by the underlying C library
- Allows for more specific and accurate representation of radar charts in Excel
- Maintains parity with the C library's capabilities 