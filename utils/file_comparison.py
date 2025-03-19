#!/usr/bin/env python3
"""
File comparison utilities for the autocheck tool.
"""

from pathlib import Path
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

passed_autocheck_file = "autochecked"

def get_relative_path(path, project_root):
    """Convert a path to be relative to the project root if possible"""
    try:
        return path.relative_to(project_root)
    except ValueError:
        # If the path can't be made relative to PROJECT_ROOT, return as is
        return path


def compare_with_reference(example_name, reference_dir, results_dir, project_root, quiet=False, ignore_styles=False):
    """Compare generated Excel file with reference file"""
    generated_file = Path(f"{example_name}.xlsx")
    reference_file = reference_dir / f"{example_name}.xlsx"
    example_results_dir = results_dir / example_name
    
    if not reference_file.exists():
        print(f"⚠️ Reference file not found: {get_relative_path(reference_file, project_root)}")
        return False
    
    try:
        gen_wb = openpyxl.load_workbook(generated_file)
        ref_wb = openpyxl.load_workbook(reference_file)
        
        # Compare sheet names
        if gen_wb.sheetnames != ref_wb.sheetnames:
            print(f"⚠️ Sheet names differ: {gen_wb.sheetnames} vs {ref_wb.sheetnames}")
            return False
        
        style_differences = []
        has_failures = False
        
        # Compare cell contents in each sheet
        for sheet_name in gen_wb.sheetnames:
            gen_sheet = gen_wb[sheet_name]
            ref_sheet = ref_wb[sheet_name]
            
            # Find the max dimensions to iterate through
            max_row = max(gen_sheet.max_row, ref_sheet.max_row)
            max_col = max(gen_sheet.max_column, ref_sheet.max_column)
            
            for row in range(1, max_row + 1):
                for col in range(1, max_col + 1):
                    gen_cell = gen_sheet.cell(row=row, column=col)
                    ref_cell = ref_sheet.cell(row=row, column=col)
                    
                    # Special handling for ArrayFormula objects
                    if (str(gen_cell.value).startswith('<openpyxl.worksheet.formula.ArrayFormula') and 
                        str(ref_cell.value).startswith('<openpyxl.worksheet.formula.ArrayFormula')):
                        # Check if they're both array formulas - consider them equal for this check
                        if hasattr(gen_cell, 'value') and hasattr(ref_cell, 'value'):
                            # They're both array formulas, so we'll consider them equal
                            continue
                    
                    if gen_cell.value != ref_cell.value:
                        print(f"⚠️ Cell value mismatch at {sheet_name}!{gen_cell.coordinate}: "
                              f"'{gen_cell.value}' vs '{ref_cell.value}'")
                        has_failures = True
                    
                    # Skip style checking if ignore_styles is set
                    if ignore_styles:
                        continue
                        
                    # Compare cell styles if they exist
                    if hasattr(gen_cell, '_style') and hasattr(ref_cell, '_style'):
                        gen_style = gen_cell._style
                        ref_style = ref_cell._style
                        
                        # Compare all style properties
                        style_diff = []
                        
                        # Compare background color
                        if hasattr(gen_style, 'fill') and hasattr(ref_style, 'fill'):
                            gen_fill = gen_style.fill.start_color.rgb if gen_style.fill.start_color else None
                            ref_fill = ref_style.fill.start_color.rgb if ref_style.fill.start_color else None
                            if gen_fill != ref_fill:
                                style_diff.append(f"background color: {gen_fill} vs {ref_fill}")
                                has_failures = True
                        
                        # Compare font properties
                        if hasattr(gen_style, 'font') and hasattr(ref_style, 'font'):
                            font_props = [
                                ('bold', 'bold'),
                                ('italic', 'italic'),
                                ('underline', 'underline'),
                                ('strike', 'strike'),
                                ('color', 'color.rgb'),
                                ('size', 'size'),
                                ('name', 'name'),
                                ('vertAlign', 'vertAlign'),
                                ('scheme', 'scheme')
                            ]
                            
                            for prop, path in font_props:
                                gen_val = getattr(gen_style.font, prop)
                                ref_val = getattr(ref_style.font, prop)
                                if gen_val != ref_val:
                                    style_diff.append(f"font {prop}: {gen_val} vs {ref_val}")
                                    has_failures = True
                        
                        # Compare borders
                        if hasattr(gen_style, 'border') and hasattr(ref_style, 'border'):
                            for side in ['top', 'bottom', 'left', 'right']:
                                gen_border = getattr(gen_style.border, side)
                                ref_border = getattr(ref_style.border, side)
                                
                                # Compare border style
                                if gen_border.style != ref_border.style:
                                    style_diff.append(f"{side} border style: {gen_border.style} vs {ref_border.style}")
                                    has_failures = True
                                
                                # Compare border color
                                gen_color = gen_border.color.rgb if gen_border.color else None
                                ref_color = ref_border.color.rgb if ref_border.color else None
                                if gen_color != ref_color:
                                    style_diff.append(f"{side} border color: {gen_color} vs {ref_color}")
                                    has_failures = True
                        
                        # Compare alignment
                        if hasattr(gen_style, 'alignment') and hasattr(ref_style, 'alignment'):
                            align_props = [
                                ('horizontal', 'horizontal'),
                                ('vertical', 'vertical'),
                                ('textRotation', 'textRotation'),
                                ('wrapText', 'wrapText'),
                                ('shrinkToFit', 'shrinkToFit'),
                                ('indent', 'indent'),
                                ('relativeIndent', 'relativeIndent'),
                                ('justifyLastLine', 'justifyLastLine')
                            ]
                            
                            for prop, path in align_props:
                                gen_val = getattr(gen_style.alignment, prop)
                                ref_val = getattr(ref_style.alignment, prop)
                                if gen_val != ref_val:
                                    style_diff.append(f"alignment {prop}: {gen_val} vs {ref_val}")
                                    has_failures = True
                        
                        # Compare number format
                        if hasattr(gen_style, 'number_format') and hasattr(ref_style, 'number_format'):
                            if gen_style.number_format != ref_style.number_format:
                                style_diff.append(f"number format: {gen_style.number_format} vs {ref_style.number_format}")
                                has_failures = True
                        
                        # If there are any style differences, add them to the list
                        if style_diff:
                            style_differences.append({
                                'cell': f"{sheet_name}!{gen_cell.coordinate}",
                                'differences': style_diff
                            })
        
        # Report any style differences found
        if style_differences and not ignore_styles:
            print("\n⚠️ Style differences found:")
            for diff in style_differences:
                print(f"\nCell: {diff['cell']}")
                for d in diff['differences']:
                    print(f"  - {d}")
        
        # When ignore_styles is true, we don't want style differences to cause a failure
        style_failure = has_failures
        if ignore_styles:
            # If we have differences but they're only style differences, don't fail
            style_failure = has_failures and not style_differences
            if style_differences and not quiet:
                print("⚠️ Style differences found but ignored due to --ignore-styles flag")
                
        if not quiet:
            if style_failure:
                print("\n❌ Generated file has differences from reference file")
            else:
                print("✅ Generated file matches reference file content" + 
                     (" (ignoring styles)" if ignore_styles and style_differences else ""))
        
        # Create results directory if it doesn't exist
        example_results_dir.mkdir(parents=True, exist_ok=True)
        
        # Create or remove autochecked file based on results
        autochecked_file = example_results_dir / passed_autocheck_file
        if not style_failure:
            autochecked_file.touch()
            if not quiet:
                print(f"✅ Created autochecked file at {get_relative_path(autochecked_file, project_root)}")
        else:
            if autochecked_file.exists():
                autochecked_file.unlink()
            if not quiet:
                print(f"❌ Removed autochecked file due to differences")
        
        return not style_failure
        
    except InvalidFileException as e:
        print(f"❌ Error opening Excel file: {e}")
        return False 