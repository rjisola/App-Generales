# AI Coding Instructions for Payroll Processing System

## Overview
This is a Python-based payroll processing system for construction workers (UOCRA/NASA/UECARA unions). It processes employee hours, calculates salaries, generates PDF receipts, and integrates heavily with Excel files containing VBA macros.

## Architecture
- **Launcher**: `1-LAUNCHER.pyw` - GUI launcher for tools
- **Core Processor**: `B-PROCESARSUELDOS.pyw` - Main payroll processing script
- **Key Modules**:
  - `data_loader.py` - Loads Excel data, maps cell colors to employee categories
  - `logic_payroll.py` - Calculates hours, rates, bonuses per employee category
  - `logic_accountant.py` - Generates accountant summaries
  - `excel_format_writer.py` - Writes results back to Excel preserving format
  - `receipt_font_formatter.py` - Applies fonts to receipt PDFs

## Data Flow
1. Read employee data from Excel `.xlsm` files (sheet 'CALCULAR HORAS')
2. Process each employee row using category-specific logic (AMARILLO/AZUL/CELESTE/etc.)
3. Calculate hours based on day inputs and holiday definitions
4. Apply category rules from `config.json` (hourly rates, bonuses, divisors)
5. Write results to modified Excel file and generate PDF receipts

## Key Conventions
- **Employee Categories**: Determined by Excel cell background colors (RGB mappings in `data_loader.py`)
  - AMARILLO: Sub-projects via cell colors (QUILMES=NARANJA, PAPELERA=VERDE, NORMAL=BLANCO)
  - AZUL/CELESTE: Base salary divisors (110/120 hours)
  - SALMON: Bonus factors (1.2x) with category-specific hourly rates
- **File Extensions**: `.pyw` for GUI scripts, `.xlsm` for Excel with VBA
- **Sheet Names**: Configured in `config.json` (CALCULAR HORAS, ENVIO CONTADOR, etc.)
- **Day Processing**: Multiple occurrences of same weekday handled with counters
- **Holiday Detection**: Row 7 markers in Excel for holiday columns

## Developer Workflows
- **Setup**: Run `instalar_dependencias.bat` to install from `requirements.txt`
- **Clean Run**: `EJECUTAR_LIMPIO.bat` clears `__pycache__` and runs receipt generator
- **Main Processing**: Execute `B-PROCESARSUELDOS.pyw` or use launcher
- **Console Mode**: Add `--console` flag for headless processing
- **Font Customization**: Use `--font Arial` for receipt formatting

## Common Patterns
- **Color Reading**: Use `openpyxl` with `data_only=False` for styles, `data_only=True` for values
- **Employee Iteration**: Loop DataFrame rows, skip empty names, preserve Excel row indices
- **Error Handling**: Continue processing on individual employee failures
- **VBA Preservation**: Always use `keep_vba=True` when loading workbooks
- **Path Handling**: Generate output files in same directory as input Excel

## Integration Points
- **Excel VBA**: Macros in `Macros Excel/` perform initial calculations
- **PDF Generation**: Uses `reportlab` for receipts, `PyPDF2`/`pikepdf` for manipulation
- **Font Application**: Post-process PDFs with `receipt_font_formatter.py`
- **Icon Loading**: PNG icons in `launcher_icons/` for GUI

## Examples
- **Category Logic**: See `logic_payroll.py` `_process_amarillo_with_colors()` for color-based sub-project detection
- **Config Usage**: Hourly rates referenced as Excel cells (e.g., "B3") in `config.json`
- **Output Verification**: `verify_output_file()` checks generated Excel integrity