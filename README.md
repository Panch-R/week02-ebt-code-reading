# Week 2 - EBT Python Code Reading

## Goal
Understand the existing EBT Python code and its design.

## What I learned

### 1. Path and file handling
- `Path`
- `glob()`
- `list()`
- `for`
- `file.name`
- `file.stem`
- `file.suffix`

### 2. Excel / COM objects
- Excel Application
- Workbook
- Worksheet
- Range
- `.Value`
- Source / Destination

### 3. Python functions
- `def`
- arguments
- `return`
- `lambda`

### 4. Error handling
- `try`
- `except`
- `continue`
- `raise`
- `retry()`

### 5. EBT code design
The source workbook is opened as read-only because it is only used as
an input source. Data is read from the source worksheet and written
to the destination workbook.

## Key takeaway

I can now explain the basic flow of the existing EBT code:

Source files
→ find the correct file
→ open the source workbook
→ select the source worksheet
→ read the data range
→ write the data to the destination worksheet
→ save the destination workbook.

## What I can explain now
- The difference between a Workbook and a Worksheet
- How to designate only one source file
- Why Source files are opened as ReadOnly
- How Excel data is transferred through Range.Value
- Why list() is used with Path.glob()
- How for / if / continue work together
- How lambda is used with retry()
- Why retry() is necessary when Excel is busy
