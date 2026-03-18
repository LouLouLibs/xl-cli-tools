# VHS tape for recording xlfilter demo GIF.
# Usage: vhs demo/xlfilter.tape
#
# Prerequisites:
#   - xlfilter binary built and in PATH
#   - Demo xlsx files created: cargo run --example create_demo

Output demo/xlfilter.gif

Set FontSize 14
Set Width 1100
Set Height 550
Set Padding 15
Set Theme "GruvboxDarkHard"

Set TypingSpeed 80ms

Set Shell "bash"

Hide
    Type 'export PS1="> "'
    Enter
    Type "clear"
    Enter
Show

# --- 1. FILTER ROWS ---
Type "xlfilter demo/sales.xlsx --where Region=East"
Enter
Sleep 3s

# --- 2. NUMERIC COMPARISON ---
Type "xlfilter demo/sales.xlsx --where 'Revenue>12000'"
Enter
Sleep 3s

# --- 3. MULTIPLE FILTERS (AND) ---
Type "xlfilter demo/sales.xlsx --where Region=East --where 'Revenue>10000'"
Enter
Sleep 3s

# --- 4. SELECT COLUMNS + SORT ---
Type "xlfilter demo/sales.xlsx --cols Region,Product,Revenue --sort Revenue:desc"
Enter
Sleep 3s

# --- 5. LIMIT OUTPUT ---
Type "xlfilter demo/sales.xlsx --sort Revenue:desc --limit 3 --cols Product,Revenue"
Enter
Sleep 2s

# --- 6. CONTAINS FILTER ---
Type "xlfilter demo/sales.xlsx --where 'Product~Widget A' --cols Date,Region,Revenue"
Enter
Sleep 3s

# --- 7. CSV OUTPUT FOR PIPING ---
Type "xlfilter demo/sales.xlsx --where Region=West --csv"
Enter
Sleep 3s
