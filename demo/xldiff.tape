# VHS tape for recording xldiff demo GIF.
# Usage: vhs demo/xldiff.tape
#
# Prerequisites:
#   - xldiff binary built and in PATH
#   - Demo xlsx files created: cargo run --example create_demo

Output demo/xldiff.gif

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

# ─── 1. SHOW THE INPUT FILES ───
Type "xlcat demo/old.xlsx"
Enter
Sleep 3s

Type "xlcat demo/new.xlsx"
Enter
Sleep 3s

# ─── 2. KEY-BASED DIFF ───
Type "xldiff demo/old.xlsx demo/new.xlsx --key ID"
Enter
Sleep 4s

# ─── 3. MARKDOWN OUTPUT ───
Type "xldiff demo/old.xlsx demo/new.xlsx --key ID --format markdown"
Enter
Sleep 4s

# ─── 4. JSON OUTPUT ───
Type "xldiff demo/old.xlsx demo/new.xlsx --key ID --format json"
Enter
Sleep 4s
