# Tutorials
A collection of useful tutorial


# Files

1. [vbscript.md](https://github.com/filovirid/tutorials/blob/main/vbscript.md) Visual Basic scripting edition tutorial.
2. [Script56.CHM](https://github.com/filovirid/tutorials/blob/main/Script56.CHM) Official Microsoft tutorial for VBS




### Generate HTML from MD file:
```bash
python3 -c 'import markdown2 as m;h=m.markdown_path("vbscript.md",extras={"fenced-code-blocks":None,"tables":None,"html-classes":{"table":"table border table-striped"}});print(h)'> output.html
```
