# CSS To DOTX
## Discription
This python program converts a Obsidian theme(CSS) to microsoft word theme(DOTX)
Support for all css may come eventually.
**This program only works on windows for now**

## Usage
**Path to css is required**
For basic usage use:
[path to css] is a string (e.g. "./Obsidian gruvbox.css")
```shell
CSSTODOTX.py [path to css]
```
To change output name use (-o, --output):
[name] is a string (e.g. "style"). **This adds the file extension automatically**
```shell
CSSTODOTX.py [path to css] -o [name]
```
Verbosity is (-v, -vv, --verbose):
1. 1-info
2. 2-warning
3. 3-debug(debug + info & warning)
```shell
CSSTODOTX.py [path to css] -v
CSSTODOTX.py [path to css] -vv
CSSTODOTX.py [path to css] --verbose
```
To install font use automatically use (-If, --install-font):
This is a bool (e.g. True). 
**This Requires Administrator mode and will only work on windows, it will download the font if it can find it.**
Also not all themes has a font that needs downloading & installing.
```shell
CSSTODOTX.py [path to css] -If True
```