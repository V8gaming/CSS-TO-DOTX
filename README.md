# CSS To DOTX

## Discription

***Requires Python 3.9.2 and above***
This python program converts a Obsidian theme(CSS) to microsoft word theme(DOTX)
Support for all css may come eventually.
**This program only works on windows for now**

## Usage

**Path to css is required**
For basic usage use: <n>
[path to css] is a string (e.g. "./Obsidian gruvbox.css")

```shell
CSSTODOTX.py [path to css]
```

To change output name use (-o, --output): <n>
[name] is a string string without the file extension (e.g. "style"). **This adds the file extension automatically**

```shell
CSSTODOTX.py [path to css] -o [name]
```

Verbosity is (-v, -vv, --verbose): <n>

1. info
2. warning
3. debug(debug + info & warning)
4. dump(all + extra)

```shell
CSSTODOTX.py [path to css] -v
CSSTODOTX.py [path to css] -vv
CSSTODOTX.py [path to css] --verbose
```

To install font automatically use (-If, --install-font): <n> **This Requires Administrator mode and will only work on windows, it will download the font if it can find it.**
Also not all themes has a font that needs downloading & installing.

```shell
CSSTODOTX.py [path to css] --install-font
```

To delete a font from a css use (-Df, --delete-font): <n>
**This Requires Administrator mode and will only work on windows**

```shell
CSSTODOTX.py [path to css] --delete-font
```

To change output directory use (-oD, --output-dir): <n>
[directory] is a string (e.g. "./output"). **This creates the directory if it does not exist**

```bash
CSSTODOTX.py [path to css] -oD [directory]
```

To change output name use (-oN, --output-name): <n>
[name] is a string string without the file extension (e.g. "style"). **This adds the file extension automatically**

```bash
CSSTODOTX.py [path to css] -oN [name]
```

For logging use (-l, --log): <n>
this is outputted in ./logs/[datetime].log

```bash
CSSTODOTX.py [path to css] -l
```

### Examples

Basic usage:

```bash
CSSTODOTX.py "./Obsidian gruvbox.css"
```

Dump verbosity:

```bash
CSSTODOTX.py "./Obsidian gruvbox.css" -vvvv
```

Change output name:

```bash
CSSTODOTX.py "./Obsidian gruvbox.css" -vv -oN "different"
```

Change output directory:

```bash
CSSTODOTX.py "./Obsidian gruvbox.css" -verbose -oD "C:/different"
```

Change output name and directory:

```bash
CSSTODOTX.py "./Obsidian gruvbox.css" -oN "name" -oD "C:/path/to/dir"
```

Change output name and directory, and install font(if available):

```bash
CSSTODOTX.py "./Obsidian gruvbox.css" -vv -oN "new" -oD "./new" -If
```

Change output name and directory, and delete font(if already installed):

```bash
CSSTODOTX.py "./Obsidian gruvbox.css" -vvv -oN "name" -oD "./path" -Df
```

Log output:

```bash
CSSTODOTX.py "./Obsidian gruvbox.css" -l
```

## Modules

[Python 3 Windows font installer](https://gist.github.com/lpsandaruwan/7661e822db3be37e4b50ec9579db61e0)
