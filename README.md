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

```shell
CSSTODOTX.py [path to css] -v
CSSTODOTX.py [path to css] -vv
CSSTODOTX.py [path to css] --verbose
```

To install font automatically use (-If, --install-font): <n>
This is a bool (e.g. True).  **This Requires Administrator mode and will only work on windows, it will download the font if it can find it.**
Also not all themes has a font that needs downloading & installing.

```shell
CSSTODOTX.py [path to css] --install-font True
```

To delete a font from a css use (-Df, --delete-font): <n>
This is a bool (e.g. True).  **This Requires Administrator mode and will only work on windows**

```shell
CSSTODOTX.py [path to css] --delete-font True
```

## Modules

[Python 3 Windows font installer](https://gist.github.com/lpsandaruwan/7661e822db3be37e4b50ec9579db61e0)