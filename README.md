# xccdf2xls

Python helping script to aggregate XCCDF result files into an XLS file

## Usage

### Help option

`./xccdf2xls -h` or `./xccdf2xls --help`

### Path option

Use this param to specify where XML files must be found.

**Default:** place where you launch

`./xccdf2xls -p` or `./xccdf2xls --path`

### Group option

Use this param to group rules result by a reference.

**Default:** *null*

`./xccdf2xls -g` or `./xccdf2xls --group`

### Output option

Use this param to specify output file.

**Default:** *result.xlsx*

`./xccdf2xls -o` or `./xccdf2xls --output`
