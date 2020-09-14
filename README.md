# xccdf2xls

Python helping script to aggregate XCCDF result files into an XLS file

## Usage

### Help option

`python3 xccdf2xls.py -h` or `python3 xccdf2xls.py --help`

### Path option

Use this param to specify where XML files must be found.

**Default:** place where you launch

`python3 xccdf2xls.py -p` or `python3 xccdf2xls.py --path`

### Group option

Use this param to group rules result by a reference.

**Default:** _null_

`python3 xccdf2xls.py -g` or `python3 xccdf2xls.py --group`

### Output option

Use this param to specify output file.

**Default:** _result.xlsx_

`python3 xccdf2xls.py -o` or `python3 xccdf2xls.py --output`

## Examples

`python3 xccdf2xls.py -p "Results_Sample/**/*" -o result.xlsx -g REF`
