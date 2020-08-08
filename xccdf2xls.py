#! /usr/bin/python3
import argparse
from os import path
from glob import glob
from xml.etree.ElementTree import parse, ParseError
from openpyxl import Workbook
from openpyxl.styles import Font


def parsingFile(filePath):
    print("Reading file {} ..".format(filePath))
    try:
        tree = parse(filePath)
    except ParseError as err:
        print("\033[91mError parsing {}\033[0m".format(filePath))
        quit()
    return tree


def getXMLNS(rootElem):
    return rootElem.tag[1:].partition("}")[0] if rootElem.tag[0] == "{" else None


def computeRefResult(ruleResults):
    if "fail" in ruleResults:
        return "fail"
    if "unknown" in ruleResults:
        return "unknown"
    if "pass" in ruleResults:
        return "pass"
    if "unchecked" in ruleResults:
        return "unchecked"
    return "notselected"


def addKeyValuePairToDict(key, value, dictionary):
    # Add a (Key, Value) pair to dictionary
    # key: the key of the pair
    # value: the value of the pair
    # dictionary: the destination dictionary
    if key in dictionary:
        dictionary[key].update(value)
    else:
        dictionary[key] = value


def flatDictKeys(dictionary):
    # Flat one level dictionary by keys
    # dictionary: the one to flat
    res = list()
    for ref, rules in dictionary.items():
        res.append("[REF]: {}".format(ref))
        [res.append(r) for r in rules.keys()]
    return res


def flatDictValues(dictionary):
    # Flat one level dictionary by values
    # dictionary: the one to flat
    res = list()
    for ref, rules in dictionary.items():
        results = rules.values()
        res.append(computeRefResult(list(results)))
        [res.append(r) for r in results]
    return res


def xccdf2json(filePath, grouped=False, group="UNREFERENCED"):
    # Convert XCCDF result xml(s) to a JSON object
    # filePath: absolute path to find XML files
    # grouped: boolean to indicate if rule results must be grouped
    # groupName: group name for grouping func
    mainDict = dict()
    for machineId, file in enumerate(glob(filePath)):
        root = parsingFile(file).getroot()
        xmlns = getXMLNS(root)
        # Map Rule IDs to Rule results
        ruleDict = dict()
        for rr in root.iter("{%s}rule-result" % xmlns):
            result = rr.find("{%s}result" % xmlns)
            ruleDict[rr.get("idref")] = result.text

        refDict = ruleDict
        # Grouping means mapping Rule refs to Rule IDs with Rule results
        if (grouped):
            refDict = dict()
            for elem in root.iter("{%s}Rule" % xmlns):
                tmp = dict()
                tmp[elem.get("id")] = ruleDict.get(elem.get("id"))
                for r in elem.iter("{%s}reference" % xmlns):
                    key = r.text if r.get("href") == group else "UNREFERENCED"
                    addKeyValuePairToDict(key, tmp, refDict)
                if len(list(elem.iter("{%s}reference" % xmlns))) == 0:
                    addKeyValuePairToDict("UNREFERENCED", tmp, refDict)
        mainDict[root.find("{%s}TestResult" % xmlns).find(
            "{%s}target" % xmlns).text] = refDict
    return mainDict


grouped = None
groupName = None
parser = argparse.ArgumentParser()
parser.add_argument(
    "-p", "--path", help="XML files, alias (*) accepted, must be quoted, default \"*\"", type=str, default="*")
parser.add_argument(
    "-g", "--group", help="reference to group result by, default null", type=str, default="")
parser.add_argument(
    "-o", "--output", help="output file name, default result.xlsx", type=str, default="result.xlsx")
args = parser.parse_args()
if args.group:
    grouped = True
    groupName = args.group

files = "{}/{}.xml".format(path.abspath(path.dirname(parser.prog)), args.path)
if len(glob(files)) == 0:
    print("\033[91mNone file found!\033[0m")
    quit()

res = xccdf2json(files, grouped, groupName)

# Create Workbook and initialize Results WorkSheet
workbook = Workbook()
worksheet = workbook.active
worksheet.title = "Results"
boldFont = Font(bold=True)

for machineNum, (machineName, mapping) in enumerate(res.items()):
    if machineNum == 0:
        # Fill first column
        firstCol = flatDictKeys(mapping) if grouped else list(mapping.keys())
        for rowIndex, rowValue in enumerate(firstCol):
            cell = worksheet.cell(row=rowIndex+2, column=1)
            cell.value = rowValue
            if "[REF]" in rowValue:
                cell.font = boldFont

    # Fill machine column
    worksheet.cell(row=1, column=machineNum+2).value = machineName
    machineCol = flatDictValues(mapping) if grouped else list(mapping.values())
    for rowIndex, rowValue in enumerate(machineCol):
        cell = worksheet.cell(row=rowIndex+2, column=machineNum+2)
        cell.value = rowValue
        if "[REF]" in worksheet.cell(row=rowIndex+2, column=1).value:
            cell.font = boldFont

# Save file and quit
workbook.save(args.output)
print("\033[92mSuccessfully merged and convert!\033[0m")
