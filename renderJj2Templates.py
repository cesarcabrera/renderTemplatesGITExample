import jinja2
import pandas as pd
import openpyxl
import json
import re
from datetime import datetime

DEB_LEVEL = 1

if __name__ == '__main__':
    start = datetime.now()
    date_prefix = start.strftime('%Y_%m_%d_%M')

    templateFolder, systemSeparator = '.', '/'
    configFile, variablesFile = 'config.xlsx', 'parameters.xlsx'

    wb = openpyxl.load_workbook(templateFolder + systemSeparator + variablesFile)
    sectionGroups = wb.sheetnames
    print(f"INFO: detected {len(sectionGroups)} sectionGroups")
    print(f"INFO: {sectionGroups}")
    env = jinja2.Environment()
    for g in sectionGroups:
        # Read template
        df = pd.read_excel(templateFolder + "/" + configFile, sheet_name=g)
        template, variables = '', []
        for row in df.iterrows():
            linea = str(row[1].values[0]) + '\n'
            var = re.findall(r"{{([\w-]+)}}", linea)
            if len(var) > 1:
                variables += var
            template += linea
        # Read variables
        df = pd.read_excel(templateFolder + "/" + variablesFile, sheet_name=g)
        print(f"DEB: Detected {len(variables)} variables")
        data = {}
        for row in df.iterrows():
            name, value = row[1][0], row[1].values[1]
            linea = str(value)
            data[name] = ''
            if re.match(r'{', linea):
                print(f"DEB: i = {row[0]}, name = {name} (Reading JSON)")
                content = json.loads(linea)
            else:
                print(f"DEB: i = {row[0]}, name = {name} (Reading Text)")
                content = linea
            data[name] = content
            print(f"DEB: {name} => {data[name]}")
        config = env.from_string(str(template))
        print(f">>> START Config for {g}")
        print(config.render(data))
        print(f"<<< END Config for {g}")

