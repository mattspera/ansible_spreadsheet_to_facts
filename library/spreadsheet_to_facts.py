#!/usr/bin/python

# Copyright: (c) 2018, Matthew Spera <speramatthew@gmail.com>
# GNU General Public License v3.0+ (see COPYING or https://www.gnu.org/licenses/gpl-3.0.txt)

ANSIBLE_METADATA = {
    'metadata_version': '1.1',
    'status': ['preview'],
    'supported_by': 'community'
}

DOCUMENTATION = '''
---
module: spreadsheet_to_facts

short_description: Read and parse an .xlsx file and output Ansible facts.

description:
    - Read and parse an .xlsx file and output Ansible facts.
    
requirements:
    - openpyxl can be obtained from PyPi (https://pypi.org/project/openpyxl)

options:
    src:
        description:
            - The name of the spreadsheet.
        required: true
    sheets:
        description:
            - List of worksheets within workbook to parse.
        type: list

author:
    - Matthew Spera (@mattspera)
'''

EXAMPLES = '''
# Parse whole spreadsheet
- name: Read spreadsheet
  spreadsheet_to_facts:
    src: test_spreadsheet.xlsx

# Parse only specific worksheets within spreadsheet
- name: Read sheet 1 & 2 from spreadsheet
  spreadsheet_to_facts:
    src: test_spreadsheet.xlsx
    sheets: ['Sheet1', 'Sheet2']
'''

RETURN = '''
facts_json:
    description: Spreadsheet data output as Ansible facts.
message:
    description: The output message generated.
'''

from ansible.module_utils.basic import AnsibleModule

try:
    import openpyxl
    
    HAS_LIB = True
except ImportError:
    HAS_LIB = False
    
def parse_xlsx_dict(input_file, sheet_list):
    result = {'ansible_facts':{}}
    spreadsheet = {}
    
    try:
        workbook = openpyxl.load_workbook(input_file, read_only=True)
    except IOError:
        return(1, 'IOError on input file: ' + input_file)
        
    if sheet_list:
        worksheets = sheet_list
    else:
        worksheets = workbook.get_sheet_names()
        
    for sheet in worksheets:
        ansible_sheet_name = 'sheet_' + sheet
        spreadsheet[ansible_sheet_name] = []
        current_sheet = workbook[sheet]
        
        header = []
        for cell in current_sheet[1]:
            header.append(cell.value)
            
        for row in list(current_sheet.rows)[1:]:
            row_data = {}
            for header_key, cell_value in zip(header, row):
                row_data[header_key] = cell_value.value
            spreadsheet[ansible_sheet_name].append(row_data)
            
    result['ansible_facts'] = spreadsheet
    
    return(0, result)

def main():
    module_args = dict(
        src = dict(type='str', required=True),
        sheets = dict(type='list')
    )

    result = dict(
        changed=False,
        facts_json='',
        message=''
    )

    module = AnsibleModule(
        argument_spec=module_args,
        supports_check_mode=False
    )
    if not HAS_LIB:
        module.fail_json(msg='Missing required library: openpyxl')
        
    ret_code, response = parse_xlsx_dict(module.params['src'], module.params['sheets'])
    
    if ret_code == 1:
        result['message'] = response
        module.fail_json(**result)
    else:
        result['facts_json'] = response
        result['message'] = 'Done'
    
    module.exit_json(**result)

if __name__ == '__main__':
    main()