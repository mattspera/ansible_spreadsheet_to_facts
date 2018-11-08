# spreadsheet_to_facts

Read and parse an .xlsx file and output Ansible facts.
Module has been developed to efficiently parse large, data-heavy spreadsheets.

To execute example and view Ansible facts output:
`ansible-playbook parse-spreadsheet.yml -vvv`

For use within playbook, place `spreadsheet_to_facts.py` into the **library/** directory within your Ansible project directory.
