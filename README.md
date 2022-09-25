# VIEScheckerMakro

Purpose:
Macro use data from example excel file and validate all VAT numbers from table.
Validation based via API from https://ec.europa.eu/taxation_customs/vies/#/vat-validation

Required references:
- https://github.com/VBA-tools/VBA-JSON
- Microsoft Scripting Runtime
- Microsoft WinHTTP Services, version 5.1

Work flow:
Read non empty cells in the request column.
Iterate thrue all requests and validate VAT numbers.
