# istfett-excel
Excel function that checks cell contents is bold


Even Excel version 2020 does'nt offer a function that checks whether the cell content (type: variant) is written in bold or not. 

This function =isfett() returns a 1 for true (written in bold type) or a 0 for not written in bold type.

To implement the function, a new module must be created within the vba environment of the workbook and the code must be copied into it.

The application in the productive environment takes the form: [= istfett(A1)]
