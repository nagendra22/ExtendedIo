# ExtendedIo
Extended Input and output for USQL

This is repository to add Input/Output functionality of U-SQL using User Defined extractors and outputters.

Currently there is excel extractor.

### How to use:
Build the project and register assemblies from bin folder:
* ExtendedIo.dll
* ExcelDataReader.dll
* ExcelDataReader.Dataset.dll

Use the custom Extractor from U-SQL Script

```
REFERENCE ASSEMBLY HRServicesDB.ExtendedIo;
REFERENCE ASSEMBLY HRServicesDB.ExcelDataReader;
REFERENCE ASSEMBLY HRServicesDB.[ExcelDataReader.DataSet];
USING ExtendedIo;

@excelData =
    EXTRACT id int,
            name string
    FROM "/sample/sampleExcel.xlsx"
    USING new ExcelExtractor(skipFirstNRows:1, verifyHeader:true, sheetName:"sampleSheet");
```
