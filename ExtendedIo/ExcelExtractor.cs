using ExcelDataReader;
using Microsoft.Analytics.Interfaces;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExtendedIo
{
    [SqlUserDefinedExtractor(AtomicFileProcessing = true)]
    public class ExcelExtractor : IExtractor
    {
        private readonly int _skipFirstNRows;
        private readonly bool _verifyHeader;
        private readonly int _sheetNumber;
        private readonly string _sheetName;
        public ExcelExtractor(int skipFirstNRows = 0, bool verifyHeader = false, int sheetNumber = 0, string sheetName = null)
        {
            _skipFirstNRows = skipFirstNRows;
            _verifyHeader = verifyHeader;
            _sheetNumber = sheetNumber;
            _sheetName = sheetName;
        }
        public override IEnumerable<IRow> Extract(IUnstructuredReader input, IUpdatableRow output)
        {
            byte[] buffer = new byte[input.Length];
            input.BaseStream.Read(buffer, 0, (int)input.Length);
            var memoryStream = new MemoryStream(buffer);

            using (var reader = ExcelReaderFactory.CreateReader(memoryStream))
            {
                //check if sheet name is passed. if not then read the passed/default sheet number
                var dataSet = reader.AsDataSet();
                DataTable sheet;
                if (_sheetName == null)
                {
                    sheet = dataSet.Tables[_sheetNumber];
                }
                else
                {
                    sheet = dataSet.Tables[_sheetName];
                }
                var rows = sheet.Rows;
                var numOfRows = rows.Count;

                //verify if the column names are matching
                if (_verifyHeader)
                {
                    int colIndex = 0;
                    foreach (var col in output.Schema)
                    {
                        var value = rows[0].ItemArray[colIndex++].ToString().ToLowerInvariant().Trim();
                        if (!col.Name.ToLowerInvariant().Trim().Equals(value))
                        {
                            throw new ColumnNameMismatchException($"Error at column index {colIndex - 1}. Expected column name: {col.Name}. Got column name: {value}");
                        }
                    }
                }

                //get all data after skipping first n rows
                for (var i = _skipFirstNRows; i < numOfRows; i++)
                {
                    int colIndex = 0;
                    foreach (var col in output.Schema)
                    {
                        var value = rows[i].ItemArray[colIndex++];
                        var typedValue = Convert.ChangeType(value, col.Type);
                        output.Set<object>(col.Name, typedValue);
                    }
                    yield return output.AsReadOnly();
                }
            }
        }
    }
}
