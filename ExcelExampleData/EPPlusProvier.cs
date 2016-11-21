using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using EPPlus.Extensions;

namespace ExcelExampleData
{
    public class EPPlusProvier : FileBaseService
    {
        public EPPlusProvier(string fileFullName)
            : base(fileFullName)
        {
        }

        public override List<string> SearchRow(string searchName)
        {
            if (string.IsNullOrEmpty(searchName))
            {
                return null;
            }
            using (var fileStream = new FileStream(_fullFileName, FileMode.Open, FileAccess.Read))
            using (var ep = new ExcelPackage(fileStream))
            {
                var ws = ep.Workbook.Worksheets["test"];
                for (int i = ws.Dimension.Start.Row; i <= ws.Dimension.End.Row; i++)
                {
                    var curRow = ws.Row(i);
                    var cell = ws.Cells[curRow.Row, 1];
                    var firstColval = cell.GetValue<string>();
                    if ((firstColval ?? "")
                        .ToUpper()
                        .Contains(searchName.ToUpper()))
                    {
                        var res = new List<string>();

                        for (int col = ws.Dimension.Start.Column; col <= ws.Dimension.End.Column; col++)
                        {
                            res.Add(ws.Cells[curRow.Row, col].GetValue<string>());
                        }

                        return res;
                    }
                }
                return null;
             }
            
        }
    }
}
