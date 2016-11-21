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
            var file = new FileInfo(_fullFileName);
            using (var ep = new ExcelPackage(file, true))
            {
                var dt = ep.ToDataSet(firstRowContainsHeader: true);
                var firstWs = dt.Tables[0];
                var res = new List<string>();
                foreach (var row in firstWs.Rows)
                {
                    //res.Add(row)
                }
                var sheetStartRow = 1;
                var ws = ep.Workbook.Worksheets[0];
                for (int i = sheetStartRow; i < ws.Dimension.End.Row; i++)
                {
                    var curRow = ws.Row(i);
                    //curRow.
                }
                var firstColumnData = ws.Cells[sheetStartRow, 1, ws.Dimension.End.Row, 1];
                var test = firstColumnData.FirstOrDefault(el => ((el.Value as string) ?? "").Contains(searchName));
                //test.Address
            }

            throw new NotImplementedException();
        }
    }
}
