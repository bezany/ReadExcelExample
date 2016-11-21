using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExampleData
{
    public interface IExcelService
    {
        List<string> SearchRow(string searchName);
    }
}
