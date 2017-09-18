using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace XLSUtil.Library
{
    public class Model2XLS : IDisposable
    {
        private Workbooks workbooks { get; set; }
        private Worksheet worksheet { get; set; }
        private Workbook workbook { get; set; }
        private Sheets sheets { get; set; }
        private Application application { get; set; }

        public Model2XLS(params string[] labelList)
        {
            application = new Application();
            workbooks = application.Workbooks;
            workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            sheets = workbook.Sheets;
            worksheet = (Worksheet)sheets[1];

        }
        public void AddHeader(params string[] labelList)
        {
            for (int i = 0; i < labelList.Count(); i++)
            {
                Range range = worksheet.get_Range(CellAddress(worksheet, 1, i + 1));
                range.Value = labelList[i];
            }
        }
        private string CellAddress(Worksheet worksheet, int row, int col)
        {
            return RangeAddress(worksheet.Cells[row, col]);
        }

        private string RangeAddress(object range)
        {
            return ((Range)range).get_AddressLocal(false, false, XlReferenceStyle.xlA1);
        }

        public void AddContent(List<List<object>> rawData, bool hasHeader)
        {
            for (int i = 0; i < rawData.Count(); i++)
            {
                var line = rawData[i];
                for (int j = 0; j < line.Count(); j++)
                {
                    var column = line[j];

                    Range range = worksheet.get_Range(CellAddress(worksheet, i + 1 + (hasHeader ? 1 : 0), j + 1));

                    if (column is String)
                        range.Value = "'" + column;
                    else if (column is Boolean)
                        range.Value = ((bool)column == true ? 1 : 0);
                    else
                        range.Value = column;
                }
            }
        }

        public void Dispose()
        {
            if (application.Visible == true)
                application.Visible = false;
            else
                application.Visible = true;

            worksheet.Columns.AutoFit();
            worksheet.Rows.AutoFit();

            application.ActiveWindow.Activate();

            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(application);
        }
    }
}
