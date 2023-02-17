using Microsoft.Office.Interop.Excel;
using Soneta.Business;
using Soneta.Business.Db;
using Soneta.Place;
using Soneta.Types;
using Point = System.Drawing.Point;
using System;
using System.Linq;

using ZgloszeniImportExcel;

[assembly: Worker(typeof(MyWorker1Worker), typeof(Wyplaty))]

namespace ZgloszeniImportExcel
{
    public class MyWorker1Worker
    {
        [Context]
        public Context Context { get; set; }

        [Context]
        public ContractorsInvoicedHoursSettlementWorkerParams Params { get; set; }

        public Session Session { get => Params.Session; }

        protected Application App = new Application();

        protected _Worksheet activeSheet;


        [Action("Lista w Excelu", Description = "Eksportuj listę do Microsoft Excel", Mode = ActionMode.SingleSession | ActionMode.Progress | ActionMode.OnlyTable, Target = ActionTarget.ToolbarWithText)]
        public string PayRollCreation()
        {
            App.Visible = true;
            Point start = new Point(1, 1);
            activeSheet = (_Worksheet)App.Workbooks.Add().ActiveSheet;
            Range range = (Range)activeSheet.Cells[1, 1];
            range.Value = "tekst";
            App.Quit();
            return "Eksport dokumentu zakończony";
        }
    }
    public class ContractorsInvoicedHoursSettlementWorkerParams : ContextBase
    {
        public ContractorsInvoicedHoursSettlementWorkerParams(Context context) : base(context)
        { }       

        [Required]
        [Caption("Okres zestawienia")]
        public FromTo Period { get; set; } = FromTo.Month(Date.Today);
    }
}
