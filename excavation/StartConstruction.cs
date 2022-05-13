using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class startConstruction : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //form2為圖面選取視窗，程式作法為偷偷開起在背景執行，需要用到時才顯示出來
            form2 form2 = new form2();
            form2.Show();
            form2.Visible = false;//隱藏form2
            Form4 form4 = new Form4(commandData.Application.ActiveUIDocument, form2);//傳入doc以及form2

            form4.Show();
            return Result.Succeeded;
        }
    }
}
