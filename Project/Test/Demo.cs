using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Codeer.Friendly.Windows;
using System.Diagnostics;
using Friendly.FarPoint;
using Codeer.Friendly.Dynamic;
using System.Windows.Forms;

namespace Test
{
    [TestClass]
    public class Demo
    {
        WindowsAppFriend _app;
        dynamic _spreadCore;
        FpSpreadDriver _spread;

        [TestInitialize]
        public void TestInitialize()
        {
            _app = new WindowsAppFriend(Process.Start("Target.exe"));
            _spreadCore = _app.Type<Application>().OpenForms[0]._spread;
            _spread = new FpSpreadDriver(_spreadCore);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            Process.GetProcessById(_app.ProcessId).CloseMainWindow();
        }

        [TestMethod]
        public void Friendlyを生で触った場合()
        {
            //メリット
            //  すべてのMethod,Field,Propertyを利用可能
            //デメリット
            //  dynamicなので、インテリセンスが効かない
            //  また、実行時までミスが分からない
            //  やれることが多すぎて、どうしていいか分からない
            _spreadCore.Select();
            _spreadCore.ActiveSheetIndex = 1;
            _spreadCore.ActiveSheet.SetActiveCell(1, 2, true);
            _spreadCore.StartCellEditing(EventArgs.Empty, true);
            _spreadCore.EditingControl.Text = "abc";
            _spreadCore.StopCellEditing();
        }

        [TestMethod]
        public void ラップした場合()
        {
            //メリット
            //  インテリセンスが効く
            //  必要最小のインターフェイス以外をマスク
            //  中身を使えば拡張可能
            _spread.EmulateChangeActiveSheet(1);
            _spread.ActiveSheet.EmulateChangeActiveCell(1, 2, true);
            _spread.EmulateEditText("abc");
        }
    }
}
