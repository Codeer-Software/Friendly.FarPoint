using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Codeer.Friendly.Windows;
using System.Diagnostics;
using Friendly.FarPoint;
using Codeer.Friendly.Dynamic;
using System.Windows.Forms;

namespace Test
{
    //TODO 非同期と詳細

    [TestClass]
    public class FpSpreadDriverTest
    {
        WindowsAppFriend _app;
        FpSpreadDriver _spread;

        [TestInitialize]
        public void TestInitialize()
        {
            _app = new WindowsAppFriend(Process.Start("Target.exe"));
            _spread = new FpSpreadDriver(_app.Type<Application>().OpenForms[0]._spread);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            Process.GetProcessById(_app.ProcessId).CloseMainWindow();
        }

        [TestMethod]
        public void Sheets()
        {
            Assert.AreEqual(2, _spread.Sheets.Count);
        }

        [TestMethod]
        public void ActiveSheetIndex()
        {
            Assert.AreEqual(0, _spread.ActiveSheetIndex);
            _spread.EmulateChangeActiveSheet(1);
            Assert.AreEqual(1, _spread.ActiveSheetIndex);
        }

        [TestMethod]
        public void ActiveCell()
        {
            _spread.Sheets[0].EmulateChangeActiveCell(3, 5, true);
            Assert.AreEqual(3, _spread.Sheets[0].ActiveCell.Row.Index);
            Assert.AreEqual(3, _spread.Sheets[0].ActiveCell.Row.Index2);
            Assert.AreEqual(5, _spread.Sheets[0].ActiveCell.Column.Index);
            Assert.AreEqual(5, _spread.Sheets[0].ActiveCell.Column.Index2);
        }

        [TestMethod]
        public void EditText()
        {
            _spread.EmualteEditText("abc");
            Assert.AreEqual("abc", _spread.ActiveSheet.ActiveCell.Text);
        }

        [TestMethod]
        public void EmulateAddSelection()
        {
            _spread.ActiveSheet.EmulateAddSelection(1, 2, 3, 5);
            _spread.ActiveSheet.EmulateAddSelection(10, 20, 3, 5);
            var selections = _spread.ActiveSheet.GetSelections();
            Assert.AreEqual(2, selections.Length);
            Assert.AreEqual(1, selections[0].Row);
            Assert.AreEqual(3, selections[0].RowCount);
            Assert.AreEqual(2, selections[0].Column);
            Assert.AreEqual(5, selections[0].ColumnCount);
            Assert.AreEqual(10, selections[1].Row);
            Assert.AreEqual(3, selections[1].RowCount);
            Assert.AreEqual(20, selections[1].Column);
            Assert.AreEqual(5, selections[1].ColumnCount);
        }

        [TestMethod]
        public void EmulateClearSelection()
        {
            _spread.ActiveSheet.EmulateAddSelection(1, 2, 3, 5);
            _spread.ActiveSheet.EmulateClearSelection();
            Assert.AreEqual(0, _spread.ActiveSheet.GetSelections().Length);
        }

        [TestMethod]
        public void EmulateRemoveSelection()
        {
            _spread.ActiveSheet.EmulateAddSelection(1, 2, 3, 5);
            _spread.ActiveSheet.EmulateAddSelection(10, 20, 3, 5);
            _spread.ActiveSheet.EmulateRemoveSelection(1, 2, 3, 5);
            var selections = _spread.ActiveSheet.GetSelections();
            Assert.AreEqual(1, selections.Length);
            Assert.AreEqual(10, selections[0].Row);
            Assert.AreEqual(3, selections[0].RowCount);
            Assert.AreEqual(20, selections[0].Column);
            Assert.AreEqual(5, selections[0].ColumnCount);
        }
    }
}
