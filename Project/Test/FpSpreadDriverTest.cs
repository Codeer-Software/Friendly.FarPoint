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
            _spread.EmulateChangeActiveSheet(1);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 0, true);
            _spread.EmualteEditText("abc");
            Assert.AreEqual("abc", _spread.ActiveSheet.ActiveCell.Text);

            _spread.EmulateChangeActiveSheet(1);
            _spread.ActiveSheet.EmulateChangeActiveCell(0, 3, true);
            _spread.EmualteEditText("bbb");
            Assert.AreEqual("bbb", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 4, true);
            _spread.EmualteEditText("1234");
            Assert.AreEqual(((char)0xa5) + "1234", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 5, true);
            _spread.EmualteEditText("2015/01/01");
            Assert.AreEqual("2015/01/01", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 8, true);
            _spread.EmualteEditText("2012/01/01 01:01:01");
            Assert.AreEqual("2012/01/01 01:01:01", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 9, true);
            _spread.EmualteEditText("abc");
            Assert.AreEqual("abc", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 10, true);
            _spread.EmualteEditText("abc");
            Assert.AreEqual("abc", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 14, true);
            _spread.EmualteEditText("123");
            Assert.AreEqual("123_", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 17, true);
            _spread.EmualteEditText("333");
            Assert.AreEqual("333.00", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 18, true);
            _spread.EmualteEditText("50");
            Assert.AreEqual("50%", _spread.ActiveSheet.ActiveCell.Text);


            _spread.ActiveSheet.EmulateChangeActiveCell(0, 20, true);
            _spread.EmualteEditText("222");
            Assert.AreEqual("222", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 21, true);
            _spread.EmualteEditText("abc");
            Assert.AreEqual("abc", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 23, true);
            _spread.EmualteEditText("efg");
            Assert.AreEqual("efg", _spread.ActiveSheet.ActiveCell.Text);
        }

        [TestMethod]
        public void EditSelectedIndex()
        {
            _spread.EmulateChangeActiveSheet(1);

            //comboBox
            _spread.ActiveSheet.EmulateChangeActiveCell(0, 3, true);
            _spread.EmualteEditSelectedIndex(1);
            Assert.AreEqual("bbb", _spread.ActiveSheet.ActiveCell.Text);

            //listbox
            _spread.ActiveSheet.EmulateChangeActiveCell(0, 13, true);
            _spread.EmualteEditSelectedIndex(1);
            Assert.AreEqual("b", _spread.ActiveSheet.ActiveCell.Text);
        }

        [TestMethod]
        public void EditValue()
        {
            _spread.EmulateChangeActiveSheet(1);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 15, true);

            _spread.EmualteEditValue("200");
            Assert.AreEqual("200", _spread.ActiveSheet.ActiveCell.Text);

            _spread.ActiveSheet.EmulateChangeActiveCell(0, 16, true);
            _spread.EmualteEditValue(2);
            Assert.AreEqual("3", _spread.ActiveSheet.ActiveCell.Text);
        }

        [TestMethod]
        public void EditCheckState()
        {
            _spread.EmulateChangeActiveSheet(1);
            _spread.ActiveSheet.EmulateChangeActiveCell(0, 2, true);
            _spread.EmualteEditCheckState(CheckState.Checked);
            Assert.AreEqual("True", _spread.ActiveSheet.ActiveCell.Text);
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
