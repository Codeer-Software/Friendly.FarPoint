using FarPoint.Win;
using FarPoint.Win.Spread.CellType;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Target
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            _spread.ActiveSheetIndex = 0;
            _spread.ActiveSheet.SelectionPolicy = FarPoint.Win.Spread.Model.SelectionPolicy.MultiRange;
            ((ComboBoxCellType)_spread.Sheets[1].Cells[0, 3].CellType).Items = new string[] { "aaa", "bbb", "ccc" };

            _spread.EditModeStarting += _spread_EditModeStarting;
            _spread.EditChange += _spread_EditChange;
            _spread.EditError += _spread_EditError;
            _spread.EditModeOff += _spread_EditModeOff;
            _spread.EditModeOn += _spread_EditModeOn;
            _spread.EditorFocused += _spread_EditorFocused;
            _spread.CellClick += _spread_CellClick;
            _spread.HyperLinkClicked += _spread_HyperLinkClicked1;

            _spread.ButtonClicked += _spread_ButtonClicked;

            _spread.ActiveSheetChanged += _spread_ActiveSheetChanged;
            _spread.EnterCell += _spread_EnterCell;

            DataTable table = new DataTable();
            table.Columns.Add("Code", Type.GetType("System.String"));
            table.Columns.Add("Name", Type.GetType("System.String"));
            table.Columns.Add("Kind", Type.GetType("System.String"));
            table.Rows.Add(100, "リンゴ", "果物");
            table.Rows.Add(200, "メロン", "野菜");
            table.Rows.Add(300, "パインアップル", "果物");

            ((MultiColumnComboBoxCellType)_spread.Sheets[1].Cells[0, 15].CellType).DataSourceList = table;
        }

        private void _spread_HyperLinkClicked1(object sender, FarPoint.Win.Spread.HyperLinkClickedEventArgs e)
        {
            _spread.ActiveSheet.Cells[1, 11].Text = "Clicked";
        }

        private void _spread_EnterCell(object sender, FarPoint.Win.Spread.EnterCellEventArgs e)
        {
        }

        private void _spread_ActiveSheetChanged(object sender, EventArgs e)
        {
        }

        void _spread_ButtonClicked(object sender, FarPoint.Win.Spread.EditorNotifyEventArgs e)
        {
            _spread.ActiveSheet.Cells[1, 1].Text = "Clicked";
        }

        void _spread_EditorFocused(object sender, FarPoint.Win.Spread.EditorNotifyEventArgs e)
        {
            Debug.Print("EditorFocused " + e.Row + ", " + e.Column);
        }

        void _spread_EditModeOn(object sender, EventArgs e)
        {
            Debug.Print("EditModeOn " + _spread.ActiveSheet.ActiveCell.Row.Index + ", " + _spread.ActiveSheet.ActiveCell.Column.Index);
        }

        void _spread_EditModeOff(object sender, EventArgs e)
        {
            Debug.Print("EditModeOff " + _spread.ActiveSheet.ActiveCell.Row.Index + ", " + _spread.ActiveSheet.ActiveCell.Column.Index);
        }

        void _spread_EditError(object sender, FarPoint.Win.Spread.EditErrorEventArgs e)
        {
            Debug.Print("EditError " + +e.Row + ", " + e.Column);
        }

        void _spread_EditChange(object sender, FarPoint.Win.Spread.EditorNotifyEventArgs e)
        {
            Debug.Print("EditChange " + +e.Row + ", " + e.Column);
        }

        void _spread_EditModeStarting(object sender, FarPoint.Win.Spread.EditModeStartingEventArgs e)
        {
            Debug.Print("EditModeStarting ", _spread.ActiveSheet.ActiveCell.Row.Index + ", " + _spread.ActiveSheet.ActiveCell.Column.Index);
        }

        void _spread_CellClick(object sender, FarPoint.Win.Spread.CellClickEventArgs e)
        {
            Debug.Print("CellClick " + +e.Row + ", " + e.Column);
        }
    }
}
