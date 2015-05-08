using FarPoint.Win.Spread.CellType;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
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

            var combo = new ComboBoxCellType();
            combo.Items = new string[] { "aaa", "bbb", "ccc" };
            _spread.Sheets[1].Cells[0, 0].CellType = combo;

            var check = new CheckBoxCellType();
            _spread.Sheets[1].Cells[0, 1].CellType = check;
        }
    }
}
