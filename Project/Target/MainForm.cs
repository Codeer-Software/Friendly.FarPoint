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
        }
    }
}
