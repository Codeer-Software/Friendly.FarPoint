using System;
using System.Collections.Generic;
using Codeer.TestAssistant.GeneratorToolKit;

namespace Friendly.FarPoint.Generator
{
    [CaptureCodeGenerator("Friendly.FarPoint.FpSpreadDriver")]
    public class FpSpreadDriverGenerator : CaptureCodeGeneratorBase
    {
        TypeFinder _typeFinder = new TypeFinder();
        string _beforeValue = "";
        dynamic _spread;
        List<Action> _removes = new List<Action>();

        protected override void Attach()
        {
            _spread = ControlObject;
            _removes.Add(EventAdapter.Add(ControlObject, "ActiveSheetChanged", ActiveSheetChanged));
            _removes.Add(EventAdapter.Add(ControlObject, "EnterCell", EnterCell));
            _removes.Add(EventAdapter.Add(ControlObject, "ButtonClicked", ButtonClicked));
            _removes.Add(EventAdapter.Add(ControlObject, "EditModeOn", EditModeOn));
            _removes.Add(EventAdapter.Add(ControlObject, "EditModeOff", EditModeOff));
        }

        protected override void Detach() => _removes.ForEach(e => e());

        void EnterCell(object sender, dynamic e)
        {
            AddSentence(new TokenName(), ".ActiveSheet.EmulateChangeActiveCell(", e.Row, ", " , e.Column, new TokenAsync(CommaType.Before), ");");
        }

        void ActiveSheetChanged(object sender, dynamic e)
        {
            AddSentence(new TokenName(), ".EmulateChangeActiveSheet(", _spread.ActiveSheetIndex, new TokenAsync(CommaType.Before), ");");
        }

        void ButtonClicked(object sender, dynamic e)
        {
            if (_spread.ActiveSheet == null || _spread.ActiveSheet.ActiveCell == null) return;
            if (_spread.ActiveSheet.ActiveCell.CellType == null) return;

            var type = _spread.ActiveSheet.ActiveCell.CellType.GetType();
            if (!_typeFinder.GetType("FarPoint.Win.Spread.CellType.ButtonCellType").IsAssignableFrom(type)) return;

            AddSentence(new TokenName(), ".EmulateButtonClick(", new TokenAsync(CommaType.Non), ");");
        }

        void EditModeOn(object sender, dynamic e)
        {
            if (_spread.ActiveSheet == null || _spread.ActiveSheet.ActiveCell == null) return;
            _beforeValue = _spread.ActiveSheet.ActiveCell.Text;
        }

        void EditModeOff(object sender, dynamic e)
        {
            if (_spread.ActiveSheet == null || _spread.ActiveSheet.ActiveCell == null) return;

            var text = _spread.ActiveSheet.ActiveCell.Text;
            if (_beforeValue == text) return;

            if (_spread.ActiveSheet.ActiveCell.CellType != null)
            {
                var type = _spread.ActiveSheet.ActiveCell.CellType.GetType();
                if (_typeFinder.GetType("FarPoint.Win.Spread.CellType.ButtonCellType").IsAssignableFrom(type))
                {
                    return;
                }
                else if (_typeFinder.GetType("FarPoint.Win.Spread.CellType.CheckBoxCellType").IsAssignableFrom(type))
                {
                    AddUsingNamespace("System.Windows.Forms");
                    var val = (text == "True" || text == "Checked") ? "CheckState.Checked" :
                              (text == "False" || text == "Unchecked") ? "CheckState.Unchecked" : "CheckState.Indeterminate";
                    AddSentence(new TokenName(), ".EmulateEditCheckState(", val, new TokenAsync(CommaType.Before), ");");
                    return;
                }
                else if (_typeFinder.GetType("FarPoint.Win.Spread.CellType.MultiColumnComboBoxCellType").IsAssignableFrom(type) ||
                         _typeFinder.GetType("FarPoint.Win.Spread.CellType.MultiOptionCellType").IsAssignableFrom(type) ||
                         _typeFinder.GetType("FarPoint.Win.Spread.CellType.SliderCellType").IsAssignableFrom(type))
                {
                    AddSentence(new TokenName(), ".EmulateEditValue(", text, new TokenAsync(CommaType.Before), ");");
                    return;
                }
            }
            AddSentence(new TokenName(), ".EmulateEditText(", GenerateUtility.AdjustText(text), new TokenAsync(CommaType.Before), ");");
        }

        public override void Optimize(List<Sentence> code)
        {
            GenerateUtility.RemoveDuplicationFunction(this, code, "EmulateChangeActiveSheet");
            GenerateUtility.RemoveDuplicationFunction(this, code, "ActiveSheet.EmulateChangeActiveCell");
        }
    }
}
