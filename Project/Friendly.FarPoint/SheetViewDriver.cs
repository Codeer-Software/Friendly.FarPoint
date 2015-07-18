using Codeer.Friendly;
using Friendly.FarPoint.Inside;
using System;
using System.Windows.Forms;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.SheetView.
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.SheetViewに対応した操作を提供します。
    /// </summary>
#endif
    public class SheetViewDriver : IAppVarOwner
    {
#if ENG
        /// <summary>
        /// Returns an AppVar for a .NET object for the corresponding window.
        /// </summary>
#else
        /// <summary>
        /// 対応するウィンドウの.Netのオブジェクトが格納されたAppVarを取得します。
        /// </summary>
#endif
        public AppVar AppVar { get; set; }

#if ENG
        /// <summary>
        /// Cells's Driver.
        /// </summary>
#else
        /// <summary>
        /// Cellsのドライバです。
        /// </summary>
#endif
        public CellsDriver Cells { get { return new CellsDriver(AppVar["Cells"]()); } }

#if ENG
        /// <summary>
        /// Active cell.
        /// </summary>
#else
        /// <summary>
        /// アクティブなセルです。
        /// </summary>
#endif
        public CellDriver ActiveCell { get { return new CellDriver(AppVar["ActiveCell"]()); } }

        internal SheetViewDriver(AppVar appVar)
        {
            AppVar = appVar;
        }

#if ENG
        /// <summary>
        /// Get selections.
        /// </summary>
        /// <returns></returns>
#else
        /// <summary>
        /// 選択範囲の取得
        /// </summary>
        /// <returns></returns>
#endif
        public CellRangeDriverArray GetSelections()
        {
            return new CellRangeDriverArray(AppVar["GetSelections"]());
        }

#if ENG
        /// <summary>
        /// Selects a certain cell.
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <param name="clearSelection">Is clear selection range.</param>
#else
        /// <summary>
        /// アクティブなシートを選択します。
        /// </summary>
        /// <param name="row">行。</param>
        /// <param name="col">列。</param>
        /// <param name="clearSelection">選択範囲をクリアするかどうか。</param>
#endif
        public void EmulateChangeActiveCell(int row, int col, bool clearSelection)
        {
            AppVar.App[GetType(), "EmulateChangeActiveCell"](AppVar, row, col, clearSelection);
        }

#if ENG
        /// <summary>
        /// Selects a certain cell.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <param name="clearSelection">Is clear selection range.</param>
        /// <param name="async">Asynchronous execution.</param>
#else
        /// <summary>
        /// アクティブなシートを選択します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="row">行。</param>
        /// <param name="col">列。</param>
        /// <param name="clearSelection">選択範囲をクリアするかどうか。</param>
        /// <param name="async">非同期オブジェクト。</param>
#endif
        public void EmulateChangeActiveCell(int row, int col, bool clearSelection, Async async)
        {
            AppVar.App[GetType(), "EmulateChangeActiveCell", async](AppVar, row, col, clearSelection);
        }
        
        static void EmulateChangeActiveCell(object sheet, int row, int col, bool clearSelection)
        {
            SelectSpread(sheet);
            Invoker.Call(sheet, "SetActiveCell", row, col, clearSelection);
        }

#if ENG
        /// <summary>
        /// Clear all selection in this sheet.
        /// </summary>
#else
        /// <summary>
        /// このシートからすべての選択範囲を削除します。
        /// </summary>
#endif
        public void EmulateClearSelection()
        {
            AppVar.App[GetType(), "EmulateClearSelection"](AppVar);
        }

#if ENG
        /// <summary>
        /// Clear all selection in this sheet.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="async">Asynchronous execution.</param>
#else
        /// <summary>
        /// このシートからすべての選択範囲を削除します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="async">非同期オブジェクト。</param>
#endif
        public void EmulateClearSelection(Async async)
        {
            AppVar.App[GetType(), "EmulateClearSelection", async](AppVar);
        }

        static void EmulateClearSelection(object sheet)
        {
            SelectSpread(sheet);
            Invoker.Call(sheet, "ClearSelection");
        }

#if ENG
        /// <summary>
        /// Add selection.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <param name="rowCount">row count.</param>
        /// <param name="columnCount">column count.</param>
#else
        /// <summary>
        /// 選択を追加します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="row">行。</param>
        /// <param name="col">列。</param>
        /// <param name="rowCount">行数。</param>
        /// <param name="columnCount">列数。</param>
#endif
        public void EmulateAddSelection(int row, int col, int rowCount, int columnCount)
        {
            AppVar.App[GetType(), "EmulateAddSelection"](AppVar, row, col, rowCount, columnCount);
        }

#if ENG
        /// <summary>
        /// Add selection.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <param name="rowCount">row count.</param>
        /// <param name="columnCount">column count.</param>
        /// <param name="async">Asynchronous execution.</param>
#else
        /// <summary>
        /// 選択を追加します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="row">行。</param>
        /// <param name="col">列。</param>
        /// <param name="rowCount">行数。</param>
        /// <param name="columnCount">列数。</param>
        /// <param name="async">非同期オブジェクト。</param>
#endif
        public void EmulateAddSelection(int row, int col, int rowCount, int columnCount, Async async)
        {
            AppVar.App[GetType(), "EmulateAddSelection", async](AppVar, row, col, rowCount, columnCount);
        }

        static void EmulateAddSelection(object sheet, int row, int col, int rowCount, int columnCount)
        {
            SelectSpread(sheet);
            Invoker.Call(sheet, "AddSelection", row, col, rowCount, columnCount);
        }

#if ENG
        /// <summary>
        /// Add selection.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <param name="rowCount">row count.</param>
        /// <param name="columnCount">column count.</param>
#else
        /// <summary>
        /// 選択を追加します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="row">行。</param>
        /// <param name="col">列。</param>
        /// <param name="rowCount">行数。</param>
        /// <param name="columnCount">列数。</param>
#endif
        public void EmulateRemoveSelection(int row, int col, int rowCount, int columnCount)
        {
            AppVar.App[GetType(), "EmulateRemoveSelection"](AppVar, row, col, rowCount, columnCount);
        }

#if ENG
        /// <summary>
        /// Remove selection.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <param name="rowCount">row count.</param>
        /// <param name="columnCount">column count.</param>
        /// <param name="async">Asynchronous execution.</param>
#else
        /// <summary>
        /// 選択を削除します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="row">行。</param>
        /// <param name="col">列。</param>
        /// <param name="rowCount">行数。</param>
        /// <param name="columnCount">列数。</param>
        /// <param name="async">非同期オブジェクト。</param>
#endif
        public void EmulateRemoveSelection(int row, int col, int rowCount, int columnCount, Async async)
        {
            AppVar.App[GetType(), "EmulateRemoveSelection", async](AppVar, row, col, rowCount, columnCount);
        }
        static void EmulateRemoveSelection(object sheet, int row, int col, int rowCount, int columnCount)
        {
            SelectSpread(sheet);
            Invoker.Call(sheet, "RemoveSelection", row, col, rowCount, columnCount);
        }

        static void SelectSpread(object sheet)
        {
            Control spread = (Control)sheet.GetType().GetProperty("FpSpread").GetGetMethod().Invoke(sheet, new object[] { });
            spread.Select();
            spread.Focus();
        }
    }
}
