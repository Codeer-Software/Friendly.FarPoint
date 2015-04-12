using Codeer.Friendly;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;
using System;
using System.Windows.Forms;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.FpSpread.
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.FpSpreadに対応した操作を提供します。
    /// </summary>
#endif
    public class FpSpreadDriver : IAppVarOwner
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
        /// Active sheet index.
        /// </summary>
#else
        /// <summary>
        /// アクティブなシートのインデックスです。
        /// </summary>
#endif
        public int ActiveSheetIndex 
        {
            get { return (int)AppVar["ActiveSheetIndex"]().Core; }
        }

#if ENG
        /// <summary>
        /// SheetViewCollection's Driver.
        /// </summary>
#else
        /// <summary>
        /// SheetViewCollectionのドライバです。
        /// </summary>
#endif
        public SheetViewCollectionDriver Sheets { get { return new SheetViewCollectionDriver(AppVar["Sheets"]()); } }

#if ENG
        /// <summary>
        /// Active sheet.
        /// </summary>
#else
        /// <summary>
        /// アクティブなシートです。
        /// </summary>
#endif
        public SheetViewDriver ActiveSheet { get { return new SheetViewDriver(AppVar["ActiveSheet"]()); } }

#if ENG
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="src">WindowControl object for the underlying control.</param>
#else
        /// <summary>
        /// コンストラクタです。
        /// </summary>
        /// <param name="src">元となるウィンドウコントロールです。</param>
#endif
        public FpSpreadDriver(WindowControl src) : this(src.AppVar) { }

#if ENG
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="appVar">Application variable object for the control.</param>
#else
        /// <summary>
        /// コンストラクタです。
        /// </summary>
        /// <param name="appVar">アプリケーション内変数。</param>
#endif
        public FpSpreadDriver(AppVar appVar)
        {
            AppVar = appVar;
            WindowsAppExpander.LoadAssembly((WindowsAppFriend)AppVar.App, typeof(FpSpreadDriver).Assembly);
        }

#if ENG
        /// <summary>
        /// Selects a certain sheet.
        /// </summary>
        /// <param name="index">Index (0-based) of the sheet to select.</param>
#else
        /// <summary>
        /// アクティブなシートを選択します。
        /// </summary>
        /// <param name="index">シートインデックス。</param>
#endif
        public void EmulateChangeActiveSheet(int index)
        {
            AppVar.App[GetType(), "EmulateChangeActiveSheet"](AppVar, index);
        }

#if ENG
        /// <summary>
        /// Selects a certain sheet.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="index">Index (0-based) of the sheet to select.</param>
        /// <param name="async">Asynchronous execution.</param>
#else
        /// <summary>
        /// アクティブなシートを選択します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="index">シートインデックス。</param>
        /// <param name="async">非同期オブジェクト。</param>
#endif
        public void EmulateChangeActiveSheet(int index, Async async)
        {
            AppVar.App[GetType(), "EmulateChangeActiveSheet", async](AppVar, index);
        }

        static void EmulateChangeActiveSheet(Control spread, int index)
        {
            spread.Select();
            spread.Focus();
            spread.GetType().GetProperty("ActiveSheetIndex").GetSetMethod().Invoke(spread, new object[] { index });
        }

#if ENG
        /// <summary>
        /// Modifies the text of a active cell.
        /// </summary>
        /// <param name="text">The text to use.</param>
#else
        /// <summary>
        /// アクティブなセルのテキストを変更します。
        /// </summary>
        /// <param name="text">テキスト。</param>
#endif
        public void EmualteEditText(string text)
        {

            AppVar.App[GetType(), "EmualteEditText"](AppVar, text);
        }
#if ENG
        /// <summary>
        /// Modifies the text of a active cell.
        /// Executes asynchronously. 
        /// </summary>
        /// <param name="text">The text to use.</param>
        /// <param name="async">Asynchronous execution.</param>
#else
        /// <summary>
        /// アクティブなセルのテキストを変更します。
        /// 非同期で実行します。
        /// </summary>
        /// <param name="text">テキスト。</param>
        /// <param name="async">非同期実行オブジェクト。</param>
#endif
        public void EmualteEditText(string text, Async async)
        {
            AppVar.App[GetType(), "EmualteEditText", async](AppVar, text);
        }

        static void EmualteEditText(Control spread, string text)
        {
            spread.Select();
            spread.Focus();
            spread.GetType().GetMethod("StartCellEditing").Invoke(spread, new object[] { EventArgs.Empty, true });
            Control edit = (Control)spread.GetType().GetProperty("EditingControl").GetGetMethod().Invoke(spread, new object[] { });
            edit.Text = text;
            spread.GetType().GetMethod("StopCellEditing").Invoke(spread, new object[0]);
        }
    }
}
