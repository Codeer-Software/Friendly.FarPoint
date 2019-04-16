using System.Drawing;
using Codeer.Friendly;
using Codeer.Friendly.Windows;
using Codeer.Friendly.Windows.Grasp;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.Cell.
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.Cellに対応した操作を提供します。
    /// </summary>
#endif
    public class CellDriver : IAppVarOwner, IUIObject
    {
        SheetViewDriver _sheet;

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
        /// Column's Driver.
        /// </summary>
#else
        /// <summary>
        /// Columnのドライバです。
        /// </summary>
#endif
        public ColumnDriver Column { get { return new ColumnDriver(AppVar["Column"]()); } }

#if ENG
        /// <summary>
        /// Row's Driver.
        /// </summary>
#else
        /// <summary>
        /// Rowのドライバです。
        /// </summary>
#endif
        public RowDriver Row { get { return new RowDriver(AppVar["Row"]()); } }

#if ENG
        /// <summary>
        /// Text.
        /// </summary>
#else
        /// <summary>
        /// セル内の書式付きテキスト。
        /// </summary>
#endif
        public string Text { get { return (string)AppVar["Text"]().Core; } }

#if ENG
        /// <summary>
        /// Returns the associated application manipulation object.
        /// </summary>
#else
        /// <summary>
        /// アプリケーション操作クラスを取得します。
        /// </summary>
#endif
        public WindowsAppFriend App => (WindowsAppFriend)AppVar.App;

#if ENG
        /// <summary>
        /// Returns the size of IUIObject.
        /// </summary>
#else
        /// <summary>
        /// IUIObjectのサイズを取得します。
        /// </summary>
#endif
        public Size Size => GetRectangle().Size;

#if ENG
        /// <summary>
        /// Convert IUIObject's client coordinates to screen coordinates.
        /// </summary>
        /// <param name="clientPoint">client coordinates.</param>
        /// <returns>screen coordinates.</returns>
#else
        /// <summary>
        /// IUIObjectのクライアント座標からスクリーン座標に変換します。
        /// </summary>
        /// <param name="clientPoint">クライアント座標</param>
        /// <returns>スクリーン座標</returns>
#endif
        public Point PointToScreen(Point clientPoint)
        {
            var location = GetRectangle().Location;
            clientPoint.Offset(location.X, location.Y);
            return _sheet.Spread.PointToScreen(clientPoint);
        }

#if ENG
        /// <summary>
        /// Make it active.
        /// </summary>
#else
        /// <summary>
        /// アクティブな状態にします。
        /// </summary>
#endif
        public void Activate() => _sheet.Spread.Activate();

        internal CellDriver(SheetViewDriver sheet, AppVar appVar)
        {
            _sheet = sheet;
            AppVar = appVar;
        }

        private Rectangle GetRectangle()
        {
            var row = Row.Index;
            var column = Column.Index;
            var rowViewportIndex = row < _sheet.FrozenRowCount ? -1 : 0;
            var columnViewportIndex = column < _sheet.FrozenRowCount ? -1 : 0;
            var rc = (Rectangle)_sheet.Spread["GetCellRectangle"](rowViewportIndex, columnViewportIndex, row, column).Core;
            return rc;
        }
    }
}
