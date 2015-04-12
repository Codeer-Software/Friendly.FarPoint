using Codeer.Friendly;

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
    public class CellDriver : IAppVarOwner
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

        internal CellDriver(AppVar appVar)
        {
            AppVar = appVar;
        }
    }
}
