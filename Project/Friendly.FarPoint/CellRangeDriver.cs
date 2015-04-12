using Codeer.Friendly;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.Model.CellRange.
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.Model.CellRangeに対応した操作を提供します。
    /// </summary>
#endif
    public class CellRangeDriver : IAppVarOwner
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
        /// Column.
        /// </summary>
#else
        /// <summary>
        /// 列。
        /// </summary>
#endif
        public int Column { get { return (int)AppVar["Column"]().Core; } }

#if ENG
        /// <summary>
        /// Column count.
        /// </summary>
#else
        /// <summary>
        /// 列数。
        /// </summary>
#endif
        public int ColumnCount { get { return (int)AppVar["ColumnCount"]().Core; } }

#if ENG
        /// <summary>
        /// Row.
        /// </summary>
#else
        /// <summary>
        /// 行。
        /// </summary>
#endif
        public int Row { get { return (int)AppVar["Row"]().Core; } }

#if ENG
        /// <summary>
        /// Row count.
        /// </summary>
#else
        /// <summary>
        /// 行数。
        /// </summary>
#endif
        public int RowCount { get { return (int)AppVar["RowCount"]().Core; } }

        internal CellRangeDriver(AppVar appVar)
        {
            AppVar = appVar;
        }
    }
}
