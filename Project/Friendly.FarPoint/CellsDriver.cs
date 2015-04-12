using Codeer.Friendly;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.Cells.
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.Cellsに対応した操作を提供します。
    /// </summary>
#endif
    public class CellsDriver : IAppVarOwner
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
        /// Cell's Driver.
        /// </summary>
        /// <param name="tag">tag.</param>
        /// <returns>SheetView's Driver.</returns>
#else
        /// <summary>
        /// Cellのドライバです。
        /// </summary>
        /// <param name="tag">タグ。</param>
        /// <returns>SheetViewのドライバです。</returns>
#endif
        public CellDriver this[string tag] { get { return new CellDriver(AppVar["[]"](tag)); } }

#if ENG
        /// <summary>
        /// Cell's Driver.
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <returns>SheetView's Driver.</returns>
#else
        /// <summary>
        /// Cellのドライバです。
        /// </summary>
        /// <param name="row">行。</param>
        /// <param name="col">列。</param>
        /// <returns>SheetViewのドライバです。</returns>
#endif
        public CellDriver this[int row, int col] { get { return new CellDriver(AppVar["[,]"](row, col)); } }

#if ENG
        /// <summary>
        /// Cell's Driver.
        /// </summary>
        /// <param name="row">row.</param>
        /// <param name="col">col.</param>
        /// <param name="row2">row2.</param>
        /// <param name="col2">col2.</param>
        /// <returns>SheetView's Driver.</returns>
#else
        /// <summary>
        /// Cellのドライバです。
        /// </summary>
        /// <param name="row">範囲開始行。</param>
        /// <param name="col">範囲開始列。</param>
        /// <param name="row2">範囲最終行。</param>
        /// <param name="col2">範囲最終列。</param>
        /// <returns>SheetViewのドライバです。</returns>
#endif
        public CellDriver this[int row, int col, int row2, int col2] { get { return new CellDriver(AppVar["[,,,]"](row, col, row2, col2)); } }

        internal CellsDriver(AppVar appVar)
        {
            AppVar = appVar;
        }
    }

}
