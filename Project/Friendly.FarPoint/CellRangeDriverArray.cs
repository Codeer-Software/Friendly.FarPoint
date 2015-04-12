using Codeer.Friendly;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.Model.CellRange[].
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.Model.CellRange[]に対応した操作を提供します。
    /// </summary>
#endif
    public class CellRangeDriverArray: IAppVarOwner
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
        /// Return count of array.
        /// </summary>
#else
        /// <summary>
        /// 配列数を取得します。
        /// </summary>
#endif
        public int Length { get { return (int)AppVar["Length"]().Core; } }

#if ENG
        /// <summary>
        /// Return a element of index.
        /// </summary>
        /// <param name="index">Index.</param>
        /// <returns>Element of array.</returns>
#else
        /// <summary>
        /// 指定のインデックスの配列要素を取得します。
        /// </summary>
        /// <param name="index">インデックス。</param>
        /// <returns>配列の要素</returns>
#endif
        public CellRangeDriver this[int index] { get { return new CellRangeDriver(AppVar["[]"](index)); } }

        internal CellRangeDriverArray(AppVar appVar)
        {
            AppVar = appVar;
        }
    }
}
