using Codeer.Friendly;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.Column.
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.Columnに対応した操作を提供します。
    /// </summary>
#endif
    public class ColumnDriver: IAppVarOwner
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
        /// The index of this column or the start index of the column range.
        /// </summary>
#else
        /// <summary>
        /// この列のインデックスまたは、列範囲の開始インデックスを取得します。
        /// </summary>
#endif
        public int Index { get { return (int)AppVar["Index"]().Core; } }

#if ENG
        /// <summary>
        /// The end index of this column range.
        /// </summary>
#else
        /// <summary>
        /// この列範囲の終了インデックスを取得します。
        /// </summary>
#endif
        public int Index2 { get { return (int)AppVar["Index2"]().Core; } }

        internal ColumnDriver(AppVar appVar)
        {
            AppVar = appVar;
        }
    }
}
