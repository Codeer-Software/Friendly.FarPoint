using Codeer.Friendly;

namespace Friendly.FarPoint
{
#if ENG
    /// <summary>
    /// Provides operations on controls of type FarPoint.Win.Spread.SheetViewCollection.
    /// </summary>
#else
    /// <summary>
    /// TypeがFarPoint.Win.Spread.SheetViewCollectionに対応した操作を提供します。
    /// </summary>
#endif
    public class SheetViewCollectionDriver : IAppVarOwner
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
        /// Return count of collection.
        /// </summary>
#else
        /// <summary>
        /// 要素数を取得します。
        /// </summary>
#endif
        public int Count { get { return (int)AppVar["Count"]().Core; } }
        
#if ENG
        /// <summary>
        /// SheetView's Driver.
        /// </summary>
        /// <param name="index">index.</param>
        /// <returns>SheetView's Driver.</returns>
#else
        /// <summary>
        /// SheetViewのドライバです。
        /// </summary>
        /// <param name="index">インデックス。</param>
        /// <returns>SheetViewのドライバです。</returns>
#endif
        public SheetViewDriver this[int index] { get { return new SheetViewDriver(AppVar["[]"](index)); } }

        internal SheetViewCollectionDriver(AppVar appVar)
        {
            AppVar = appVar;
        }
    }
}
