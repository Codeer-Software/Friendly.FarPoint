namespace Target
{
    partial class MainForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this._spread = new FarPoint.Win.Spread.FpSpread();
            this._spread_Sheet1 = new FarPoint.Win.Spread.SheetView();
            this._spread_Sheet2 = new FarPoint.Win.Spread.SheetView();
            ((System.ComponentModel.ISupportInitialize)(this._spread)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._spread_Sheet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this._spread_Sheet2)).BeginInit();
            this.SuspendLayout();
            // 
            // _spread
            // 
            this._spread.AccessibleDescription = "";
            this._spread.Dock = System.Windows.Forms.DockStyle.Fill;
            this._spread.Location = new System.Drawing.Point(0, 0);
            this._spread.Name = "_spread";
            this._spread.Sheets.AddRange(new FarPoint.Win.Spread.SheetView[] {
            this._spread_Sheet1,
            this._spread_Sheet2});
            this._spread.Size = new System.Drawing.Size(284, 261);
            this._spread.TabIndex = 0;
            this._spread.ActiveSheetIndex = 1;
            // 
            // _spread_Sheet1
            // 
            this._spread_Sheet1.Reset();
            this._spread_Sheet1.SheetName = "Sheet1";
            // 
            // _spread_Sheet2
            // 
            this._spread_Sheet2.Reset();
            this._spread_Sheet2.SheetName = "Sheet2";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this._spread);
            this.Name = "MainForm";
            this.Text = "Target";
            ((System.ComponentModel.ISupportInitialize)(this._spread)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._spread_Sheet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this._spread_Sheet2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private FarPoint.Win.Spread.FpSpread _spread;
        private FarPoint.Win.Spread.SheetView _spread_Sheet1;
        private FarPoint.Win.Spread.SheetView _spread_Sheet2;
    }
}

