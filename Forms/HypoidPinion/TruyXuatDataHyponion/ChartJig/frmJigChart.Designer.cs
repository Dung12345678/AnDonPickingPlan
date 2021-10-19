
namespace BMS
{
	partial class frmJigChart
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			DevExpress.XtraCharts.XYDiagram xyDiagram1 = new DevExpress.XtraCharts.XYDiagram();
			DevExpress.XtraCharts.SecondaryAxisY secondaryAxisY1 = new DevExpress.XtraCharts.SecondaryAxisY();
			DevExpress.XtraCharts.Series series1 = new DevExpress.XtraCharts.Series();
			DevExpress.XtraCharts.PointSeriesLabel pointSeriesLabel1 = new DevExpress.XtraCharts.PointSeriesLabel();
			DevExpress.XtraCharts.LineSeriesView lineSeriesView1 = new DevExpress.XtraCharts.LineSeriesView();
			DevExpress.XtraCharts.Series series2 = new DevExpress.XtraCharts.Series();
			DevExpress.XtraCharts.LineSeriesView lineSeriesView2 = new DevExpress.XtraCharts.LineSeriesView();
			DevExpress.XtraCharts.Series series3 = new DevExpress.XtraCharts.Series();
			DevExpress.XtraCharts.LineSeriesView lineSeriesView3 = new DevExpress.XtraCharts.LineSeriesView();
			DevExpress.XtraCharts.LineSeriesView lineSeriesView4 = new DevExpress.XtraCharts.LineSeriesView();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.chartControl1 = new DevExpress.XtraCharts.ChartControl();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.chartControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(xyDiagram1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(secondaryAxisY1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(series1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(pointSeriesLabel1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(series2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(series3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView4)).BeginInit();
			this.SuspendLayout();
			// 
			// chartControl1
			// 
			this.chartControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.chartControl1.AppearanceNameSerializable = "Chameleon";
			this.chartControl1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
			this.chartControl1.BorderOptions.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			this.chartControl1.BorderOptions.Visible = false;
			xyDiagram1.AxisX.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisX.Label.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisX.Title.Font = new System.Drawing.Font("Tahoma", 10F);
			xyDiagram1.AxisX.Title.Text = "Ngày/Tháng";
			xyDiagram1.AxisX.VisibleInPanesSerializable = "-1";
			xyDiagram1.AxisX.WholeRange.AutoSideMargins = true;
			xyDiagram1.AxisY.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisY.Label.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisY.Title.Font = new System.Drawing.Font("Tahoma", 10F);
			xyDiagram1.AxisY.Title.Text = "";
			xyDiagram1.AxisY.Title.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.AxisY.Title.Visible = true;
			xyDiagram1.AxisY.VisibleInPanesSerializable = "-1";
			xyDiagram1.DefaultPane.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
			xyDiagram1.DefaultPane.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			xyDiagram1.DefaultPane.FillStyle.FillMode = DevExpress.XtraCharts.FillMode.Solid;
			xyDiagram1.Margins.Bottom = 13;
			xyDiagram1.Margins.Left = 13;
			xyDiagram1.Margins.Right = 13;
			xyDiagram1.Margins.Top = 13;
			xyDiagram1.PaneDistance = 18;
			secondaryAxisY1.AxisID = 0;
			secondaryAxisY1.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			secondaryAxisY1.Label.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			secondaryAxisY1.Name = "Secondary AxisY 1";
			secondaryAxisY1.Title.Font = new System.Drawing.Font("Tahoma", 10F);
			secondaryAxisY1.Title.Text = "";
			secondaryAxisY1.Title.TextColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			secondaryAxisY1.Title.Visible = true;
			secondaryAxisY1.VisibleInPanesSerializable = "-1";
			xyDiagram1.SecondaryAxesY.AddRange(new DevExpress.XtraCharts.SecondaryAxisY[] {
            secondaryAxisY1});
			this.chartControl1.Diagram = xyDiagram1;
			this.chartControl1.EmptyChartText.Text = "Không tải được dữ liêu\r\n";
			this.chartControl1.Legend.AlignmentHorizontal = DevExpress.XtraCharts.LegendAlignmentHorizontal.Center;
			this.chartControl1.Legend.AlignmentVertical = DevExpress.XtraCharts.LegendAlignmentVertical.TopOutside;
			this.chartControl1.Legend.Direction = DevExpress.XtraCharts.LegendDirection.LeftToRight;
			this.chartControl1.Legend.EquallySpacedItems = false;
			this.chartControl1.Legend.Margins.Bottom = 0;
			this.chartControl1.Legend.Margins.Left = 0;
			this.chartControl1.Legend.Margins.Right = 0;
			this.chartControl1.Legend.Margins.Top = 0;
			this.chartControl1.Legend.UseCheckBoxes = true;
			this.chartControl1.Location = new System.Drawing.Point(156, 109);
			this.chartControl1.Name = "chartControl1";
			this.chartControl1.Padding.Bottom = 4;
			this.chartControl1.Padding.Left = 4;
			this.chartControl1.Padding.Right = 4;
			this.chartControl1.Padding.Top = 4;
			this.chartControl1.PaletteName = "Palette 1";
			this.chartControl1.PaletteRepository.Add("Palette 1", new DevExpress.XtraCharts.Palette("Palette 1", DevExpress.XtraCharts.PaletteScaleMode.Repeat, new DevExpress.XtraCharts.PaletteEntry[] {
                new DevExpress.XtraCharts.PaletteEntry(System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(206)))), ((int)(((byte)(52))))), System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(206)))), ((int)(((byte)(52)))))),
                new DevExpress.XtraCharts.PaletteEntry(System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(202)))), ((int)(((byte)(202))))), System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(202)))), ((int)(((byte)(202)))))),
                new DevExpress.XtraCharts.PaletteEntry(System.Drawing.Color.Red, System.Drawing.Color.Red)}));
			this.chartControl1.RuntimeHitTesting = true;
			this.chartControl1.RuntimeSeriesSelectionMode = DevExpress.XtraCharts.SeriesSelectionMode.Point;
			series1.CheckableInLegend = false;
			pointSeriesLabel1.Angle = 90;
			pointSeriesLabel1.Antialiasing = true;
			pointSeriesLabel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			pointSeriesLabel1.LineColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			pointSeriesLabel1.LineStyle.DashStyle = DevExpress.XtraCharts.DashStyle.Dot;
			pointSeriesLabel1.LineStyle.Thickness = 2;
			pointSeriesLabel1.LineVisible = false;
			pointSeriesLabel1.ResolveOverlappingMode = DevExpress.XtraCharts.ResolveOverlappingMode.HideOverlapped;
			series1.Label = pointSeriesLabel1;
			series1.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True;
			series1.Name = "Thực tế";
			lineSeriesView1.LineMarkerOptions.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			lineSeriesView1.LineMarkerOptions.Color = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			lineSeriesView1.LineMarkerOptions.FillStyle.FillMode = DevExpress.XtraCharts.FillMode.Hatch;
			lineSeriesView1.LineMarkerOptions.Size = 7;
			lineSeriesView1.MarkerVisibility = DevExpress.Utils.DefaultBoolean.True;
			series1.View = lineSeriesView1;
			series2.CheckableInLegend = false;
			series2.Name = "Max";
			series2.View = lineSeriesView2;
			series3.CheckableInLegend = false;
			series3.Name = "Min";
			series3.View = lineSeriesView3;
			this.chartControl1.SeriesSerializable = new DevExpress.XtraCharts.Series[] {
        series1,
        series2,
        series3};
			this.chartControl1.SeriesTemplate.View = lineSeriesView4;
			this.chartControl1.SideBySideBarDistanceVariable = 0.1D;
			this.chartControl1.Size = new System.Drawing.Size(788, 337);
			this.chartControl1.TabIndex = 222;
			// 
			// frmJigChart
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1108, 662);
			this.Controls.Add(this.chartControl1);
			this.Name = "frmJigChart";
			this.Text = "BIỂU ĐỒ GIA CÔNG";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.frmKnifeProcessedChart_Load);
			((System.ComponentModel.ISupportInitialize)(secondaryAxisY1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(xyDiagram1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(pointSeriesLabel1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(series1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(series2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(series3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(lineSeriesView4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chartControl1)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion
		private System.Windows.Forms.Timer timer1;
		private DevExpress.XtraCharts.ChartControl chartControl1;
		private System.Windows.Forms.ToolTip toolTip1;
	}
}