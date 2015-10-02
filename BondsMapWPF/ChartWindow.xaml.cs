using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace BondsMapWPF
{
    /// <summary>
    /// Логика взаимодействия для ChartWindow.xaml
    /// </summary>
    public partial class ChartWindow
    {
        public ChartWindow(BondsGroup[] selectedRecordsTables)
        {
            InitializeComponent();

            var notNullRecords = selectedRecordsTables.AsEnumerable().SelectMany(s=>s.BondItems)
                .Where(w => w.Duration.HasValue && w.YieldClose.HasValue).ToArray();

            var minDuration = notNullRecords.Min(s => Convert.ToInt32(s.Duration));
            var maxDuration = notNullRecords.Max(s => Convert.ToInt32(s.Duration));
            var minYield = notNullRecords.Min(s => Convert.ToDouble(s.YieldClose));
            var maxYield = notNullRecords.Max(s => Convert.ToDouble(s.YieldClose));

            const double xInterval = 30.0;
            const double yInterval = 1.0;

            var minXScale = Convert.ToInt32(Math.Floor(minDuration/ xInterval) * xInterval);
            var maxXScale = Convert.ToInt32(Math.Ceiling(maxDuration/ xInterval) * xInterval);
            var minYScale = Math.Floor(minYield / yInterval) * yInterval;
            var maxYScale = Math.Ceiling(maxYield / yInterval) * yInterval;

            BondsMapChart.ChartAreas.Add(new ChartArea("BondsMap")
            {
                AxisX =
                {
                    
                    Title = "Дюрация",
                    Interval = xInterval,
                    Minimum = minXScale,
                    Maximum = maxXScale,
                    MajorGrid =
                    {
                        LineColor = Color.LightGray, 
                        LineDashStyle = ChartDashStyle.Dash
                    }
                },
                AxisY =
                {
                    Title = "Доходность",
                    Interval = yInterval,
                    Minimum = minYScale,
                    Maximum = maxYScale,
                    MajorGrid =
                    {
                        LineColor = Color.LightGray, 
                        LineDashStyle = ChartDashStyle.Dash
                    },
                    MinorGrid =
                    {
                        LineColor = Color.LightGray,
                        LineDashStyle = ChartDashStyle.Dash
                    }
                }
            });

            BondsMapChart.Palette = ChartColorPalette.Bright;
            PaletteComboBox.ItemsSource = Enum.GetValues(typeof(ChartColorPalette)).Cast<ChartColorPalette>();

            foreach (var recordsTable in selectedRecordsTables)
            {
                /*recordsTable.Columns.Add("ChartLabel", typeof (string),
                    string.Format("SecShortName+' [{0}]'", recordsTable.Name));*/

                var notNullTable = recordsTable.BondItems.
                    Where(w => w.Duration.HasValue && w.YieldClose.HasValue).ToArray();

                if (!notNullTable.Any()) continue;

                BondsMapChart.Series.Add(new Series(recordsTable.Name)
                {
                    ChartArea = "BondsMap",
                    ChartType = SeriesChartType.Point,
                    MarkerStyle = MarkerStyle.Circle,
                    MarkerSize = 7,
                    MarkerBorderColor = Color.Black
                });
                BondsMapChart.Series[recordsTable.Name].Points.DataBind(notNullTable, "Duration", "YieldClose",
                    "ToolTip=ChartTip,Label=SecShortName,LabelToolTip=SecurityId");

                var trend = new Trend(notNullTable.Select(s => Convert.ToInt32(s.Duration)).ToList(),
                    notNullTable.Select(s => Convert.ToDouble(s.YieldClose)).ToList(), Trend.Type.Logarithmic);

                var trendDurations = Enumerable.Range(1, maxXScale+1000).ToArray();

                BondsMapChart.ApplyPaletteColors();
                BondsMapChart.Series.Add(new Series(recordsTable.Name + " (Тренд)")
                {
                    Palette = ChartColorPalette.None,
                    ChartArea = "BondsMap",
                    ChartType = SeriesChartType.Line,
                    Color = BondsMapChart.Series[recordsTable.Name].Color,
                    MarkerStyle = MarkerStyle.None
                });
                BondsMapChart.Series[recordsTable.Name + " (Тренд)"].Points.DataBindXY(trendDurations,
                    trendDurations.Select(s => trend.Y(s)).ToArray());
            }

            BondsMapChart.Legends.Add(new Legend
            {
                Docking = Docking.Bottom
            });
        }

        private void PaletteComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BondsMapChart.ApplyPaletteColors();
            foreach (var seria in BondsMapChart.Series.Where(seria => seria.Name.EndsWith(" (Тренд)")))
                seria.Color = BondsMapChart.Series[seria.Name.Substring(0, seria.Name.IndexOf(" (Тренд)", StringComparison.Ordinal))].Color;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            BondsMapChart.Printing.PageSetup();
            BondsMapChart.Printing.PrintPreview();
        }

        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            var excel = new Excel.Application { SheetsInNewWorkbook = 1 };

            var newbook = excel.Workbooks.Add(Excel.XlWBATemplate.xlWBATChart);

            Excel.Chart chart = newbook.Charts[1];
            chart.ChartType = Excel.XlChartType.xlXYScatter;

            List<double> allDurations = new List<double>();
            List<double> allYields = new List<double>();

            BondsMapChart.ApplyPaletteColors();
            foreach (var bs in BondsMapChart.Series.Where(w => !w.Name.EndsWith(" (Тренд)")))
            {
                Excel.Series seria;

                foreach (var bond in bs.Points)
                {
                    seria = chart.SeriesCollection().NewSeries();
                    seria.XValues = bond.XValue;
                    seria.Values = bond.YValues;
                    seria.Name = bond.Label;
                    seria.ApplyDataLabels(ShowSeriesName: true);

                    seria.MarkerForegroundColor = (int)Excel.XlRgbColor.rgbBlack;
                    seria.MarkerBackgroundColor = ColorTranslator.ToOle(bond.Color);
                    seria.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                    seria.MarkerSize = 6;

                    seria.DataLabels().ShowSeriesName = true;
                    seria.DataLabels().ShowValue = false;
                    seria.DataLabels().ShowCategoryName = false;
                }

                
                seria = chart.SeriesCollection().NewSeries();
                seria.XValues = bs.Points.Select(s => s.XValue).ToArray();
                seria.Values = bs.Points.Select(s => s.YValues[0]).ToArray();
                seria.Name = bs.Name;
                Excel.Trendline trend = seria.Trendlines().Add(Type: Excel.XlTrendlineType.xlLogarithmic);
                trend.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(bs.Color);

                seria.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

                allDurations.AddRange(bs.Points.Select(s => s.XValue));
                allYields.AddRange(bs.Points.Select(s => s.YValues[0]));
            }


            chart.PlotArea.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            Excel.Axis axes = chart.Axes(Excel.XlAxisType.xlValue);
            axes.HasMajorGridlines = true;
            axes.HasMinorGridlines = true;
            axes.MajorGridlines.Border.ColorIndex = 16;
            axes.MajorGridlines.Border.Weight = Excel.XlBorderWeight.xlHairline;
            axes.MajorGridlines.Border.LineStyle = Excel.XlLineStyle.xlDash;
            axes.MinorGridlines.Border.ColorIndex = 48;
            axes.MinorGridlines.Border.Weight = Excel.XlBorderWeight.xlHairline;
            axes.MinorGridlines.Border.LineStyle = Excel.XlLineStyle.xlDot;
            axes.MinimumScale = (int)(allYields.Min() / 0.5) * 0.5;
            axes.MaximumScale = ((int)(allYields.Max() / 0.5) + 1) * 0.5;
            axes.MinorUnit = 0.1;
            axes.MajorUnit = 0.5;

            axes = chart.Axes(Excel.XlAxisType.xlCategory);
            axes.HasMajorGridlines = true;
            axes.HasMinorGridlines = true;
            axes.MajorGridlines.Border.ColorIndex = 16;
            axes.MajorGridlines.Border.Weight = Excel.XlBorderWeight.xlHairline;
            axes.MajorGridlines.Border.LineStyle = Excel.XlLineStyle.xlDash;
            axes.MinorGridlines.Border.ColorIndex = 48;
            axes.MinorGridlines.Border.Weight = Excel.XlBorderWeight.xlHairline;
            axes.MinorGridlines.Border.LineStyle = Excel.XlLineStyle.xlDot;
            axes.MinimumScale = allDurations.Min() / 30 * 30;
            axes.MaximumScale = allDurations.Max() / 30 * 30;
            axes.MinorUnit = 30;
            axes.MajorUnit = 90;

            chart.Location(Excel.XlChartLocation.xlLocationAsNewSheet);
            chart.HasTitle = true;
            chart.ChartTitle.Characters.Text = "Карта облигаций"; // + bs.Date.ToString("d");
            chart.Axes(Excel.XlAxisType.xlCategory).HasTitle = true;
            chart.Axes(Excel.XlAxisType.xlCategory).AxisTitle.Characters.Text = "Дюрация, дн.";
            chart.Axes(Excel.XlAxisType.xlValue).HasTitle = true;
            chart.Axes(Excel.XlAxisType.xlValue).AxisTitle.Characters.Text = "Доходность, %";
            chart.HasLegend = false;

            chart.PageSetup.LeftMargin = excel.InchesToPoints(0.393700787401575);
            chart.PageSetup.RightMargin = excel.InchesToPoints(0.393700787401575);
            chart.PageSetup.TopMargin = excel.InchesToPoints(0.393700787401575);
            chart.PageSetup.BottomMargin = excel.InchesToPoints(0.393700787401575);
            chart.PageSetup.HeaderMargin = excel.InchesToPoints(0.393700787401575);
            chart.PageSetup.FooterMargin = excel.InchesToPoints(0.393700787401575);

            excel.Visible = true;
        }
    }
}
