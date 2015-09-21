using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms.DataVisualization.Charting;

namespace BondsMapWPF
{
    /// <summary>
    /// Логика взаимодействия для ChartWindow.xaml
    /// </summary>
    public partial class ChartWindow
    {
        public ChartWindow(DataTable[] selectedRecordsTables)
        {
            InitializeComponent();

            var notNullRecords = selectedRecordsTables.AsEnumerable().SelectMany(s=>s.AsEnumerable())
                .Where(w => !w.IsNull("Duration") && !w.IsNull("YieldClose")).ToArray();

            var minDuration = notNullRecords.Min(s => Convert.ToInt32(s["Duration"]));
            var maxDuration = notNullRecords.Max(s => Convert.ToInt32(s["Duration"]));
            var minYield = notNullRecords.Min(s => Convert.ToDouble(s["YieldClose"]));
            var maxYield = notNullRecords.Max(s => Convert.ToDouble(s["YieldClose"]));

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
                    string.Format("SecShortName+' [{0}]'", recordsTable.TableName));*/

                var notNullTable = recordsTable.AsEnumerable().
                    Where(w => !w.IsNull("Duration") && !w.IsNull("YieldClose")).ToArray();

                if (!notNullTable.Any()) continue;

                BondsMapChart.Series.Add(new Series(recordsTable.TableName)
                {
                    ChartArea = "BondsMap",
                    ChartType = SeriesChartType.Point,
                    MarkerStyle = MarkerStyle.Circle,
                    MarkerSize = 8,
                    MarkerBorderColor = Color.Black,
                });
                BondsMapChart.Series[recordsTable.TableName].Points.DataBind(notNullTable, "Duration", "YieldClose",
                    "ToolTip=ChartTip,Label=SecShortName");

                Trend trend = new Trend(notNullTable.Select(s => Convert.ToInt32(s["Duration"])).ToList(),
                    notNullTable.Select(s => Convert.ToDouble(s["YieldClose"])).ToList(), Trend.Type.Logarithmic);

                var trendDurations = Enumerable.Range(minXScale == 0 ? 1 : minXScale, maxXScale).ToArray();

                BondsMapChart.ApplyPaletteColors();
                BondsMapChart.Series.Add(new Series(recordsTable.TableName + " (Тренд)")
                {
                    Palette = ChartColorPalette.None,
                    ChartArea = "BondsMap",
                    ChartType = SeriesChartType.Line,
                    Color = BondsMapChart.Series[recordsTable.TableName].Color,
                    MarkerStyle = MarkerStyle.None
                });
                BondsMapChart.Series[recordsTable.TableName + " (Тренд)"].Points.DataBindXY(trendDurations,
                    trendDurations.Select(s => trend.Y(s)).ToArray());
            }

            BondsMapChart.Legends.Add(new Legend()
            {
                Docking = Docking.Bottom
            });
        }

        private void PaletteComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            BondsMapChart.ApplyPaletteColors();
            foreach (var seria in BondsMapChart.Series.Where(seria => seria.Name.EndsWith(" (Тренд)")))
                seria.Color = BondsMapChart.Series[seria.Name.Substring(0, seria.Name.IndexOf(" (Тренд)", StringComparison.Ordinal))].Color;
        }
    }
}
