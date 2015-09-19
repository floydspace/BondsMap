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
            var selectedRecordsTable = selectedRecordsTables[0];

            var notNullRecords = selectedRecordsTables.AsEnumerable().SelectMany(s=>s.AsEnumerable())
                .Where(w => !w.IsNull("Duration") && !w.IsNull("YieldClose")).ToArray();

            if (!notNullRecords.Any())
                return;

            var minDuration = notNullRecords.Min(s => Convert.ToInt32(s["Duration"]));
            var maxDuration = notNullRecords.Max(s => Convert.ToInt32(s["Duration"]));
            var minYield = notNullRecords.Min(s => Convert.ToDouble(s["YieldClose"]));
            var maxYield = notNullRecords.Max(s => Convert.ToDouble(s["YieldClose"]));

            var minXScale = Convert.ToInt32(Math.Floor(minDuration/30d)*30);
            var maxXScale = Convert.ToInt32(Math.Ceiling(maxDuration/30d)*30);

            BondsMapChart.ChartAreas.Add(new ChartArea("BondsMap")
            {
                AxisX =
                {
                    
                    Title = "Дюрация",
                    Interval = 30,
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
                    /*Interval = 1,
                    Minimum = Math.Floor(minYield / 1d) * 1,
                    Maximum = Math.Ceiling(maxYield / 1d) * 1,*/
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

            foreach (var recordsTable in selectedRecordsTables)
            {
                var notNullTable = recordsTable.AsEnumerable().
                    Where(w => !w.IsNull("Duration") && !w.IsNull("YieldClose")).ToArray();

                BondsMapChart.Series.Add(new Series(recordsTable.TableName)
                {
                    ChartArea = "BondsMap",
                    ChartType = SeriesChartType.Point,
                    MarkerStyle = MarkerStyle.Circle,
                    MarkerSize = 8,
                    MarkerBorderColor = Color.Black,
                    //MarkerColor = Color.Black
                });
                BondsMapChart.Series[recordsTable.TableName].Points.DataBind(notNullTable, "Duration", "YieldClose",
                    "ToolTip=ChartTip,Label=SecShortName");

                Trend trend = new Trend(notNullTable.Select(s => Convert.ToInt32(s["Duration"])).ToList(),
                    notNullTable.Select(s => Convert.ToDouble(s["YieldClose"])).ToList(), Trend.Type.Logarithmic);

                var trendDurations = Enumerable.Range(minXScale == 0 ? 1 : minXScale, maxXScale).ToArray();

                BondsMapChart.ApplyPaletteColors();
                BondsMapChart.Series.Add(new Series(recordsTable.TableName + "Trend")
                {
                    Palette = ChartColorPalette.None,
                    ChartArea = "BondsMap",
                    ChartType = SeriesChartType.Line,
                    Color = BondsMapChart.Series[recordsTable.TableName].Color,
                    MarkerStyle = MarkerStyle.None
                });
                BondsMapChart.Series[recordsTable.TableName + "Trend"].Points.DataBindXY(trendDurations,
                    trendDurations.Select(s => trend.Y(s)).ToArray());
            }
        }
    }
}
