using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Xml;
using MessageBox = System.Windows.Forms.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace BondsMapWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly DataSet _reportsSet;
        private int _i;
        public MainWindow()
        {
            InitializeComponent();

            _reportsSet = new DataSet();
            _reportsSet.ReadXmlSchema("SEM21_02062014.xsd");
            _reportsSet.Tables["BOARD"].Columns.Add("TradeDate", typeof(DateTime), @"Parent(SEM21_BOARD).TradeDate");
            _reportsSet.Tables["RECORDS"].Columns.Add("TradeDate", typeof(DateTime), @"Parent(BOARD_RECORDS).TradeDate");
            _reportsSet.Tables["RECORDS"].Columns.Add("BoardName", typeof(string), @"Parent(BOARD_RECORDS).BoardName");
            _reportsSet.Tables["RECORDS"].Columns.Add("DurationYears", typeof(double), "Duration/365");
            _reportsSet.Tables["RECORDS"].Columns.Add("ChartTip", typeof(string), @"'Доходность: '+YieldClose+' | Дюрация: '+Duration");

            var existingXmlFiles = Directory.GetFiles(Environment.CurrentDirectory, @"*_SEM21_*.xml");
            foreach (var existingXmlFile in existingXmlFiles) FillReportSet(existingXmlFile);

            FillCalendar();
        }

        private void CreateGroup(string s)
        {
            var groupTable = new DataTable(s);
            foreach (DataColumn column in _reportsSet.Tables["RECORDS"].Columns)
                groupTable.Columns.Add(column.ColumnName, column.DataType,
                    column.ColumnName.Equals("TradeDate") || column.ColumnName.Equals("BoardName")
                        ? string.Empty
                        : column.Expression);
            
            GroupsComboBox.Items.Add(groupTable);
            GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
        }

        private void FillCalendar()
        {
            if (!_reportsSet.Tables.Contains("SEM21")) return;

            var availableDates = _reportsSet.Tables["SEM21"].AsEnumerable().Select(s => (DateTime)s["TradeDate"]);

            if (!availableDates.Any()) return;

            var startDate = availableDates.Min();
            var endDate = availableDates.Max();
            CalendarReports.DisplayDateStart = startDate;
            CalendarReports.DisplayDateEnd = endDate;
            CalendarReports.DisplayDate = endDate;
            CalendarReports.BlackoutDates.Clear();
            for (var date = startDate; date <= endDate; date = date.AddDays(1))
                if (!availableDates.Contains(date)) CalendarReports.BlackoutDates.Add(new CalendarDateRange(date));

            CalendarReports.IsEnabled = true;
            CalendarReports.SelectedDate = endDate;
        }

        private bool FillReportSet(string xmlFile)
        {
            try
            {
                _reportsSet.ReadXml(xmlFile);
                return true;
            }
            catch (XmlException xmlEx)
            {
                MessageBox.Show(string.Format("Файл {0} не является файлом биржевой информации (SEM21)",
                    Path.GetFileName(xmlFile)), xmlEx.Source,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void LoadReportsButton_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Файлы биржевой информации (SEM21)|*_SEM21_*.xml"
            };
            if (ofd.ShowDialog() != true) return;

            foreach (var fileName in ofd.FileNames)
            {
                if (!FillReportSet(fileName)) continue;
                var fileShortName = Path.GetFileName(fileName);
                if (fileShortName != null && !File.Exists(Path.Combine(Environment.CurrentDirectory, fileShortName)))
                    File.Copy(fileName, Path.Combine(Environment.CurrentDirectory, fileShortName));
            }

            FillCalendar();
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FillFoundedRecordsListBox();
        }

        private void FillFoundedRecordsListBox()
        {
            if (!CalendarReports.IsEnabled) return;
            var report = _reportsSet.Tables["SEM21"].AsEnumerable().First(w => w["TradeDate"].Equals(CalendarReports.SelectedDate));
            var boards = report.GetChildRows("SEM21_BOARD").Where(w => w["BoardType"].Equals("MAIN"));
            var records = boards.SelectMany(s => s.GetChildRows("BOARD_RECORDS")).Where(w => w["SecurityType"].Equals("об"));
            var foundedRecords = records.Where(w =>
                        ((string)w["SecurityId"]).ToLowerInvariant().Contains(SearchTextBox.Text.ToLowerInvariant()) ||
                        ((string)w["SecShortName"]).ToLowerInvariant().Contains(SearchTextBox.Text.ToLowerInvariant()) ||
                        ((string)w["EngName"]).ToLowerInvariant().Contains(SearchTextBox.Text.ToLowerInvariant()) ||
                        ((string)w["RegNumber"]).ToLowerInvariant().Contains(SearchTextBox.Text.ToLowerInvariant())).ToArray();

            FoundedRecordsListBox.ItemsSource = foundedRecords.Any() ? foundedRecords.CopyToDataTable().DefaultView : new DataView();
        }

        private void CalendarReports_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            FillFoundedRecordsListBox();
        }

        private void AddAll_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                    " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);

            if (GroupsComboBox.SelectedItem != null)
                FoundedRecordsListBox.Items.Cast<DataRowView>().Select(s => s.Row)
                    .CopyToDataTable((DataTable)GroupsComboBox.SelectedItem, LoadOption.OverwriteChanges);
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                    " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);

            if (GroupsComboBox.SelectedItem != null)
                FoundedRecordsListBox.SelectedItems.Cast<DataRowView>().Select(s => s.Row)
                    .CopyToDataTable((DataTable)GroupsComboBox.SelectedItem, LoadOption.OverwriteChanges);
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            while (SelectedRecordsDataGrid.SelectedItems.Count > 0)
                ((DataTable)GroupsComboBox.SelectedItem).Rows.Remove(((DataRowView)SelectedRecordsDataGrid.SelectedItem).Row);
        }

        private void RemoveAll_Click(object sender, RoutedEventArgs e)
        {
            ((DataTable)GroupsComboBox.SelectedItem).Rows.Clear();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            Remove_Click(null, null);
        }

        private void BuildBondsMapButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedRecordsTables = GroupsComboBox.Items.Cast<DataTable>().ToArray();
            foreach (var selectedRecordsTable in selectedRecordsTables)
                selectedRecordsTable.AcceptChanges();
            
            if (selectedRecordsTables.SelectMany(s => s.AsEnumerable()).All(w => w.RowState == DataRowState.Deleted || w.IsNull("Duration") || w.IsNull("YieldClose"))) return;
            ChartWindow chartWindow = new ChartWindow(selectedRecordsTables);
            chartWindow.Show();
        }

        private void CreateGroupButton_Click(object sender, RoutedEventArgs e)
        {
            new InputBox(
                string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                    " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);
        }

        private void GroupsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectedRecordsDataGrid.ItemsSource = e.AddedItems.Count > 0
                ? e.AddedItems.Cast<DataTable>().First().DefaultView
                : new DataView();
        }

        private void DeleteGroupButton_Click(object sender, RoutedEventArgs e)
        {
            GroupsComboBox.Items.Remove(GroupsComboBox.SelectedItem);
            GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
        }

        private void AddNewRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                    string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                        " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(
                            this);

            var table = ((DataView) SelectedRecordsDataGrid.ItemsSource).Table;
            table.Rows.Add(table.NewRow());
        }

        private void ImportFromExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                    string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                        " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(
                            this);

            var ofd = new OpenFileDialog
            {
                Filter = "Файлы CSV|*.csv"
            };
            if (ofd.ShowDialog() != true) return;

            var lines = File.ReadAllLines(ofd.FileName);

            const int columnNo = 0;
            var isins = lines.Select(s => s.Split(';').Count() > columnNo ? s.Split(';')[columnNo].Trim() : "")
                    .Distinct().ToArray();

            if (GroupsComboBox.SelectedItem != null)
                FoundedRecordsListBox.Items.Cast<DataRowView>().Select(s => s.Row)
                    .Where(w => isins.Contains(w["SecurityId"]) || isins.Contains(w["RegNumber"]))
                    .CopyToDataTable((DataTable)GroupsComboBox.SelectedItem, LoadOption.OverwriteChanges);
        }

        private void CalendarReports_GotMouseCapture(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (Mouse.Captured is Calendar || Mouse.Captured is System.Windows.Controls.Primitives.CalendarItem)
                Mouse.Capture(null);
        }
    }
}
