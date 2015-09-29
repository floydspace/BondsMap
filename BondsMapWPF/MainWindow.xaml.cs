using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Xml;
using FloydSpace.MICEX.Portable;
using MessageBox = System.Windows.Forms.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace BondsMapWPF
{
    class BondsGroup
    {
        public string TableName { get; set; }
        public ObservableCollection<Bond> BondItems { get; set; }
    }

    class Bond : INotifyPropertyChanged 
    {
        private DateTime _tradeDate;
        private string _boardName;
        private string _securityId;
        private string _secShortName;
        private string _regNumber;
        private string _currencyId;
        private double _yieldClose;
        private int _duration;
        private double _durationYears;
        private DateTime _matDate;

        public DateTime TradeDate
        {
            get { return _tradeDate; }
            set
            {
                _tradeDate = value;
                OnPropertyChanged(new PropertyChangedEventArgs("TradeDate"));
            }
        }

        public string BoardName
        {
            get { return _boardName; }
            set
            {
                _boardName = value;
                OnPropertyChanged(new PropertyChangedEventArgs("BoardName"));
            }
        }

        public string SecurityId
        {
            get { return _securityId; }
            set
            {
                _securityId = value;
                OnPropertyChanged(new PropertyChangedEventArgs("SecurityId"));
            }
        }

        public string SecShortName
        {
            get { return _secShortName; }
            set
            {
                _secShortName = value;
                OnPropertyChanged(new PropertyChangedEventArgs("SecShortName"));
            }
        }

        public string RegNumber
        {
            get { return _regNumber; }
            set
            {
                _regNumber = value;
                OnPropertyChanged(new PropertyChangedEventArgs("RegNumber"));
            }
        }

        public string CurrencyId
        {
            get { return _currencyId; }
            set
            {
                _currencyId = value;
                OnPropertyChanged(new PropertyChangedEventArgs("CurrencyId"));
            }
        }

        public double YieldClose
        {
            get { return _yieldClose; }
            set
            {
                _yieldClose = value;
                OnPropertyChanged(new PropertyChangedEventArgs("YieldClose"));
            }
        }

        public int Duration
        {
            get { return _duration; }
            set
            {
                _duration = value;
                OnPropertyChanged(new PropertyChangedEventArgs("Duration"));
            }
        }

        public double DurationYears
        {
            get { return _durationYears; }
            set
            {
                _durationYears = value;
                OnPropertyChanged(new PropertyChangedEventArgs("DurationYears"));
            }
        }

        public DateTime MatDate
        {
            get { return _matDate; }
            set
            {
                _matDate = value;
                OnPropertyChanged(new PropertyChangedEventArgs("MatDate"));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, e);
        }
    }

    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private int _i;
        public MainWindow()
        {
            InitializeComponent();

            CalendarReports.SelectedDate = DateTime.Today;
            /*var favoritesDirectory = Path.Combine(Environment.CurrentDirectory, "Favorites");
            if (!Directory.Exists(favoritesDirectory)) return;
            foreach (var favoritesFile in Directory.GetFiles(favoritesDirectory))
            {
                var groupTable = new DataTable();
                groupTable.ReadXml(favoritesFile);
                GroupsComboBox.Items.Add(groupTable);
                GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
            }*/
        }

        private void CreateGroup(string s)
        {
            var bondsGroup = new BondsGroup { TableName = s, BondItems = new ObservableCollection<Bond>() };
            GroupsComboBox.Items.Add(bondsGroup);
            GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
        }

        private async void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (SearchTextBox.Text.Length < 3) return;
            var securities = await MicexGrabber.FindSecuritiesAsync(SearchTextBox.Text);
            FoundedRecordsListBox.ItemsSource = securities.Where(w=>w.Group.Contains("bonds")).OrderBy(o=>o.ShortName);
        }

        private async void AddAll_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                    " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);

            if (GroupsComboBox.SelectedItem == null) return;
            foreach (MicexGrabber.SecurityItem item in FoundedRecordsListBox.Items)
            {
                var bond = new Bond
                {
                    TradeDate = CalendarReports.SelectedDate ?? DateTime.Now,
                    SecurityId = item.SecId,
                    SecShortName = item.ShortName,
                    RegNumber = item.RegNumber
                };
                var board = (await MicexGrabber.GetSecurityBoardsAsync(bond.SecurityId)).First();
                bond.BoardName = board.Title;
                var boardGroupInfo = await MicexGrabber.GetBoardGroupInfoAsync(board.BoardGroupId, bond.SecurityId);
                var marketData = await MicexGrabber.GetMarketDataAsync(bond.TradeDate, bond.SecurityId, boardGroupInfo);
                bond.YieldClose = marketData.YieldClose ?? new double();
                bond.Duration = marketData.Duration ?? new int();
                bond.DurationYears = bond.Duration / 365d;
                bond.CurrencyId = marketData.CurrencyId;
                bond.MatDate = marketData.MatDate ?? new DateTime();

                ((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Add(bond);
            }
        }

        private async void Add_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                    " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);

            if (GroupsComboBox.SelectedItem == null) return;
            foreach (MicexGrabber.SecurityItem item in FoundedRecordsListBox.SelectedItems)
            {
                var bond = new Bond
                {
                    TradeDate = CalendarReports.SelectedDate ?? DateTime.Now,
                    SecurityId = item.SecId,
                    SecShortName = item.ShortName,
                    RegNumber = item.RegNumber
                };
                var board = (await MicexGrabber.GetSecurityBoardsAsync(bond.SecurityId)).First();
                bond.BoardName = board.Title;
                var boardGroupInfo = await MicexGrabber.GetBoardGroupInfoAsync(board.BoardGroupId, bond.SecurityId);
                var marketData = await MicexGrabber.GetMarketDataAsync(bond.TradeDate, bond.SecurityId, boardGroupInfo);
                bond.YieldClose = marketData.YieldClose ?? new double();
                bond.Duration = marketData.Duration ?? new int();
                bond.DurationYears = bond.Duration / 365d;
                bond.CurrencyId = marketData.CurrencyId;
                bond.MatDate = marketData.MatDate ?? new DateTime();

                ((BondsGroup)GroupsComboBox.SelectedItem).BondItems.Add(bond);
            }
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
            new ChartWindow(selectedRecordsTables).Show();
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
                ? e.AddedItems.Cast<BondsGroup>().First().BondItems
                : new ObservableCollection<Bond>();

            ((Image)FavoritesButton.Content).Source = !IsInFavorites
                    ? new BitmapImage(new Uri(@"pack://application:,,,/Images/favorAdd.jpg", UriKind.RelativeOrAbsolute))
                    : new BitmapImage(new Uri(@"pack://application:,,,/Images/favorRemove.jpg", UriKind.RelativeOrAbsolute));
        }

        private void DeleteGroupButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedTable = GroupsComboBox.SelectedItem as DataTable;
            if (selectedTable == null) return;
            var favoritesDirectory = Path.Combine(Environment.CurrentDirectory, "Favorites");
            var fileName = Path.Combine(favoritesDirectory, selectedTable.TableName + ".xml");
            if (File.Exists(fileName)) File.Delete(fileName);

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
            var isins = lines.Where(w =>
            {
                var temp = w.Split(';');
                return temp.Length > columnNo && temp[columnNo] != string.Empty;
            }).Select(s => s.Split(';')[columnNo].Trim())
                .Distinct().ToArray();

            if (GroupsComboBox.SelectedItem == null) return;
            ((DataTable)GroupsComboBox.SelectedItem).AcceptChanges();
            FoundedRecordsListBox.Items.Cast<DataRowView>().Select(s => s.Row)
                .Where(w => isins.Contains(w["SecurityId"]) || isins.Contains(w["RegNumber"]))
                .Where(w => !((DataTable)GroupsComboBox.SelectedItem).Rows.Cast<DataRow>()
                    .Any(row => (string)row["BoardName"] == (string)w["BoardName"] &&
                                (string)row["SecurityId"] == (string)w["SecurityId"]))
                .CopyToDataTable((DataTable)GroupsComboBox.SelectedItem, LoadOption.OverwriteChanges);
        }

        private void CalendarReports_GotMouseCapture(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (Mouse.Captured is Calendar || Mouse.Captured is System.Windows.Controls.Primitives.CalendarItem)
                Mouse.Capture(null);
        }

        public bool IsInFavorites {
            get
            {
                var selectedTable = GroupsComboBox.SelectedItem as DataTable;
                if (selectedTable == null) return false;

                var fileName = Path.Combine(Environment.CurrentDirectory, "Favorites", selectedTable.TableName + ".xml");
                return File.Exists(fileName);
            }
        }

        private void FavoritesButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedTable = GroupsComboBox.SelectedItem as DataTable;
            if (selectedTable == null) return;

            var favoritesDirectory = Path.Combine(Environment.CurrentDirectory, "Favorites");
            var fileName = Path.Combine(favoritesDirectory, selectedTable.TableName + ".xml");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
                ((Image)FavoritesButton.Content).Source =
                new BitmapImage(new Uri(@"pack://application:,,,/Images/favorAdd.jpg", UriKind.RelativeOrAbsolute));
            }
            else
            {
                Directory.CreateDirectory(favoritesDirectory);
                selectedTable.WriteXml(fileName, XmlWriteMode.WriteSchema);
                ((Image)FavoritesButton.Content).Source =
                new BitmapImage(new Uri(@"pack://application:,,,/Images/favorRemove.jpg", UriKind.RelativeOrAbsolute));
            }
        }
    }
}
