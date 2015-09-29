using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using FloydSpace.MICEX.Portable;
using Microsoft.Win32;

namespace BondsMapWPF
{
    public class BondsGroup
    {
        public string TableName { get; set; }
        public ObservableCollection<Bond> BondItems { get; set; }
    }

    public class Bond : INotifyPropertyChanged 
    {
        private DateTime? _tradeDate;
        private string _boardName;
        private string _securityId;
        private string _secShortName;
        private string _regNumber;
        private string _currencyId;
        private double? _yieldClose;
        private int? _duration;
        private double? _durationYears;
        private DateTime? _matDate;

        public DateTime? TradeDate
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

        public double? YieldClose
        {
            get { return _yieldClose; }
            set
            {
                _yieldClose = value;
                OnPropertyChanged(new PropertyChangedEventArgs("YieldClose"));
            }
        }

        public int? Duration
        {
            get { return _duration; }
            set
            {
                _duration = value;
                OnPropertyChanged(new PropertyChangedEventArgs("Duration"));
            }
        }

        public double? DurationYears
        {
            get { return _durationYears; }
            set
            {
                _durationYears = value;
                OnPropertyChanged(new PropertyChangedEventArgs("DurationYears"));
            }
        }

        public DateTime? MatDate
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

        private void FillDataGrid(IEnumerable<MicexGrabber.SecurityItem> secItems)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                    string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "Группа " + ++_i : SearchTextBox.Text,
                        " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);

            if (GroupsComboBox.SelectedItem == null) return;

            ProgressImage.Visibility = Visibility.Visible;
            Grid1.IsEnabled = false;
            ThreadPool.QueueUserWorkItem(state =>
            {
                Parallel.ForEach(secItems, item =>
                {
                    var bond = new Bond();
                    CalendarReports.Dispatcher.Invoke(new Action(() => { bond.TradeDate = CalendarReports.SelectedDate; }));
                    bond.SecurityId = item.SecId;
                    bond.SecShortName = item.ShortName;
                    bond.RegNumber = item.RegNumber;


                    var secBoardGroup = MicexGrabber.GetSecurityBoardGroups(bond.SecurityId).First();
                    var boardGroupInfo = MicexGrabber.GetBoardGroupInfo(secBoardGroup.Id, bond.SecurityId);
                    bond.BoardName = boardGroupInfo.BoardItem.Title;
                    var marketData = MicexGrabber.GetMarketData(bond.TradeDate ?? DateTime.Now, bond.SecurityId,
                        boardGroupInfo);
                    if (marketData.Equals(new MicexGrabber.MarketData())) return;
                    bond.YieldClose = marketData.YieldClose;
                    bond.Duration = marketData.Duration;
                    bond.DurationYears = bond.Duration/365d;
                    bond.CurrencyId = marketData.CurrencyId;
                    bond.MatDate = marketData.MatDate;

                    GroupsComboBox.Dispatcher.Invoke(
                        new Action(() => { ((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Add(bond); }));
                });
                ProgressImage.Dispatcher.Invoke(new Action(() =>
                {
                    ProgressImage.Visibility = Visibility.Hidden;
                    Grid1.IsEnabled = true;
                }));
            }, null);
        }
        private void AddAll_Click(object sender, RoutedEventArgs e)
        {
            FillDataGrid(FoundedRecordsListBox.Items.Cast<MicexGrabber.SecurityItem>());
        }
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            FillDataGrid(FoundedRecordsListBox.SelectedItems.Cast<MicexGrabber.SecurityItem>());
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            while (SelectedRecordsDataGrid.SelectedItems.Count > 0)
                ((BondsGroup)GroupsComboBox.SelectedItem).BondItems.Remove(((Bond)SelectedRecordsDataGrid.SelectedItem));
        }

        private void RemoveAll_Click(object sender, RoutedEventArgs e)
        {
            ((BondsGroup)GroupsComboBox.SelectedItem).BondItems.Clear();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            Remove_Click(null, null);
        }

        private void BuildBondsMapButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedRecordsTables = GroupsComboBox.Items.Cast<BondsGroup>().ToArray();
            if (selectedRecordsTables.SelectMany(s => s.BondItems)
                .All(w => !w.Duration.HasValue || !w.YieldClose.HasValue)) return;
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

            if (GroupsComboBox.SelectedItem != null)
                ((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Add(new Bond());
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

        private void CalendarReports_GotMouseCapture(object sender, MouseEventArgs e)
        {
            if (Mouse.Captured is Calendar || Mouse.Captured is CalendarItem)
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
