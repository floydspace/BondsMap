using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Xml;
using System.Xml.Serialization;
using FloydSpace.MICEX.Portable;
using Microsoft.Win32;

namespace BondsMapWPF
{
    public class BondsGroup
    {
        public string Name { get; set; }
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
        private DateTime? _buyBackDate;
        private DateTime? _matDate;

        public DateTime? TradeDate
        {
            get { return _tradeDate; }
            set
            {
                _tradeDate = value;
                OnPropertyChanged();
            }
        }

        public string BoardName
        {
            get { return _boardName; }
            set
            {
                _boardName = value;
                OnPropertyChanged();
            }
        }

        public string SecurityId
        {
            get { return _securityId; }
            set
            {
                _securityId = value;
                OnPropertyChanged();
            }
        }

        public string SecShortName
        {
            get { return _secShortName; }
            set
            {
                _secShortName = value;
                OnPropertyChanged();
            }
        }

        public string RegNumber
        {
            get { return _regNumber; }
            set
            {
                _regNumber = value;
                OnPropertyChanged();
            }
        }

        public string CurrencyId
        {
            get { return _currencyId; }
            set
            {
                _currencyId = value;
                OnPropertyChanged();
            }
        }

        public double? YieldClose
        {
            get { return _yieldClose; }
            set
            {
                _yieldClose = value;
                OnPropertyChanged();
            }
        }

        public int? Duration
        {
            get { return _duration; }
            set
            {
                _duration = value;
                DurationYears = value/365d;
                OnPropertyChanged();
            }
        }

        public double? DurationYears
        {
            get { return _durationYears; }
            set
            {
                _durationYears = value;
                OnPropertyChanged();
            }
        }

        public DateTime? BuyBackDate
        {
            get { return _buyBackDate; }
            set
            {
                _buyBackDate = value;
                OnPropertyChanged();
            }
        }

        public DateTime? MatDate
        {
            get { return _matDate; }
            set
            {
                _matDate = value;
                OnPropertyChanged();
            }
        }

        public string BoardId { get; set; }

        public string ChartTip {
            get
            {
                return string.Format("Доходность: {0:N2} | Дюрация: {1:N0}", YieldClose.GetValueOrDefault(),
                    Duration.GetValueOrDefault());
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private int _i;
        public List<MarketData> MarketDatas { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            CalendarReports.SelectedDate = DateTime.Today;
            var favoritesDirectory = Path.Combine(Environment.CurrentDirectory, "Favorites");
            if (!Directory.Exists(favoritesDirectory)) return;
            foreach (var favoritesFile in Directory.GetFiles(favoritesDirectory))
                using (var reader = XmlReader.Create(favoritesFile))
                    GroupsComboBox.Items.Add(new XmlSerializer(typeof (BondsGroup)).Deserialize(reader));

            GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SecGroupsComboBox.ItemsSource = await MicexGrabber.GetSecurityCollectionsAsync((int) SecurityGroup.Bonds);
            SecGroupsComboBox.SelectedIndex = 0;
        }

        private void SecGroupsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!CalendarReports.SelectedDate.HasValue || e.AddedItems.Count == 0) return;
            BindFoundedBonds(e.AddedItems.Cast<MicexItem>().First().Id, CalendarReports.SelectedDate.Value);
        }

        private void CalendarReports_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SecGroupsComboBox.SelectedIndex < 0 || e.AddedItems.Count == 0) return;
            BindFoundedBonds(((MicexItem) SecGroupsComboBox.SelectedItem).Id, e.AddedItems.Cast<DateTime>().First());
        }

        private async void BindFoundedBonds(int bondsCollectionId, DateTime date)
        {
            ProgressImageListBox.Visibility = Visibility.Visible;
            FoundedRecordsListBox.IsEnabled = false;

            MarketDatas =
                (await MicexGrabber.GetMarketDataAsync(date, bondsCollectionId, 7, "stock", "bonds")).Where(
                    w => !w.BoardId.StartsWith("SP")).ToList();
            MarketDatas.AddRange(await MicexGrabber.GetMarketDataAsync(date, bondsCollectionId, 58, "stock", "bonds"));

            FilterFoundedBonds(SearchTextBox.Text.Split(new[] {' ', ',', ';'}, StringSplitOptions.RemoveEmptyEntries));

            ProgressImageListBox.Visibility = Visibility.Hidden;
            FoundedRecordsListBox.IsEnabled = true;
        }

        private void FilterFoundedBonds(params string[] expressions)
        {
            FoundedRecordsListBox.ItemsSource = MarketDatas.Where(w => expressions.All(expression =>
                new[] {w.SecId ?? "", w.ShortName ?? ""}.Any(
                    tm => tm.ToLowerInvariant().Contains(expression.ToLowerInvariant())))).OrderBy(o => o.ShortName);
        }

        private void CreateGroup(string s)
        {
            var bondsGroup = new BondsGroup {Name = s, BondItems = new ObservableCollection<Bond>()};
            GroupsComboBox.Items.Add(bondsGroup);
            GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            FilterFoundedBonds(SearchTextBox.Text.Split(new[] {' ', ',', ';'}, StringSplitOptions.RemoveEmptyEntries));
        }

        private async void FillDataGrid(IEnumerable<MarketData> secItems)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                    string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "" : SearchTextBox.Text + " - ",
                        ((MicexItem) SecGroupsComboBox.SelectedItem).Title,
                        " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(
                            this);

            if (GroupsComboBox.SelectedItem == null) return;

            ProgressImage.Visibility = Visibility.Visible;
            Grid1.IsEnabled = false;
            var task = new Task(() =>
            {
                Parallel.ForEach(secItems, item =>
                {
                    var currentDate = item.TradeDate ?? item.SysTime ?? DateTime.Today;

                    var marketData = MicexGrabber.GetMarketData(currentDate, item.SecId,
                        item.BoardId, "stock", "bonds");
                    if (marketData.Equals(new MarketData())) return;

                    var bond = new Bond
                    {
                        TradeDate = marketData.SysTime ?? marketData.TradeDate,
                        SecurityId = marketData.SecId,
                        SecShortName = marketData.ShortName,
                        RegNumber = marketData.RegNumber,
                        BoardName = marketData.BoardName,
                        BoardId = marketData.BoardId,
                        YieldClose = marketData.YieldClose ?? (marketData.Yield != new int() ? marketData.Yield : new int?()) ?? (marketData.CloseYield != new int() ? marketData.CloseYield : new int?()),
                        Duration = marketData.Duration != new int() ? marketData.Duration : new int?(),
                        CurrencyId = marketData.CurrencyId,
                        BuyBackDate = marketData.BuyBackDate,
                        MatDate = marketData.MatDate
                    };

                    GroupsComboBox.Dispatcher.Invoke(
                        new Action(() => { ((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Add(bond); }));
                });
            });
            task.Start();
            await task;
            ProgressImage.Visibility = Visibility.Hidden;
            Grid1.IsEnabled = true;
        }

        private void AddAll_Click(object sender, RoutedEventArgs e)
        {
            FillDataGrid(FoundedRecordsListBox.Items.Cast<MarketData>());
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            FillDataGrid(FoundedRecordsListBox.SelectedItems.Cast<MarketData>());
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            while (SelectedRecordsDataGrid.SelectedItems.Count > 0)
                ((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Remove(
                    ((Bond) SelectedRecordsDataGrid.SelectedItem));
        }

        private void RemoveAll_Click(object sender, RoutedEventArgs e)
        {
            ((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Clear();
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
                string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "" : SearchTextBox.Text + " - ",
                    ((MicexItem) SecGroupsComboBox.SelectedItem).Title,
                    " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);
        }

        private void GroupsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ((Image) FavoritesButton.Content).Source = !IsInFavorites
                ? new BitmapImage(new Uri(@"pack://application:,,,/Images/favorAdd.jpg", UriKind.RelativeOrAbsolute))
                : new BitmapImage(new Uri(@"pack://application:,,,/Images/favorRemove.jpg", UriKind.RelativeOrAbsolute));
        }

        private void DeleteGroupButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedTable = GroupsComboBox.SelectedItem as BondsGroup;
            if (selectedTable == null) return;
            var favoritesDirectory = Path.Combine(Environment.CurrentDirectory, "Favorites");
            var fileName = Path.Combine(favoritesDirectory, selectedTable.Name + ".xml");
            if (File.Exists(fileName)) File.Delete(fileName);

            GroupsComboBox.Items.Remove(GroupsComboBox.SelectedItem);
            GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
        }

        private void AddNewRowButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                    string.Concat("Группа " + ++_i,
                        " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(
                            this);

            if (GroupsComboBox.SelectedItem != null)
                ((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Add(new Bond());
        }

        private void ImportFromExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsComboBox.SelectedItem == null)
                new InputBox(
                    string.Concat("Импорт - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")),
                    CreateGroup).ShowDialog(this);

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

            FillDataGrid(
                MarketDatas.Where(w => isins.Contains(w.SecId) || isins.Contains(w.RegNumber) || isins.Contains(w.Isin))
                    .Where(
                        w =>
                            !((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Any(
                                row => row.BoardId == w.BoardId && row.SecurityId == w.SecId)).ToList());
        }

        public bool IsInFavorites
        {
            get
            {
                var selectedTable = GroupsComboBox.SelectedItem as BondsGroup;
                if (selectedTable == null) return false;

                var fileName = Path.Combine(Environment.CurrentDirectory, "Favorites", selectedTable.Name + ".xml");
                return File.Exists(fileName);
            }
        }

        private void FavoritesButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedTable = GroupsComboBox.SelectedItem as BondsGroup;
            if (selectedTable == null) return;

            var favoritesDirectory = Path.Combine(Environment.CurrentDirectory, "Favorites");
            var fileName = Path.Combine(favoritesDirectory, selectedTable.Name + ".xml");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
                ((Image) FavoritesButton.Content).Source =
                    new BitmapImage(new Uri(@"pack://application:,,,/Images/favorAdd.jpg", UriKind.RelativeOrAbsolute));
            }
            else
            {
                Directory.CreateDirectory(favoritesDirectory);

                using (XmlWriter writer = XmlWriter.Create(fileName))
                    new XmlSerializer(typeof (BondsGroup)).Serialize(writer, GroupsComboBox.SelectedItem);

                ((Image) FavoritesButton.Content).Source =
                    new BitmapImage(new Uri(@"pack://application:,,,/Images/favorRemove.jpg", UriKind.RelativeOrAbsolute));
            }
        }

        private void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            new AboutWindow {Owner = this}.ShowDialog();
        }
    }
}