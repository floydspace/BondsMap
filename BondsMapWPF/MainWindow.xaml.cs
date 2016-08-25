using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using System.Runtime.Serialization;
using System.Windows.Media.Animation;
using FloydSpace.MICEX.Portable;
using FloydSpace.MICEX.Portable.Maps;
using Microsoft.Win32;

namespace BondsMapWPF
{
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
            foreach (var favoritesFile in Directory.GetFiles(favoritesDirectory, "*.xml"))
                using (var reader = XmlReader.Create(favoritesFile))
                    GroupsComboBox.Items.Add(new DataContractSerializer(typeof (BondsGroup)).ReadObject(reader));

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
            var animation = (Storyboard)TryFindResource("ProgressStoryboardListBox");
            animation.Begin();

            MarketDatas =
                (await MicexGrabber.GetMarketDataAsync(date, bondsCollectionId, 7, "stock", "bonds")).Where(
                    w => !w.BoardId.StartsWith("SP")).ToList();
            MarketDatas.AddRange(await MicexGrabber.GetMarketDataAsync(date, bondsCollectionId, 58, "stock", "bonds"));

            FilterFoundedBonds(SearchTextBox.Text.Split(new[] {' ', ',', ';'}, StringSplitOptions.RemoveEmptyEntries));

            ProgressImageListBox.Visibility = Visibility.Hidden;
            FoundedRecordsListBox.IsEnabled = true;
            animation.Stop();
        }

        private void FilterFoundedBonds(params string[] expressions)
        {
            FoundedRecordsListBox.ItemsSource = MarketDatas.Where(w => expressions.All(expression =>
                new[] {w.SecId ?? "", w.ShortName ?? ""}.Any(
                    tm => tm.ToLowerInvariant().Contains(expression.ToLowerInvariant())))).OrderBy(o => o.ShortName);
        }

        private bool CreateGroup(string s)
        {
            var bondsGroup = new BondsGroup
            {
                Name = s,
                BondItems = new ObservableCollection<Bond>(),
                IsFavorite = false,
                ShowOnChart = true
            };
            if ((GroupsComboBox.Items.Cast<BondsGroup>()).Select(item => item.Id).Contains(bondsGroup.Id))
                return false;
            GroupsComboBox.Items.Add(bondsGroup);
            GroupsComboBox.SelectedIndex = GroupsComboBox.Items.Count - 1;
            return true;
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
            var animation = (Storyboard)TryFindResource("ProgressStoryboard");
            animation.Begin();
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
                        YieldClose =
                            marketData.YieldClose ??
                            (marketData.Yield != new int() ? marketData.Yield : new int?()) ??
                            (marketData.CloseYield != new int() ? marketData.CloseYield : new int?()),
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
            animation.Stop();
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


        private void CreateGroupButton_Click(object sender, RoutedEventArgs e)
        {
            new InputBox(
                string.Concat(string.IsNullOrWhiteSpace(SearchTextBox.Text) ? "" : SearchTextBox.Text + " - ",
                    ((MicexItem) SecGroupsComboBox.SelectedItem).Title,
                    " - ", CalendarReports.SelectedDate.GetValueOrDefault().ToString("d")), CreateGroup).ShowDialog(this);
        }

        private void DeleteGroupButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedTable = GroupsComboBox.SelectedItem as BondsGroup;
            if (selectedTable == null) return;

            selectedTable.IsFavorite = false;
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
                    .Where(w => !((BondsGroup) GroupsComboBox.SelectedItem).BondItems.Any(
                        row => row.BoardId == w.BoardId && row.SecurityId == w.SecId)).ToList());
        }

        private void AboutButton_Click(object sender, RoutedEventArgs e)
        {
            new AboutWindow {Owner = this}.ShowDialog();
        }

        private void BuildBondsMapButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedRecordsTables = GroupsComboBox.Items.Cast<BondsGroup>().Where(w => w.ShowOnChart).ToArray();
            if (selectedRecordsTables.SelectMany(s => s.BondItems)
                .All(w => !w.Duration.HasValue || !w.YieldClose.HasValue)) return;
            new ChartWindow(selectedRecordsTables).Show();
        }
    }
}