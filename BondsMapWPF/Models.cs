using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Xml;

namespace BondsMapWPF
{
    [DataContract]
    public class BondsGroup
    {
        [DataMember]
        public int Id
        {
            get { return Name.GetHashCode(); }
        }

        [DataMember]
        public string Name { get; set; }

        [DataMember]
        public bool ShowOnChart { get; set; }

        [DataMember(Name = "IsFavorite")]
        private bool _isFavorite;

        public bool IsFavorite
        {
            get { return _isFavorite; }
            set
            {
                _isFavorite = value;

                var favoritesDirectory = Path.Combine(Environment.CurrentDirectory, "Favorites");
                var fileName = Path.Combine(favoritesDirectory, Name + ".xml");
                if (_isFavorite)
                {
                    Directory.CreateDirectory(favoritesDirectory);
                    using (var writer = XmlWriter.Create(fileName))
                        new DataContractSerializer(typeof(BondsGroup)).WriteObject(writer, this);
                }
                else
                {
                    if (File.Exists(fileName))
                        File.Delete(fileName);
                }
            }
        }

        [DataMember]
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
                DurationYears = value / 365d;
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

        public string ChartTip
        {
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
}
