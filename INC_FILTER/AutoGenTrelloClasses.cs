using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace INC_FILTER
{
    public class Trello
    {
        public int board { get; set; }
        public int card { get; set; }
    }

    public class Label
    {
        public string id { get; set; }
        public string name { get; set; }
    }

    public class Value
    {
        public string text { get; set; }
        public DateTime? date { get; set; }
    }

    public class CustomFieldItem
    {
        public string id { get; set; }
        public Value value { get; set; }
        public string idCustomField { get; set; }
        public string idModel { get; set; }
        public string modelType { get; set; }
    }

    public class Card : INotifyPropertyChanged
    {
        public DateTime submitDate { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        public void RaiseChange()
        {
            labelDisplay = string.Join(" | ", labels.Select(x => x.name));

            OnPropertyChanged("labelDisplay");
            OnPropertyChanged("lastMailAt");
            OnPropertyChanged("remind1st");
            OnPropertyChanged("remind2nd");
            OnPropertyChanged("remind3rd");
            OnPropertyChanged("followUp");

        }

        public string id { get; set; }
        public DateTime dateLastActivity { get; set; }
        public string desc { get; set; }
        public string idList { get; set; }
        public string listName { get; set; }
        public IList<string> idLabels { get; set; }
        public string name { get; set; }
        public IList<Label> labels { get; set; }
        public string labelDisplay { get; set; }
        public string url { get; set; }
        public IList<CustomFieldItem> customFieldItems { get; set; }

        public Base64CodeData Base64CodeData { get;set;}

        public DateTime? lastMailAt { get; set; }
        public DateTime? remind1st { get; set; }
        public DateTime? remind2nd { get; set; }
        public DateTime? remind3rd { get; set; }

        public DateTime? followUp { get; set; }
        public string status { get;  set; }
        public string assignee { get; set; }
    }

    public class TrelloData
    {
        public IList<Card> cards { get; set; }
    }

    public class TrelloList : INotifyPropertyChanged
    {
        private bool isChecked;

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        public string id { get; set; }
        public string name { get; set; }

        public bool IsChecked
        {
            get => isChecked; set
            {
                isChecked = value;
                OnPropertyChanged();
            }
        }
    }

    public class Base64CodeData
    {
        public string displayId { get; set; }
        public string priority { get; set; }
        public string id { get; set; }
        public string status { get; set; }
        public string type { get; set; }
        public string customerName { get; set; }
        public string customerEmail { get; set; }
        public string assignee { get; set; }
        public long submitDate { get; set; }
        public long modifiedDate { get; set; }
    }

    public class TrelloLabel
    {
        public string id { get; set; }
        public string idBoard { get; set; }
        public string name { get; set; }
        public string color { get; set; }
    }

}
