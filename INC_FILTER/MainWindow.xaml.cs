using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Flurl.Http;
using System.Net;
using System.Diagnostics;
using System.Windows.Threading;
using Newtonsoft.Json;
using System.IO;
using System.Threading;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using Microsoft.Office.Core;

namespace INC_FILTER
{
    public partial class MainWindow
    {
        private Dictionary<string, List<IncEmailItem>> incEmails = new Dictionary<string, List<IncEmailItem>>();
        private Dictionary<string, List<IncEmailItem>> otherEmails = new Dictionary<string, List<IncEmailItem>>();

        private List<Card> trelloCards;
        private List<TrelloList> trelloList;
        private Dictionary<string, string> trelloLabel;
        private Settings settings;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
            incName.SelectionChanged += IncName_SelectionChanged;
            mailLists.MouseDoubleClick += MailLists_MouseDoubleClick;
        }

        private void MailLists_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OpenMOutlook_Click(null, null);
        }

        private void IncName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (e.AddedItems.Count == 0) {
                    mailLists.ItemsSource = new object[0];
                    return;
                };
                var s = ((dynamic)e.AddedItems[0]).IncName;
                if (showOtherCheckBox.IsChecked.Value)
                {

                    if (otherEmails.ContainsKey(s))
                        mailLists.ItemsSource = otherEmails[s];
                    return;
                }
                if (e.AddedItems.Count > 0)
                {
                    mailLists.ItemsSource = incEmails[s];

                    var card = (incName.SelectedItem as dynamic).TrelloCard as Card;
                    if (card != null)
                    {
                        lastMailTimePicker.SelectedDate = card.lastMailAt;
                        remind1stTimePicker.SelectedDate = card.remind1st;
                        remind2ndTimePicker.SelectedDate = card.remind2nd;
                        remind3rdTimePicker.SelectedDate = card.remind3rd;
                        followUpTimePicker.SelectedDate = card.followUp;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void UiRenderEmailListView()
        {
            this.Dispatcher.Invoke(() =>
            {
                loadingGrid.Visibility = Visibility.Visible;
                contentGrid.Visibility = Visibility.Collapsed;

                var text = searchText.Text.Trim().ToLower();
                if (showOtherCheckBox.IsChecked.Value)
                {
                    incName.ItemsSource = this.otherEmails.Keys.Select(x => new
                    {
                        LastestMail = this.otherEmails[x].FirstOrDefault(),
                        IncName = x,
                        TrelloCard = trelloCards.Where(c => c.name.Contains(x)).FirstOrDefault()
                    })
                    .Where(x =>
                    {
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            return x.IncName.ToUpper().Contains(text.ToUpper().Trim());
                        }
                        return true;

                    }).OrderByDescending(x => x.LastestMail.ReceivedTime).ToList();

                    summaryTextBlock.Text = "Total INC: " + incName.Items.Count + " items";

                    loadingGrid.Visibility = Visibility.Collapsed;
                    contentGrid.Visibility = Visibility.Visible;

                    return;
                }

                var noFilter = noFilterCheckBox.IsChecked ?? false;
                var filterNoTrelloCard = noTrelloCheckBox.IsChecked ?? false;
              
                filterTrello.ItemsSource = trelloList;
                var ids = trelloList.Where(f => f.IsChecked).Select(f => f.id).ToList();
                incName.ItemsSource = this.incEmails.Keys.Select(x => new
                {
                    LastestMail = this.incEmails[x].FirstOrDefault(),
                    IncName = x,
                    TrelloCard = trelloCards.Where(c => c.name.Contains(x)).FirstOrDefault()
                })
                .Where(x =>
                {
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return x.IncName.ToUpper().Contains(text.ToUpper().Trim());
                    }

                    if (noFilter) return true;

                    return (filterNoTrelloCard && x.TrelloCard == null) || (x.TrelloCard != null && ids.Any(l => x.TrelloCard.idList.Contains(l)));

                }).OrderByDescending(x => x.LastestMail.ReceivedTime).ToList();

                summaryTextBlock.Text = "Total INC: " + incName.Items.Count + " items";

                loadingGrid.Visibility = Visibility.Collapsed;
                contentGrid.Visibility = Visibility.Visible;

            var stats = trelloCards.Where(x => x.Base64CodeData != null)
                .Select(x => x.Base64CodeData.status).GroupBy(x => x).Select(x => new { x.Key, Count = x.Count() });

              var stats2 = trelloCards.Where(x => x.Base64CodeData != null && settings.MyXFptMembers.Any(y => x.Base64CodeData.assignee.ToLower().Contains(y.ToLower())))
                .Select(x => x.Base64CodeData.status).GroupBy(x => x).Select(x => new { x.Key, Count = x.Count() });

              var all =  stats.Select(x => x.Key + ": " + (stats2.FirstOrDefault(t => t.Key == x.Key)?.Count ?? 0) + "/" + x.Count);

                this.Title = DateTime.Now.ToShortTimeString() + " | X-INCIDENT TOOL | " + "" + string.Join(", ", all);
            });
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {

            var t = new DispatcherTimer(new TimeSpan(0, 0, 0, 0, 300), DispatcherPriority.Background, t_Tick, Dispatcher.CurrentDispatcher) { IsEnabled = true };
            t.Start();
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            Task.Run(async () =>
            {
                try
                {
                    settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText("settings.json"));
                    await LoadTrello();
                    LoadIncEmail();

                    UiRenderEmailListView();

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }).ConfigureAwait(true);

        }

        int counter = 0;

        private void t_Tick(object sender, EventArgs e)
        {
            counter++;
            string loadingText = "LOADING...";
            this.loadinText.Text = loadingText.Insert((counter % 10) + 1, " ");
        }

        DateTime d = DateTime.Now;
        private void Log(String x) => Dispatcher.Invoke(() => loadedComponent.Text = loadedComponent.Text + "\r\n" + (DateTime.Now - d).ToString(@"mm\:ss\.fff") + " : " + x);

        private async Task LoadTrello()
        {
            Log("Load Trello card");
            trelloCards = (await "https://api.trello.com/1/boards/replace-hardcode-board-ic?keyreplace-hardcode-key&token=replace-hardcode-token&cards=all&card_customFieldItems=true"
                .GetJsonAsync<TrelloData>()).cards.ToList();

            Log("Load Trello list");
            trelloList = (await "https://api.trello.com/1/boards/replace-hardcode-board-ic/lists?keyreplace-hardcode-key&token=replace-hardcode-token"
                .GetJsonAsync<List<TrelloList>>());

            Log("Load Trello label");
            trelloLabel = (await "https://api.trello.com/1/boards/replace-hardcode-board-ic/labels?keyreplace-hardcode-key&token=replace-hardcode-token"
                .GetJsonAsync<List<TrelloLabel>>()).ToDictionary(x => x.id, x => x.name);

            var dict = trelloList.ToDictionary(x => x.id, x => x.name);

            Log("Proceed Trello Data: " + trelloCards.Count + " cards");
            foreach (var card in trelloCards)
            {
                card.listName = dict[card.idList];
                card.labelDisplay = string.Join(" | ", card.labels.Select(x => x.name));

                try
                {
                    var base64Text = card.customFieldItems.Where(t => t.idCustomField == "5eb3fc7e6cfebd31477e4ac5").FirstOrDefault()?.value.text;
                    if (!string.IsNullOrWhiteSpace(base64Text))
                    {
                        string base64Decoded;
                        byte[] data = System.Convert.FromBase64String(base64Text);
                        base64Decoded = System.Text.ASCIIEncoding.ASCII.GetString(data);
                        card.Base64CodeData = JsonConvert.DeserializeObject<Base64CodeData>(base64Decoded);
                    }

                    card.lastMailAt = card.customFieldItems.Where(t => t.idCustomField == "5ec49c1ec38f8c68ddc270be").FirstOrDefault()?.value.date;
                    card.remind1st = card.customFieldItems.Where(t => t.idCustomField == "5ecf24c0a327d374b4dca89f").FirstOrDefault()?.value.date;
                    card.remind2nd = card.customFieldItems.Where(t => t.idCustomField == "5ecf24cb32402064b8c00b67").FirstOrDefault()?.value.date;
                    card.remind3rd = card.customFieldItems.Where(t => t.idCustomField == "5ecf24d57c11cd0e5f705772").FirstOrDefault()?.value.date;
                    card.followUp = card.customFieldItems.Where(t => t.idCustomField == "5eb3f6b4c933c517d89d0816").FirstOrDefault()?.value.date;
                    
                    if (card.Base64CodeData != null)
                    {
                        card.submitDate = UnixTimeStampToDateTime(card.Base64CodeData.submitDate);
                        card.assignee = card.Base64CodeData.assignee;
                        card.status = card.Base64CodeData.status;
                    }
                }
                catch (System.Exception ex)
                {
                    Log("parse base 64 err for card " + card.name);
                }
            }

           

        }

        private MAPIFolder closedFolder = null;
        private Microsoft.Office.Interop.Outlook.Application oApp;
        private Microsoft.Office.Interop.Outlook.NameSpace oNS;
        private Stores stores;

        private void LoadIncEmail()
        {
            Log("Begin Load IncEmail from Local Machine");
            var dict = new Dictionary<string, List<IncEmailItem>>();
            var otherEmail = new Dictionary<string, List<IncEmailItem>>();

            Func<string, IncEmailItem, List<string>> ExtractIncidentName = (s, i) =>
            {
                s = s.Trim();
                var matches1 = Regex.Matches(s, @"ICT_INC\d+");
                var matches2 = Regex.Matches(s, @"ICT_WO\d+");
                var matchesValues = matches1.Cast<Match>().Where(x => x.Success).Select(x => x.Value).Union(matches2.Cast<Match>().Where(x => x.Success).Select(x => x.Value)).ToList();
                foreach (var matchItem in matchesValues)
                {
                    if (!dict.ContainsKey(matchItem))
                    {
                        dict.Add(matchItem, new List<IncEmailItem>());
                    }

                    dict[matchItem].Add(i);
                }

                if (!matchesValues.Any())
                {
                    var upperTitle = s.ToUpper();
                    if (upperTitle.StartsWith("FW:") || upperTitle.StartsWith("RE:") || upperTitle.StartsWith("CT:"))
                    {
                        if (upperTitle.Length > 3)
                        {
                            s = s.Substring(3).Trim();
                        }
                    }
                    if (!otherEmail.ContainsKey(s))
                    {
                        otherEmail.Add(s, new List<IncEmailItem>());
                    }

                    otherEmail[s].Add(i);
                }

                return matchesValues;
            };


            oApp = new Microsoft.Office.Interop.Outlook.Application();
            oNS = oApp.GetNamespace("mapi");
            stores = oNS.Stores;
            var folders = settings.ScanFolders;

            foreach (Store store in stores)
            {
                // continue;
                Log("Read Store: " + store.DisplayName);
                MAPIFolder YOURFOLDERNAME = store.GetRootFolder();
                Log("MAPIFolder: " + YOURFOLDERNAME.Name);
                foreach (MAPIFolder subF in YOURFOLDERNAME.Folders)
                {
                    var f = subF.Name;
                    Log("MAPIFolder: " + f + "#");
                    if (folders.Contains(f))
                    {
                        ScanFolder(ExtractIncidentName, subF);
                    }
                    else
                    {
                        // Log("SKip read Mail in folder " + subF.Name);
                    }
                }
            }

            Log("Proceed mail...");
            foreach (var item in dict)
            {
                this.incEmails.Add(item.Key, item.Value.OrderByDescending(x => x.ReceivedTime).ToList());
            }

            foreach (var item in otherEmail)
            {
                this.otherEmails.Add(item.Key, item.Value.OrderByDescending(x => x.ReceivedTime).ToList());
            }

            Log("Log off outlook...");
            //Log off.
            oNS.Logoff();

        }

        private void ScanFolder(Func<string, IncEmailItem, List<string>> ExtractIncidentName, MAPIFolder subF, string folderPath=null)
        {
            Log("Read Mail in folder " + subF.Name);
            var mailItems = subF.Items.Cast<object>()
                .Where(x => x is Microsoft.Office.Interop.Outlook.MailItem)
                .Cast<Microsoft.Office.Interop.Outlook.MailItem>()
                .Where(x => x.Subject != null && x.SenderName != "hard code email here")
                .OrderByDescending(x => x.ReceivedTime);


            foreach (var item in mailItems)
            {
                string body = item.Body ?? string.Empty;
                body = body.Split(new[] { "From: " }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;
                body = (body.Length > 500 ? body.Substring(0, 500) : body);
                body = Regex.Replace(body, @"\t|\n|\r", " ");
                ExtractIncidentName(item.Subject, new IncEmailItem
                {
                    SenderName = item.SenderName,
                    To = item.To,
                    CC = item.CC,
                    PreviewBody = body,
                    Subject = item.Subject,
                    ReceivedTime = item.SentOn,
                    MailItem = item,
                    FolderPath = subF.FolderPath,
                    ShowModal = () => item.Display(false),
                    ShowReply = () => item.ReplyAll()
                });
            }

            if (settings.ScanChildFolder)
            {
                foreach(var childFolder in subF.Folders.Cast<MAPIFolder>())
                {
                    ScanFolder(ExtractIncidentName, childFolder);
                }
            }

        }

        private void Search_TextChanged(object sender, TextChangedEventArgs e) => UiRenderEmailListView();
        private void FilterClicked(object sender, RoutedEventArgs e) => UiRenderEmailListView();

        private void OpenMOutlook_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (mailLists.SelectedItem is IncEmailItem)
                {
                    (mailLists.SelectedItem as IncEmailItem).ShowModal();
                }
                else
                {
                    MessageBox.Show("Pls choose mail item.");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        
        private async void UpdateTrello_Verified(object sender, RoutedEventArgs e)
        {
            if (incName.SelectedItem == null)
            {
                MessageBox.Show("Pls choose Inc item.");
                return;
            }

            var card = (incName.SelectedItem as dynamic).TrelloCard as Card;
            if (card == null)
            {
                MessageBox.Show("No Trello card data.");
                return;
            }

            try
            {
                var data = new { value = "5ea065017669b2254965a046" };

                var str = JsonConvert.SerializeObject(data);
                await $"https://api.trello.com/1/cards/{card.id}/idLabels?keyreplace-hardcode-key&token=replace-hardcode-token"
                    .PostJsonAsync(data);

                ((incName.SelectedItem as dynamic).TrelloCard as Card).labels.Add(new Label() { id = "5ea065017669b2254965a046", name = "USER-VERIFIED" });
                ((incName.SelectedItem as dynamic).TrelloCard as Card).RaiseChange();

                MessageBox.Show("Update label done.");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        private void CopyIncNameToClipboard_Click(object sender, RoutedEventArgs e)
        {
            if (incName.SelectedItem != null)
            {
                Clipboard.SetText((incName.SelectedItem as dynamic).IncName);
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private Card GetSelectedCard()
        {
            return (incName.SelectedItem as dynamic)?.TrelloCard as Card;
        }

        private void CopyCustomerEmailToClipboard_Click(object sender, RoutedEventArgs e)
        {
            if (incName.SelectedItem != null)
            {
                var card = GetSelectedCard();
                Clipboard.SetText(card?.Base64CodeData?.customerEmail);
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private void CopyCustomerNameToClipboard_Click(object sender, RoutedEventArgs e)
        {
            if (incName.SelectedItem != null)
            {
                var card = GetSelectedCard();
                Clipboard.SetText(card?.Base64CodeData?.customerName);
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private void CopyRemindEmailToClipboard_Click(object sender, RoutedEventArgs e)
        {
            if (mailLists.SelectedItem != null)
            {
                var card = GetSelectedCard();

                if (card == null || card.Base64CodeData == null)
                {
                    MessageBox.Show("Card not found");
                    return;
                }

                var link = $"https://x-url/smartit/app/#/{(card.Base64CodeData.displayId.StartsWith("ICT_INC") ? "incident" : "workorder")}/{card?.Base64CodeData?.id.ToUpper()}";

                string msg = $@"Dear {card.Base64CodeData?.customerName},

May I know your incident status. Is it resolved?
Your incident id: {card?.Base64CodeData?.displayId}
Your incident link: {link}

Can u revert to us the result.

Thanks,
Diep";

                // (mailLists.SelectedItem as IncEmailItem).ShowModal();


                Clipboard.SetText(msg);
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private void OpenInTrello_Click(object sender, RoutedEventArgs e)
        {
            if (incName.SelectedItem != null)
            {
                var card = (incName.SelectedItem as dynamic).TrelloCard as Card;
                if (card == null)
                {
                    MessageBox.Show("No Trello card.");
                    return;
                }

                Process.Start("open-trello.bat", card.url);
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private void OpenInX_Click(object sender, RoutedEventArgs e)
        {
            if (incName.SelectedItem != null)
            {
                var card = (incName.SelectedItem as dynamic).TrelloCard as Card;
                if (card == null || card.Base64CodeData == null)
                {
                    MessageBox.Show("No Trello card data.");
                    return;
                }

                Process.Start("open-myX.bat", $"https://x-url/smartit/app/#/{(card.Base64CodeData.displayId.StartsWith("ICT_INC") ? "incident" : "workorder")}/{card.Base64CodeData.id.ToUpper()}");
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private async void updateLastEmailDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateField("5ec49c1ec38f8c68ddc270be", lastMailTimePicker.SelectedDate, lastMailTimePicker, "lastMailAt");
        }

        private async void updateRemind1stDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateField("5ecf24c0a327d374b4dca89f", remind1stTimePicker.SelectedDate, remind1stTimePicker, "remind1st");
        }

        private async void updateRemind2ndDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateField("5ecf24cb32402064b8c00b67", remind2ndTimePicker.SelectedDate, remind2ndTimePicker, "remind2nd");
        }

        private async void updateRemind3rdDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateField("5ecf24d57c11cd0e5f705772", remind3rdTimePicker.SelectedDate, remind3rdTimePicker, "remind3rd");
        }
        private async void updateFollowUpDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateField("5eb3f6b4c933c517d89d0816", followUpTimePicker.SelectedDate, followUpTimePicker, "followUp");
        }

        private async void autoLastEmailDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ec49c1ec38f8c68ddc270be", (mailLists.SelectedItem as IncEmailItem)?.ReceivedTime, lastMailTimePicker, "lastMailAt");
        }

        private async void updateNowEmailDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ec49c1ec38f8c68ddc270be", DateTime.Now, lastMailTimePicker, "lastMailAt");
        }

        private async void autoupdateRemind1stDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ecf24c0a327d374b4dca89f", (mailLists.SelectedItem as IncEmailItem)?.ReceivedTime, remind1stTimePicker, "remind1st");
        }

        private async void updateNowRemind1stDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ecf24c0a327d374b4dca89f", DateTime.Now, remind1stTimePicker, "remind1st");
        }


        private async void autoupdateRemind2ndDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ecf24cb32402064b8c00b67", (mailLists.SelectedItem as IncEmailItem)?.ReceivedTime, remind2ndTimePicker, "remind2nd");
        }

        private async void updateNowRemind2ndDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ecf24cb32402064b8c00b67", DateTime.Now, remind2ndTimePicker, "remind2nd");
        }

        private async void autoupdateRemind3rdDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ecf24d57c11cd0e5f705772", (mailLists.SelectedItem as IncEmailItem)?.ReceivedTime, remind3rdTimePicker, "remind3rd");
        }

        private async void updateNowRemind3rdDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5ecf24d57c11cd0e5f705772", DateTime.Now, remind3rdTimePicker, "remind3rd");
        }

        private async void autoupdateFollowUpDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5eb3f6b4c933c517d89d0816", (mailLists.SelectedItem as IncEmailItem)?.ReceivedTime, followUpTimePicker, "followUp");
        }
        private async void updateNowFollowUpDateinTrello_Click(object sender, RoutedEventArgs e)
        {
            await UpdateDateFieldBySelectedMail("5eb3f6b4c933c517d89d0816", DateTime.Now, followUpTimePicker, "followUp");
        }

        private async Task UpdateDateFieldBySelectedMail(string fieldId, DateTime? value, DatePicker picker, string fieldName)
        {
            try
            {
                await UpdateDateField(fieldId, value, picker, fieldName);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private async Task UpdateDateField(string fieldId, DateTime? value, DatePicker picker, string fieldName)
        {
            if (incName.SelectedItem == null)
            {
                MessageBox.Show("Pls choose Inc item.");
                return;
            }

            var card = (incName.SelectedItem as dynamic).TrelloCard as Card;
            if (card == null)
            {
                MessageBox.Show("No Trello card data.");
                return;
            }

            try
            {
                value = value?.ToUniversalTime();

                var data = new { value = new { date = value } };
                await $"https://api.trello.com/1/cards/{card.id}/customField/{fieldId}/item?keyreplace-hardcode-key&token=replace-hardcode-token"
                    .PutJsonAsync(data);

                picker.SelectedDate = value;

                card.GetType().GetProperty(fieldName).SetValue(card, value);
                card.RaiseChange();

                await this.ShowMessageAsync("Sync success to Trello", "Field Name: " + fieldName + "\r\nValue: " + value?.ToLocalTime().ToLongDateString() + " " + value?.ToLocalTime().ToLongDateString());
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        MAPIFolder GetCloseFolder()
        {
            try
            {
                if (closedFolder != null) return closedFolder;
                var allFolders = new Queue<MAPIFolder>(stores.Cast<Store>().SelectMany(x => x.GetRootFolder().Folders.Cast<MAPIFolder>()));
                var folderName = settings.ClosedFolderName;

                while (allFolders.Any())
                {
                    var folderX = allFolders.Dequeue();

                    if (folderX.Name == folderName)
                    {
                        closedFolder = folderX;
                        break;
                    };

                    foreach (var explorerFolder in folderX.Folders.Cast<MAPIFolder>())
                    {
                        allFolders.Enqueue(explorerFolder);
                    }
                }

                return closedFolder;
            } catch(System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private void SaveEmailList_Click(object sender, RoutedEventArgs e)
        {
            if (incName.SelectedItem != null)
            {
                string inc = (incName.SelectedItem as dynamic).IncName;
                var mails = incEmails[inc];
                string folder = Path.Combine(inc_mail_save_folder.Text, inc);

                Directory.CreateDirectory(folder);
                foreach (var mail in mails)
                {
                    try
                    {
                        var fileName = RemoveInvalidChars(mail.ReceivedTime.ToString("dd MMM yy HH-mm") + "__" + mail.SenderName + "__" + mail.Subject + ".msg");
                        mail.MailItem.InternetCodepage = 65001;
                        mail.MailItem.SaveAs(Path.Combine(folder, fileName));
                    }
                    catch (System.Exception ex) {
                        MessageBox.Show(inc + "\t" + ex.Message);
                    }
                }

                MessageBox.Show("Done");
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private void MoveListEmailByInc(object sender, RoutedEventArgs e)
        {
            GetCloseFolder();

            if (incName.SelectedItem != null)
            {
                string inc = (incName.SelectedItem as dynamic).IncName;
                var mails = incEmails[inc];

                try
                {
                    if (closedFolder != null)
                    {
                        var subFolder = closedFolder.Folders.Cast<MAPIFolder>().FirstOrDefault(x => x.Name == inc);

                        if (subFolder == null)
                        {
                            subFolder = closedFolder.Folders.Add(inc, Type.Missing) as MAPIFolder;
                        }

                        foreach (var mail in mails)
                        {
                            mail.MailItem.Move(subFolder);
                        }
                      
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(inc + "\t" + ex.Message);
                }

                MessageBox.Show("Moved to Folder : " + inc + ":" + mails.Count);
            }
            else
            {
                MessageBox.Show("Pls choose Inc item.");
            }
        }

        private void MoveAllEmailList_Click(object sender, RoutedEventArgs e)
        {
            GetCloseFolder();

            if (closedFolder == null)
            {
                MessageBox.Show("CLOSE FOLDER NOT FOUND");
                return;
            }

            if (MessageBox.Show("WARNING! Do you want to sort all these email?", "WARNING", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
            {
                MessageBox.Show("You cancelled it");
                return;
            }

            var itemsource = incName.ItemsSource.Cast<dynamic>().Select(x => x.IncName).Cast<string>();
            List<string> err = new List<string>();
            foreach (var inc in itemsource)
            {
                var mails = incEmails[inc];

                try
                {

                    var subFolder = closedFolder.Folders.Cast<MAPIFolder>().FirstOrDefault(x => x.Name == inc);

                    if (subFolder == null)
                    {
                        subFolder = closedFolder.Folders.Add(inc, Type.Missing) as MAPIFolder;
                    }

                    foreach (var mail in mails)
                    {
                        mail.MailItem.Move(subFolder);
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(inc + "\t" + ex.Message);
                }
            }

            MessageBox.Show("Finish, fail mesage = " + string.Join("\r\n", err));
        }

        public string RemoveInvalidChars(string filename)
        {
            var t = string.Concat(filename.Split(Path.GetInvalidFileNameChars()));

            if (t.Length > 250)
            {
                t = t.Substring(0, 250);
            }

            return t;
        }

        private void SaveAllEmailList_Click(object sender, RoutedEventArgs e)
        {
            var itemsource = incName.ItemsSource.Cast<dynamic>().Select(x => x.IncName).Cast<string>();
            List<string> err = new List<string>();
            foreach (var inc in itemsource)
            {
                var mails = incEmails[inc];
                string folder = Path.Combine(inc_mail_save_folder.Text, inc);

                Directory.CreateDirectory(folder);
                foreach (var mail in mails)
                {
                    try
                    {
                        var fileName = RemoveInvalidChars(mail.ReceivedTime.ToString("dd MMM yy HH-mm") + "__" + mail.SenderName + "__" + mail.Subject + ".msg");
                        mail.MailItem.SaveAs(Path.Combine(folder, fileName));
                    }
                    catch (System.Exception ex) { err.Add(inc + "\t" + ex.Message); }
                }
            }

            MessageBox.Show("Finish, fail mesage = " + string.Join("\r\n", err));
        }

        private void IncHasNoEmail_Click(object sender, RoutedEventArgs e)
        {
            var emails = incEmails.Keys.ToArray();
            var list = trelloCards
                .Where(x => x.Base64CodeData != null)
                .Where(x => !emails.Any(t => x.name.Contains(t)) && (x.Base64CodeData.status == "In Progress" || x.Base64CodeData.status == "Pending"))
                .OrderByDescending(x => x.Base64CodeData.submitDate)
                .Select(x => UnixTimeStampToDateTime(x.Base64CodeData.submitDate).ToString("dd MMM yyyy HH:mm") + " | " +  x.name + " | " + x.Base64CodeData.status).ToList();

            TrelloHasNoEmail w = new TrelloHasNoEmail(list);
            w.ShowDialog();
        }

        public static DateTime UnixTimeStampToDateTime(double unixTimeStamp)
        {
            System.DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddMilliseconds(unixTimeStamp).ToLocalTime();
            return dtDateTime;
        }

        private void RestartWindow(object sender, RoutedEventArgs e)
        {
            Process.Start("restart.bat");

        }

        private void OpenSettings(object sender, RoutedEventArgs e)
        {
            Process.Start("open-settings.bat");
        }
    }



    public class IncEmailItem
    {
        public string SenderName { get; set; }
        public string To { get; set; }
        public string CC { get; set; }
        public string PreviewBody { get; set; }
        public DateTime ReceivedTime { get; set; }
        public string Subject { get; set; }

        public MailItem MailItem { get; set; }

        public System.Action ShowModal { get; set; }
        public System.Action ShowReply { get; set; }

        public string FolderPath { get; set; }

    }

    public class Settings
    {
        public string ClosedFolderName { get; set; }

        public bool ScanChildFolder { get; set; }

        public string[] ScanFolders { get; set; }

        public HighlightValue[] HighlightSenders { get; set; }
        public HighlightValue[] HighlightTrelloStatus { get; set; }
        public string[] MyXFptMembers { get; set; }
        public class HighlightValue
        {
            public byte[] Color { get; set; }

            public string[] Values { get; set; }
        }
    }
}
