using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace GraphSDKDemo
{
    public sealed partial class MessagesPage : Page
    {
        GraphServiceClient graphClient = null;

        MailFolder inbox = null;
        ObservableCollection<Models.Message> MyMessages = null;

        Message myMessage = null;
        Models.Message selectedMessage = null;

        public MessagesPage()
        {
            this.InitializeComponent();
        }

        private async void GetMessagesButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {
                inbox = await graphClient.Me.MailFolders.Inbox.Request().GetAsync();
                IUserMessagesCollectionPage messages =
                    await graphClient.Me.Messages.Request().Top(20)
                                                 .Select("sender,from, subject, importance").GetAsync();

                MyMessages = new ObservableCollection<Models.Message>();

                foreach (var message in messages)
                {
                    MyMessages.Add(new Models.Message
                    {
                        Id = message.Id,
                        Sender = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Name :
                                  "Unknown name",
                        From = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Address :
                                  "Unknown email",
                        Subject = message.Subject ?? "No subject",
                        Importance = message.Importance.ToString()
                    });
                }

                MessageCountTextBlock.Text = $"You have {inbox.TotalItemCount} messages, " +
                    $"{inbox.UnreadItemCount} of them are unread. Here are the first 20:";
                MessagesListView.ItemsSource = MyMessages;
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not get messages: {ex.Error.Message}";
            }
        }

        private async void GetInboxMessagesButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();
            string inboxId = string.Empty;

            try
            {
                inbox = await graphClient.Me.MailFolders.Inbox.Request().GetAsync();
                IMailFolderMessagesCollectionPage messages =
                    await graphClient.Me.MailFolders.Inbox.Messages.Request().Top(20).GetAsync();

                MyMessages = new ObservableCollection<Models.Message>();

                foreach (var message in messages)
                {
                    MyMessages.Add(new Models.Message
                    {
                        Id = message.Id,
                        Sender = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Name :
                                  "Unknown name",
                        From = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Address :
                                  "Unknown email",
                        Subject = message.Subject ?? "No subject",
                        Importance = message.Importance.ToString()
                    });
                }

                MessageCountTextBlock.Text = $"You have {inbox.TotalItemCount} messages, " +
                    $"{inbox.UnreadItemCount} of them are unread. Here are the first 20:";
                MessagesListView.ItemsSource = MyMessages;
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not get messages: {ex.Error.Message}";
            }
        }

        private async void GetHighImportanceMessagesButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();
            string inboxId = string.Empty;

            try
            {
                IMailFolderMessagesCollectionPage messages =
                    await graphClient.Me.MailFolders.Inbox.Messages.Request().Filter("importance eq 'high'").GetAsync();

                MyMessages = new ObservableCollection<Models.Message>();

                foreach (var message in messages)
                {
                    MyMessages.Add(new Models.Message
                    {
                        Id = message.Id,
                        Sender = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Name :
                                  "Unknown name",
                        From = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Address :
                                  "Unknown email",
                        Subject = message.Subject ?? "No subject",
                        Importance = message.Importance.ToString()
                    });
                }

                MessageCountTextBlock.Text = $"You have {MyMessages.Count()} red bang messages:";
                MessagesListView.ItemsSource = MyMessages;
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not get messages: {ex.Error.Message}";
            }
        }

        private async void GetRogreenHighImportanceMessagesButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();
            string inboxId = string.Empty;

            try
            {
                IMailFolderMessagesCollectionPage messages =
                    await graphClient.Me.MailFolders.Inbox.Messages.Request().
                    Filter("importance eq 'high' and sender/emailaddress/address eq 'rogreen@microsoft.com'").GetAsync();

                MyMessages = new ObservableCollection<Models.Message>();

                foreach (var message in messages)
                {
                    MyMessages.Add(new Models.Message
                    {
                        Id = message.Id,
                        Sender = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Name :
                                  "Unknown name",
                        From = (message.Sender != null) ?
                                  message.Sender.EmailAddress.Address :
                                  "Unknown email",
                        Subject = message.Subject ?? "No subject",
                        Importance = message.Importance.ToString()
                    });
                }

                MessageCountTextBlock.Text = $"You have {MyMessages.Count()} red bang messages:";
                MessagesListView.ItemsSource = MyMessages;
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not get messages: {ex.Error.Message}";
            }
        }

        private async void MessagesListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            if (MessagesListView.SelectedItem != null)
            {
                selectedMessage = ((Models.Message)MessagesListView.SelectedItem);

                myMessage = await graphClient.Me.Messages[selectedMessage.Id].Request().GetAsync();

                SenderTextBlock.Text = (myMessage.Sender != null) ?
                                        myMessage.Sender.EmailAddress.Name :
                                        "Unknown name";
                FromTextBlock.Text = (myMessage.Sender != null) ?
                                      myMessage.Sender.EmailAddress.Address :
                                      "Unknown email";
                SubjectTextBlock.Text = myMessage.Subject ?? "No subject";
                PreviewTextBlock.Text = myMessage.BodyPreview;
                DateTextBlock.Text = (myMessage.SentDateTime != null) ?
                                      $"{myMessage.SentDateTime.Value.Date:M/d/yyyy}" :
                                      "unknown date";
                ImportanceTextBlock.Text = myMessage.Importance.ToString();
                IsReadTextBlock.Text = (myMessage.IsRead == true) ? "Yes" : "No";
                AttachmentsTextBlock.Text = (myMessage.HasAttachments == true) ? "Yes" : "No";
            }
        }

        private async void SendMessageButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            var recipients = new List<Recipient>();
            recipients.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = "rgreen2005@msn.com"
                }
            });

            var messageToSend = new Message
            {
                ToRecipients = recipients,
                Subject = "Urgent",
                Body = new ItemBody
                {
                    Content = "Call me immediately if you don't get this message!!",
                    ContentType = BodyType.Html
                },
            };

            try
            {
                await graphClient.Me.SendMail(messageToSend, true).Request().PostAsync();
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not send this message: {ex.Error.Message}";
            }
        }

        private async void ReplyMessageButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            string replyText = "Thanks for your mail. I will treat it with all the seriousness it deserves.";

            try
            {
                await graphClient.Me.Messages[selectedMessage.Id].Reply(replyText).Request().PostAsync();
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not reply to this message: {ex.Error.Message}";
            }
        }

        private async void ForwardMessageButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            var recipients = new List<Recipient>
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = "rgreen2005@msn.com"
                    }
                }
            };
            string forwardText = "Thought you might be interested in this. I am not.";

            try
            {
                await graphClient.Me.Messages[selectedMessage.Id].Forward(forwardText, recipients).Request().PostAsync();
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not forward this message: {ex.Error.Message}";
            }
        }

        private async void DeleteMessageButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                await graphClient.Me.Messages[selectedMessage.Id].Request().DeleteAsync();
            }
            catch (ServiceException ex)
            {
                MessageCountTextBlock.Text = $"We could not get delete this message: {ex.Error.Message}";
            }
        }

        private void ShowSliptView(object sender, RoutedEventArgs e)
        {
            MySamplesPane.SamplesSplitView.IsPaneOpen = !MySamplesPane.SamplesSplitView.IsPaneOpen;
        }
    }
}
