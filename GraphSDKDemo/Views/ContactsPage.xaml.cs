using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace GraphSDKDemo
{
    public sealed partial class ContactsPage : Page
    {
        GraphServiceClient graphClient = null;

        IUserContactsCollectionPage contacts = null;
        ObservableCollection<Models.Contact> MyContacts = null;

        Contact myContact = null;
        Models.Contact selectedContact = null;

        public ContactsPage()
        {
            this.InitializeComponent();

            graphClient = (App.Current as App).GraphClient;
        }

        private async void GetContactsButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                //contacts = await graphClient.Me.Contacts.Request().Top(20).GetAsync();
                // contacts = await graphClient.Me.Contacts.Request().OrderBy("displayName").GetAsync();
                contacts = await graphClient.Me.Contacts.Request().OrderBy("displayName")
                                            .Select("displayName,emailAddresses").GetAsync();
                //contacts = await graphClient.Me.Contacts.Request()
                //                            .Filter("startswith(displayName,'A'")
                //                            .Select("displayName,emailAddresses").GetAsync();

                MyContacts = new ObservableCollection<Models.Contact>();

                while (true)
                {
                    foreach (var contact in contacts)
                    {
                        if (contact.DisplayName.StartsWith("A"))
                        {
                            MyContacts.Add(new Models.Contact
                            {
                                Id = contact.Id,
                                DisplayName = contact.DisplayName,
                                EmailAddress = (contact.EmailAddresses.Count() > 0) ?
                                               $"{contact.EmailAddresses.First().Address}" :
                                               "Unknown email",
                            });
                        }
                    }

                    if (MyContacts.Count() >= 20) { break; }

                    if (contacts.NextPageRequest == null) { break; }

                    contacts = await contacts.NextPageRequest.GetAsync();
                }

                ContactCountTextBlock.Text = $"Here are your first {MyContacts.Count()} contacts that start with A";
                ContactsDataGrid.ItemsSource = MyContacts;

            }
            catch (ServiceException ex)
            {
                ContactCountTextBlock.Text = $"We could not get contacts: {ex.Error.Message}";
            }
        }

        private async void ContactsDataGrid_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {
            if (ContactsDataGrid.SelectedItem != null)
            {
                selectedContact = ((Models.Contact)ContactsDataGrid.SelectedItem);

                // Note: This api does not support using a filter, 
                // so you can only get a particular contact via the Id

                myContact = await graphClient.Me.Contacts[selectedContact.Id]
                                             .Request().GetAsync();

                DisplayNameTextBlock.Text = (myContact.DisplayName != string.Empty) ?
                                             myContact.DisplayName :
                                             "Unknown name";
                EmailAddressTextBlock.Text = (myContact.EmailAddresses.Count() > 0) ?
                                              $"{myContact.EmailAddresses.First().Address}" :
                                              "Unknown email";
                CompanyTextBlock.Text = myContact.CompanyName ?? "";
                JobTitleTextBlock.Text = myContact.JobTitle ?? "";
                BusinessPhoneTextBlock.Text = (myContact.BusinessPhones.Count() > 0) ?
                                               myContact.BusinessPhones.First() : "";
                HomePhoneTextBlock.Text = (myContact.HomePhones.Count() > 0) ?
                                           myContact.HomePhones.First() : "";
                MobilePhoneTextBlock.Text = myContact.MobilePhone ?? "";
                NotesTextBlock.Text = myContact.PersonalNotes ?? "";
            }
        }

        private async void UpdateContactButton_Click(Object sender, RoutedEventArgs e)
        {
            var contactToUpdate = new Contact();
            contactToUpdate.PersonalNotes = "My best friend ever!";

            try
            {
                var updatedContact = await graphClient.Me.Contacts[selectedContact.Id]
                                                      .Request()
                                                      .UpdateAsync(contactToUpdate);
            }
            catch (ServiceException ex)
            {
                ContactCountTextBlock.Text = $"We could not get update this contact: {ex.Error.Message}";
            }
        }


        private async void AddContactButton_Click(Object sender, RoutedEventArgs e)
        {
            var contactToAdd = new Contact()
            {
                GivenName = "Rufus T Firefly",
                DisplayName = "Rufus T Firefly"
            };
            var emailAddresses = new List<EmailAddress>();
            var emailAddress = new EmailAddress()
            {
                Address = "rufus@northwind.com"
            };
            emailAddresses.Add(emailAddress);
            contactToAdd.EmailAddresses = emailAddresses;

            contactToAdd.CompanyName = "Northwind Traders";
            contactToAdd.JobTitle = "CEO";
            contactToAdd.HomePhones = null;

            var businessPhones = new List<String>
            {
                "555-555-1212"
            };
            contactToAdd.BusinessPhones = businessPhones;

            contactToAdd.MobilePhone = "555-555-1213";
            contactToAdd.PersonalNotes = "";

            try
            {
                var updatedContact = await graphClient.Me.Contacts.Request()
                                                      .AddAsync(contactToAdd);
            }
            catch (ServiceException ex)
            {
                ContactCountTextBlock.Text = $"We could not get add this contact: {ex.Error.Message}";
            }
        }

        private async void DeleteContactButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                await graphClient.Me.Contacts[selectedContact.Id].Request().DeleteAsync();
            }
            catch (ServiceException ex)
            {
                ContactCountTextBlock.Text = $"We could not get delete this contact: {ex.Error.Message}";
            }
        }

    }
}
