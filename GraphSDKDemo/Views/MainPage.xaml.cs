using Microsoft.Toolkit.Services.MicrosoftGraph;
using System;
using Windows.UI.Xaml.Controls;

namespace GraphSDKDemo
{
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();

            // Initialize auth state to false
            SetAuthState(false);

            // Load OAuth settings
            var oauthSettings = Windows.ApplicationModel.Resources.ResourceLoader.GetForCurrentView("OAuth");
            var appId = oauthSettings.GetString("AppId");
            var scopes = oauthSettings.GetString("Scopes");

            if (string.IsNullOrEmpty(appId) || string.IsNullOrEmpty(scopes))
            {
                Notification.Show("Could not load OAuth Settings from resource file.");
            }
            else
            {
                // Initialize Graph
                MicrosoftGraphService.Instance.AuthenticationModel = MicrosoftGraphEnums.AuthenticationModel.V2;
                MicrosoftGraphService.Instance.Initialize(appId,
                    MicrosoftGraphEnums.ServicesToInitialize.UserProfile,
                    scopes.Split(' '));

                // Navigate to HomePage.xaml
                RootFrame.Navigate(typeof(HomePage));
            }
        }

        private void SetAuthState(bool isAuthenticated)
        {
            (App.Current as App).IsAuthenticated = isAuthenticated;
        }

        private void NavView_ItemInvoked(NavigationView sender, NavigationViewItemInvokedEventArgs args)
        {
            var invokedItem = args.InvokedItem as string;

            switch (invokedItem.ToLower())
            {
                case "messages":
                    RootFrame.Navigate(typeof(MessagesPage));
                    break;
                case "contacts":
                    RootFrame.Navigate(typeof(ContactsPage));
                    break;
                case "events":
                    RootFrame.Navigate(typeof(EventsPage));
                    break;
                case "driveitems":
                    RootFrame.Navigate(typeof(DriveItemsPage));
                    break;
                case "home":
                default:
                    RootFrame.Navigate(typeof(HomePage));
                    break;
            }
        }

        private void Login_SignInCompleted(object sender, Microsoft.Toolkit.Uwp.UI.Controls.Graph.SignInEventArgs e)
        {
            // Set the auth state
            SetAuthState(true);
            // Reload the home page
            RootFrame.Navigate(typeof(HomePage));
        }

        private void Login_SignOutCompleted(object sender, EventArgs e)
        {
            // Set the auth state
            SetAuthState(false);
            // Reload the home page
            RootFrame.Navigate(typeof(HomePage));
        }
    }
}
