using Microsoft.Graph;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;

namespace GraphSDKDemo
{
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }
        private void ShowSplitView(object sender, RoutedEventArgs e)
        {
            MySamplesPane.SamplesSplitView.IsPaneOpen = !MySamplesPane.SamplesSplitView.IsPaneOpen;
        }

        protected override async void OnNavigatedTo(NavigationEventArgs e)
        {
            GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

            var currentUser = await graphClient.Me.Request().GetAsync();
            UserNameTextBlock.Text = $"Welcome {currentUser.DisplayName}";
        }

    }
}
