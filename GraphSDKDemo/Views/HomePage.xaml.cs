using Microsoft.Toolkit.Services.MicrosoftGraph;
using Windows.UI.Xaml.Controls;

namespace GraphSDKDemo
{
    public sealed partial class HomePage : Page
    {
        public HomePage()
        {
            this.InitializeComponent();

            if ((App.Current as App).IsAuthenticated)
            {
                // Get the Graph client from the service
                (App.Current as App).GraphClient = MicrosoftGraphService.Instance.GraphProvider;

                HomePageMessage.Text = "Welcome! Please use the menu to the left to select a view.";
            }
        }
    }
}
