using System;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace GraphSDKDemo
{
    public sealed partial class SamplesPane : UserControl
    {
        public SamplesPane()
        {
            this.InitializeComponent();
        }
        private void MessagesButton_Click(Object sender, RoutedEventArgs e)
        {
            ((Frame)Window.Current.Content).Navigate(typeof(MessagesPage));
        }

        private void ContactsButton_Click(Object sender, RoutedEventArgs e)
        {
            ((Frame)Window.Current.Content).Navigate(typeof(ContactsPage));
        }

        private void EventsButton_Click(Object sender, RoutedEventArgs e)
        {
            ((Frame)Window.Current.Content).Navigate(typeof(EventsPage));
        }

        private void NavigateToHome(object sender, RoutedEventArgs e)
        {
            ((Frame)Window.Current.Content).Navigate(typeof(MainPage));
        }

        private void DriveItemsButton_Click(Object sender, RoutedEventArgs e)
        {
            ((Frame)Window.Current.Content).Navigate(typeof(DriveItemsPage));
        }
    }
}
