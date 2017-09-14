using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

namespace GraphSDKDemo
{
    public sealed partial class DriveItemsPage : Page
    {
        GraphServiceClient graphClient = null;

        IDriveItemChildrenCollectionPage files = null;
        IDriveItemChildrenCollectionPage folders = null;
        ObservableCollection<Models.File> MyFiles = null;
        ObservableCollection<Models.Folder> MyFolders = null;

        Models.File myFile = null;
        Models.File selectedFile = null;
        Models.Folder myFolder = null;
        Models.Folder selectedFolder = null;

        string driveItemSubject = string.Empty;

        public DriveItemsPage()
        {
            this.InitializeComponent();
        }

        private async void GetFoldersButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            FoldersListView.Visibility = Visibility.Visible;
            FilesListView.Visibility = Visibility.Collapsed;

            try
            {
                folders= await graphClient.Me.Drive.Root.Children.Request()
                                             .Select("name,folder").Filter("file eq null").GetAsync();

                MyFolders = new ObservableCollection<Models.Folder>();

                while (true)
                {
                    foreach (var folder in folders)
                    {
                        MyFolders.Add(new Models.Folder
                        {
                            Id = folder.Id,
                            Name = folder.Name,
                            FileCount = (int)folder.Folder.ChildCount
                        });
                    }

                    if (folders.NextPageRequest == null)
                    {
                        break;
                    }
                    folders = await folders.NextPageRequest.GetAsync();
                }

                DriveItemCountTextBlock.Text = $"You have {MyFolders.Count()} folders";
                FoldersListView.ItemsSource = MyFolders;
            }
            catch (ServiceException ex)
            {
                DriveItemCountTextBlock.Text = $"We could not get folders: {ex.Error.Message}";
            }
        }

        private async void GetFilesButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            FilesListView.Visibility = Visibility.Visible;
            FoldersListView.Visibility = Visibility.Collapsed;

            try
            {
                files = await graphClient.Me.Drive.Root.Children.Request()
                                            .Select("name,size").Filter("folder eq null").GetAsync();

                MyFiles = new ObservableCollection<Models.File>();

                while (true)
                {
                    foreach (var file in files)
                    {
                        MyFiles.Add(new Models.File
                        {
                            Id = file.Id,
                            Name = file.Name,
                            Size = Convert.ToInt64(file.Size)
                        });
                    }

                    if (files.NextPageRequest == null)
                    {
                        break;
                    }
                    files = await files.NextPageRequest.GetAsync();
                }

                DriveItemCountTextBlock.Text = $"You have {MyFiles.Count()} files";
                FilesListView.ItemsSource = MyFiles;
            }
            catch (ServiceException ex)
            {
                DriveItemCountTextBlock.Text = $"We could not get files: {ex.Error.Message}";
            }
        }

        private async void CheckFolderButton_Click(Object sender, RoutedEventArgs e)
        {

        }

        private async void CheckFileButton_Click(Object sender, RoutedEventArgs e)
        {

        }

        private async void CreateFolderButton_Click(Object sender, RoutedEventArgs e)
        {

        }

        private async void CreateFileButton_Click(Object sender, RoutedEventArgs e)
        {

        }

        private async void UpdateFileButton_Click(Object sender, RoutedEventArgs e)
        {

        }

        private async void DeleteFileButton_Click(Object sender, RoutedEventArgs e)
        {

        }

        private async void EventsListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {

        }

        private void ShowSplitView(object sender, RoutedEventArgs e)
        {
            MySamplesPane.SamplesSplitView.IsPaneOpen = !MySamplesPane.SamplesSplitView.IsPaneOpen;
        }

        private void FoldersListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {

        }

        private void FilesListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
