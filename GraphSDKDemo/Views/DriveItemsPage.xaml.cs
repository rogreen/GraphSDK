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

        DriveItem myFile = null;
        Models.File selectedFile = null;
        DriveItem myFolder = null;
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
            FolderScrollViewer.Visibility = Visibility.Visible;
            FileScrollViewer.Visibility = Visibility.Collapsed;

            try
            {
                folders = await graphClient.Me.Drive.Root.Children.Request()
                                             .Select("id,name,folder").Filter("file eq null").GetAsync();

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
            FileScrollViewer.Visibility = Visibility.Visible;
            FolderScrollViewer.Visibility = Visibility.Collapsed;

            try
            {
                if (FoldersListView.SelectedItem != null)
                {
                    selectedFolder = ((Models.Folder)FoldersListView.SelectedItem);

                    files = await graphClient.Me.Drive.Items[selectedFolder.Id].Children.Request()
                                             .Select("id,name,size").Filter("folder eq null").GetAsync();
                }
                else
                {
                    files = await graphClient.Me.Drive.Root.Children.Request()
                                            .Select("id,name,size").Filter("folder eq null").GetAsync();
                }

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

        private async void FoldersListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            if (FoldersListView.SelectedItem != null)
            {
                selectedFolder = ((Models.Folder)FoldersListView.SelectedItem);

                myFolder = await graphClient.Me.Drive.Items[selectedFolder.Id].Request().GetAsync();

                FolderNameTextBlock.Text = myFolder.Name;
                FileCountTextBlock.Text = ((int)myFolder.Folder.ChildCount).ToString("N0");
                FolderCreatedTextBlock.Text = myFolder.CreatedDateTime.GetValueOrDefault().ToString("d");
                FolderLastModifiedTextBlock.Text = myFolder.LastModifiedDateTime.GetValueOrDefault().ToString("d");
                FolderSharedTextBlock.Text = (myFolder.Shared != null) ? "Yes" : "No";
                FolderHyperlinkButton.NavigateUri = new Uri(myFolder.WebUrl);
            }
        }

        private async void FilesListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            if (FilesListView.SelectedItem != null)
            {
                selectedFile = ((Models.File)FilesListView.SelectedItem);

                myFile = await graphClient.Me.Drive.Items[selectedFile.Id].Request().GetAsync();

                FileNameTextBlock.Text = myFile.Name;
                FileSizeTextBlock.Text = (Convert.ToInt64(myFile.Size)).ToString("N0");
                FileCreatedTextBlock.Text = myFile.CreatedDateTime.GetValueOrDefault().ToString("d");
                FileLastModifiedTextBlock.Text = myFile.LastModifiedDateTime.GetValueOrDefault().ToString("d");
                FileSharedTextBlock.Text = (myFile.Shared != null) ? "Yes" : "No";
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

        private void ShowSplitView(object sender, RoutedEventArgs e)
        {
            MySamplesPane.SamplesSplitView.IsPaneOpen = !MySamplesPane.SamplesSplitView.IsPaneOpen;
        }

    }
}
