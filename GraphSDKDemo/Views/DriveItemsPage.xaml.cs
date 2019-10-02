using Microsoft.Graph;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media.Imaging;

namespace GraphSDKDemo
{
    public sealed partial class DriveItemsPage : Page
    {
        GraphServiceClient graphClient = null;

        IDriveItemChildrenCollectionPage files = null;
        IDriveItemChildrenCollectionPage folders = null;
        ObservableCollection<Models.File> MyFiles = null;
        ObservableCollection<Models.Folder> MyFolders = null;

        IDriveItemSearchCollectionPage searchedItems = null;

        DriveItem myFile = null;
        Models.File selectedFile = null;
        DriveItem myFolder = null;
        Models.Folder selectedFolder = null;

        Stream downloadedFile = null;

        public DriveItemsPage()
        {
            this.InitializeComponent();

            graphClient = (App.Current as App).GraphClient;
        }

        private async void GetFoldersButton_Click(Object sender, RoutedEventArgs e)
        {
            ShowFolders();
            ClearFoldersAndFiles();

            try
            {
                folders = await graphClient.Me.Drive.Root.Children.Request()
                                              .Select("id,name,folder")
                                              .Filter("file eq null").GetAsync();

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
            ShowFiles();

            try
            {
                if (FoldersListView.SelectedItem != null)
                {
                    selectedFolder = ((Models.Folder)FoldersListView.SelectedItem);

                    files = await graphClient.Me.Drive.Items[selectedFolder.Id]
                                             .Children.Request()
                                             .Select("id,name,size,weburl")
                                             .Filter("folder eq null").GetAsync();
                }
                else
                {
                    files = await graphClient.Me.Drive.Root.Children.Request()
                                            .Select("id,name,size,weburl")
                                            .Filter("folder eq null").GetAsync();
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
                            Size = Convert.ToInt64(file.Size),
                            Url = file.WebUrl
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
            if (FoldersListView.SelectedItem != null)
            {
                selectedFolder = ((Models.Folder)FoldersListView.SelectedItem);

                myFolder = await graphClient.Me.Drive.Items[selectedFolder.Id]
                                            .Request().GetAsync();

                FolderNameTextBlock.Text = myFolder.Name;
                FileCountTextBlock.Text = 
                    ((int)myFolder.Folder.ChildCount).ToString("N0");
                FolderCreatedTextBlock.Text = 
                    myFolder.CreatedDateTime.GetValueOrDefault().ToString("d");
                FolderLastModifiedTextBlock.Text = 
                    myFolder.LastModifiedDateTime.GetValueOrDefault().ToString("d");
                FolderSharedTextBlock.Text = (myFolder.Shared != null) ? "Yes" : "No";
                FolderHyperlinkButton.NavigateUri = new Uri(myFolder.WebUrl);
            }
        }

        private async void FilesListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {
            if (FilesListView.SelectedItem != null)
            {
                selectedFile = ((Models.File)FilesListView.SelectedItem);

                myFile = await graphClient.Me.Drive.Items[selectedFile.Id]
                                          .Request().GetAsync();

                FileNameTextBlock.Text = myFile.Name;
                FileSizeTextBlock.Text = (Convert.ToInt64(myFile.Size)).ToString("N0");
                FileCreatedTextBlock.Text = 
                    myFile.CreatedDateTime.GetValueOrDefault().ToString("d");
                FileLastModifiedTextBlock.Text = 
                    myFile.LastModifiedDateTime.GetValueOrDefault().ToString("d");
                FileSharedTextBlock.Text = (myFile.Shared != null) ? "Yes" : "No";
            }
        }

        private async void DisplayButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                downloadedFile = await graphClient.Me.Drive.Items[selectedFile.Id]
                                                  .Content.Request().GetAsync();

                var memStream = new MemoryStream();

                // Convert the stream to the memory stream, because a memory stream supports seeking.
                await downloadedFile.CopyToAsync(memStream);

                // Set the start position.
                memStream.Position = 0;

                // Create a new bitmap image.
                var bitmap = new BitmapImage();

                // Set the bitmap source to the stream, which is converted to a IRandomAccessStream.
                bitmap.SetSource(memStream.AsRandomAccessStream());

                // Set the image control source to the bitmap.
                FileImage.Source = bitmap;

            }
            catch (Exception ex)
            {
                string wtf = ex.Message;
            }
        }

        private async void SearchFoldersButton_Click(Object sender, RoutedEventArgs e)
        {
            ShowFolders();

            try
            {
                searchedItems = await graphClient.Me.Drive.Root
                                                 .Search("tax").Request().GetAsync();

                MyFolders = new ObservableCollection<Models.Folder>();

                foreach (var item in searchedItems)
                {
                    if (item.Folder != null)
                    {
                        MyFolders.Add(new Models.Folder
                        {
                            Id = item.Id,
                            Name = item.Name,
                            FileCount = (int)item.Folder.ChildCount
                        });
                    }
                }

                DriveItemCountTextBlock.Text = $"You have {MyFolders.Count()} folders";
                FoldersListView.ItemsSource = MyFolders;
            }
            catch (ServiceException ex)
            {
                DriveItemCountTextBlock.Text = $"We could not get folders: {ex.Error.Message}";
            }
        }

        private async void SearchFilesButton_Click(Object sender, RoutedEventArgs e)
        {
            ShowFiles();

            try
            {
                if (FoldersListView.SelectedItem != null)
                {
                    selectedFolder = ((Models.Folder)FoldersListView.SelectedItem);

                    searchedItems = await graphClient.Me.Drive.Items[selectedFolder.Id]
                                                     .Search("Direct Deposit")
                                                     .Request().GetAsync();
                }

                MyFiles = new ObservableCollection<Models.File>();

                foreach (var file in searchedItems)
                {
                    MyFiles.Add(new Models.File
                    {
                        Id = file.Id,
                        Name = file.Name,
                        Size = Convert.ToInt64(file.Size),
                        Url = file.WebUrl
                    });
                }

                DriveItemCountTextBlock.Text = $"You have {MyFiles.Count()} files";
                FilesListView.ItemsSource = MyFiles;
            }
            catch (ServiceException ex)
            {
                DriveItemCountTextBlock.Text = $"We could not get files: {ex.Error.Message}";
            }
        }

        private void ShowFolders()
        {
            FoldersListView.Visibility = Visibility.Visible;
            FilesListView.Visibility = Visibility.Collapsed;
            FolderScrollViewer.Visibility = Visibility.Visible;
            FileScrollViewer.Visibility = Visibility.Collapsed;
        }

        private void ShowFiles()
        {
            FilesListView.Visibility = Visibility.Visible;
            FoldersListView.Visibility = Visibility.Collapsed;
            FileScrollViewer.Visibility = Visibility.Visible;
            FolderScrollViewer.Visibility = Visibility.Collapsed;
        }

        private void ClearFoldersAndFiles()
        {
            FolderNameTextBlock.Text = string.Empty;
            FileCountTextBlock.Text = string.Empty;
            FolderCreatedTextBlock.Text = string.Empty;
            FolderLastModifiedTextBlock.Text = string.Empty;
            FolderSharedTextBlock.Text = string.Empty;
            FileNameTextBlock.Text = string.Empty;
            FileSizeTextBlock.Text = string.Empty;
            FileCreatedTextBlock.Text = string.Empty;
            FileLastModifiedTextBlock.Text = string.Empty;
            FileSharedTextBlock.Text = string.Empty;
            FileImage.Source = null;
        }

   }
}
