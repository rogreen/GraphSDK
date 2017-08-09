using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace GraphSDKDemo
{
    public sealed partial class EventsPage : Page
    {
        GraphServiceClient graphClient = null;

        ObservableCollection<Models.Event> MyEvents = null;

        Event myEvent = null;
        Models.Event selectedEvent = null;

        string eventSubject = string.Empty;

        public EventsPage()
        {
            this.InitializeComponent();
        }

        private async void GetEventsButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {
                //IUserEventsCollectionPage events =
                //    await graphClient.Me.Events.Request()
                //                               .Filter("start/datetime ge '2017-08-06'")
                //                               .Select("subject,organizer,location,start,end").GetAsync();


                var options = new List<QueryOption>
                {
                    new QueryOption("startDateTime", DateTime.Now.ToString("o")),
                    new QueryOption("endDateTime", DateTime.Now.AddDays(7).ToString("o"))
                };
                ICalendarCalendarViewCollectionPage calendar =
                    await graphClient.Me.Calendar.CalendarView.Request(options)
                                                              .Select("subject,organizer,location,start,end")
                                                              .GetAsync();

                MyEvents = new ObservableCollection<Models.Event>();

                foreach (var myEvent in calendar)
                {
                    MyEvents.Add(new Models.Event
                    {
                        Id = myEvent.Id,
                        Subject = myEvent.Subject ?? "No subject",
                        Location = (myEvent.Location != null) ?
                                    myEvent.Location.DisplayName :
                                    "Unknown location",
                        Organizer = (myEvent.Organizer != null) ?
                                     myEvent.Organizer.EmailAddress.Name :
                                     "Unknown organizer",
                        Start = DateTime.Parse(myEvent.Start.DateTime).ToLocalTime(),
                        End = DateTime.Parse(myEvent.End.DateTime).ToLocalTime()
                    });
                }

                EventCountTextBlock.Text = $"You have {calendar.Count()} events in the next week:";
                EventsListView.ItemsSource = MyEvents;
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not get events: {ex.Error.Message}";
            }
        }

        private async void GetTodayEventsButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {
                var options = new List<QueryOption>
                {
                    new QueryOption("startDateTime", DateTime.Now.ToString("o")),
                    new QueryOption("endDateTime", DateTime.Now.ToString("o"))
                };
                ICalendarCalendarViewCollectionPage calendar =
                    await graphClient.Me.Calendar.CalendarView.Request(options)
                                                              .Select("subject,organizer,location,start,end")
                                                              .GetAsync();

                MyEvents = new ObservableCollection<Models.Event>();

                foreach (var myEvent in calendar)
                {
                    MyEvents.Add(new Models.Event
                    {
                        Id = myEvent.Id,
                        Subject = (myEvent.Subject != null) ?
                                   myEvent.Subject :
                                   "No subject",
                        Location = (myEvent.Location != null) ?
                                    myEvent.Location.DisplayName :
                                    "Unknown location",
                        Organizer = (myEvent.Organizer != null) ?
                                     myEvent.Organizer.EmailAddress.Name :
                                     "Unknown organizer",
                        Start = DateTime.Parse(myEvent.Start.DateTime).ToLocalTime(),
                        End = DateTime.Parse(myEvent.End.DateTime).ToLocalTime()
                    });
                }

                EventCountTextBlock.Text = $"You have {calendar.Count()} events today:";
                EventsListView.ItemsSource = MyEvents;
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not get today's events: {ex.Error.Message}";
            }
        }

        private async void GetBirthdaysButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {
                var options = new List<QueryOption>
                {
                    new QueryOption("startDateTime", DateTime.Now.ToString("o")),
                    new QueryOption("endDateTime", DateTime.Now.AddDays(7).ToString("o"))
                };
                ICalendarCalendarViewCollectionPage calendar =
                    await graphClient.Me.Calendar.CalendarView.Request(options)
                                                              .Filter("categories/any(categories: categories eq 'Birthday')")
                                                              .Select("subject,start,categories").GetAsync();

                MyEvents = new ObservableCollection<Models.Event>();

                foreach (var myEvent in calendar)
                {
                    MyEvents.Add(new Models.Event
                    {
                        Id = myEvent.Id,
                        Subject = myEvent.Subject ?? "No subject",
                        Start = DateTime.Parse(myEvent.Start.DateTime).ToLocalTime()
                    });
                }

                EventCountTextBlock.Text = $"You have {MyEvents.Count()} birthdays in the next week:";
                EventsListView.ItemsSource = MyEvents;
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not get birthdays: {ex.Error.Message}";
            }
        }

        private async void EventsListView_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            if (EventsListView.SelectedItem != null)
            {
                selectedEvent = ((Models.Event)EventsListView.SelectedItem);

                myEvent = await graphClient.Me.Events[selectedEvent.Id].Request().GetAsync();

                SubjectTextBlock.Text = myEvent.Subject ?? "No subject";
                eventSubject = myEvent.Subject;

                LocationTextBlock.Text = (myEvent.Location != null) ?
                                          $"{myEvent.Location.DisplayName}\n" +
                                          $"{myEvent.Location.Address.Street}\n" +
                                          $"{myEvent.Location.Address.City} " +
                                          $"{myEvent.Location.Address.State}" :
                                          "Unknown location";
                OrganizerTextBlock.Text = (myEvent.Organizer != null) ?
                                           $"{myEvent.Organizer.EmailAddress.Name}\n" +
                                           $"{myEvent.Organizer.EmailAddress.Address}" :
                                           "Unknown organizer";
                StartTextBlock.Text = DateTime.Parse(myEvent.Start.DateTime).ToLocalTime().ToString();
                EndTextBlock.Text = DateTime.Parse(myEvent.End.DateTime).ToLocalTime().ToString();
                AllDayTextBlock.Text = (myEvent.IsAllDay == true) ? "Yes" : "No";
                RecurringTextBlock.Text = (myEvent.Recurrence != null) ? "Yes" : "No";

                string categories = string.Empty;
                foreach (var category in myEvent.Categories)
                {
                    categories += $"{category} ";
                }
                CategoriesTextBlock.Text = categories;
            }
        }

        private async void CreateEventButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();


            //List of attendees
            var attendees = new List<Attendee>();
            var attendee = new Attendee();
            var emailAddress = new EmailAddress()
            {
                Address = "rgreen2005@msn.com"
            };
            attendee.EmailAddress = emailAddress;
            attendee.Type = AttendeeType.Required;
            attendees.Add(attendee);

            //Event body
            var eventBody = new ItemBody()
            {
                Content = "Status updates, blocking issues, and next steps",
                ContentType = BodyType.Text
            };

            //Event start and end time
            var eventStartTime = new DateTimeTimeZone()
            {
                DateTime = DateTime.Today.AddDays(1).ToString("o"),
                TimeZone = "UTC"
            };
            var eventEndTime = new DateTimeTimeZone()
            {
                TimeZone = "UTC",
                DateTime = DateTime.Today.AddDays(1).ToString("o")
            };

            //Create an event to add to the events collection

            var location = new Location()
            {
                DisplayName = "Water cooler"
            };
            var newEvent = new Event()
            {
                Subject = "Weekly sync",
                Location = location,
                Attendees = attendees,
                Body = eventBody,
                Start = eventStartTime,
                End = eventEndTime
            };
            try
            {
                await graphClient.Me.Events.Request().AddAsync(newEvent);
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not create this event: {ex.Error.Message}";
            }
        }

        private async void UpdateEventButton_Click(Object sender, RoutedEventArgs e)
        {
            graphClient = AuthenticationHelper.GetAuthenticatedClient();

            var eventToUpdate = new Event()
            {
                Subject = $"{eventSubject} (Updated)"
            };
            try
            {
                await graphClient.Me.Events[selectedEvent.Id].Request().UpdateAsync(eventToUpdate);
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not update this event: {ex.Error.Message}";
            }
        }

        private async void DeleteEventButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                await graphClient.Me.Events[selectedEvent.Id].Request().DeleteAsync();
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not get delete this event: {ex.Error.Message}";
            }
        }

        private void ShowSliptView(object sender, RoutedEventArgs e)
        {
            MySamplesPane.SamplesSplitView.IsPaneOpen = !MySamplesPane.SamplesSplitView.IsPaneOpen;
        }
    }
}
