﻿using Microsoft.Graph;
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

        ICalendarCalendarViewCollectionPage calendar = null;

        ObservableCollection <Models.Event> MyEvents = null;

        Event myEvent = null;
        Models.Event selectedEvent = null;

        string eventSubject = string.Empty;

        public EventsPage()
        {
            this.InitializeComponent();

            graphClient = (App.Current as App).GraphClient;
        }

        private async void GetEventsButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                var options = new List<QueryOption>
                {
                    new QueryOption("startDateTime", DateTime.Today.ToString("o")),
                    new QueryOption("endDateTime", DateTime.Today.AddDays(7).ToString("o"))
                };
                calendar = await graphClient.Me.Calendar.CalendarView.Request(options)
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
                EventsDataGrid.ItemsSource = MyEvents;
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not get events: {ex.Error.Message}";
            }
        }

        private async void GetTodayEventsButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                var options = new List<QueryOption>
                {
                    new QueryOption("startDateTime", DateTime.Today.ToString("o")),
                    new QueryOption("endDateTime", DateTime.Today.AddDays(1).ToString("o"))
                };
                calendar = await graphClient.Me.Calendar.CalendarView.Request(options)
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
                EventsDataGrid.ItemsSource = MyEvents;
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not get today's events: {ex.Error.Message}";
            }
        }

        private async void GetBirthdaysButton_Click(Object sender, RoutedEventArgs e)
        {
            try
            {
                var options = new List<QueryOption>
                {
                    new QueryOption("startDateTime", DateTime.Now.ToString("o")),
                    new QueryOption("endDateTime", DateTime.Now.AddDays(14).ToString("o"))
                };
                calendar = await graphClient.Me.Calendar.CalendarView.Request(options)
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

                EventCountTextBlock.Text = $"You have {MyEvents.Count()} birthdays in the next 2 weeks:";
                EventsDataGrid.ItemsSource = MyEvents;
            }
            catch (ServiceException ex)
            {
                EventCountTextBlock.Text = $"We could not get birthdays: {ex.Error.Message}";
            }
        }

        private async void EventsDataGrid_SelectionChanged(Object sender, SelectionChangedEventArgs e)
        {
            if (EventsDataGrid.SelectedItem != null)
            {
                selectedEvent = ((Models.Event)EventsDataGrid.SelectedItem);

                myEvent = await graphClient.Me.Events[selectedEvent.Id].Request().GetAsync();

                SubjectTextBlock.Text = myEvent.Subject ?? "No subject";
                eventSubject = myEvent.Subject;

                LocationTextBlock.Text = (myEvent.Location != null) ?
                                          $"{myEvent.Location.DisplayName}\n" +
                                          $"{myEvent.Location.Address?.Street}\n" +
                                          $"{myEvent.Location.Address?.City} " +
                                          $"{myEvent.Location.Address?.State}" :
                                          "Unknown location";
                OrganizerTextBlock.Text = (myEvent.Organizer != null) ?
                                           $"{myEvent.Organizer.EmailAddress.Name}\n" +
                                           $"{myEvent.Organizer.EmailAddress.Address}" :
                                           "Unknown organizer";
                StartTextBlock.Text = 
                    DateTime.Parse(myEvent.Start.DateTime).ToLocalTime().ToString();
                EndTextBlock.Text = 
                    DateTime.Parse(myEvent.End.DateTime).ToLocalTime().ToString();
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

             var location = new Location()
            {
                DisplayName = "Big conf room"
            };

           //Event start and end time

            DateTime eventDay = DateTime.Today.AddDays(1);

            var eventStartTime = new DateTimeTimeZone()
            {
                DateTime = 
                new DateTime(eventDay.Year, eventDay.Month, eventDay.Day, 9, 0, 0).ToString("o"),
                TimeZone = "Pacific Standard Time"
            };
            var eventEndTime = new DateTimeTimeZone()
            {
                DateTime = 
                new DateTime(eventDay.Year, eventDay.Month, eventDay.Day, 10, 0, 0).ToString("o"),
                TimeZone = "Pacific Standard Time"
            };

            //Create an event to add to the events collection

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
            var eventToUpdate = new Event()
            {
                Subject = $"{eventSubject} (Updated)"
            };
            try
            {
                await graphClient.Me.Events[selectedEvent.Id].Request()
                                 .UpdateAsync(eventToUpdate);
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

    }
}
