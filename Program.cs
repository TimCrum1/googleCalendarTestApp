using System;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using System.Configuration;
using System.IO;
using System.Threading;
using System.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace GoogleCalendarTestApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Get the OAuth2 creds from file (for CEF we will set up a different way to build the UserCredential)
                string credentialFilePath = "C:\\Data\\Clients\\VALHAL\\GoogleCalendarTestApp\\GoogleCalendarTestApp\\client_secret.json";
                UserCredential credential;
                using (var stream = new FileStream(credentialFilePath, FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        new[] { CalendarService.Scope.Calendar },
                        "user",
                        CancellationToken.None).Result;
                }
                // Create the CalendarService
                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "TestProject",
                });
                // Create the new CalendarEvent
                Event newEvent = new Event()
                {
                    Summary = "Test Meeting",
                    Description = "Test Meeting first go around",
                    Start = new EventDateTime()
                    {
                        DateTimeDateTimeOffset = DateTime.Now.AddHours(2),
                        TimeZone = "EST",
                    },
                    End = new EventDateTime()
                    {
                        DateTimeDateTimeOffset = DateTime.Now.AddHours(2.5),
                        TimeZone = "EST",
                    },
                    // ConferenceData is required for adding the Google Meet video conference to the Calendar Event
                    ConferenceData = new ConferenceData
                    {
                        // All fields required to get Google meet to work
                        ConferenceSolution = new ConferenceSolution()
                        {
                            Key = new ConferenceSolutionKey()
                            {
                                Type = "hangoutsMeet",
                            },
                        },
                        CreateRequest = new CreateConferenceRequest
                        {
                            RequestId = Guid.NewGuid().ToString(),
                            ConferenceSolutionKey = new ConferenceSolutionKey()
                            {
                                Type = "hangoutsMeet"
                            }
                        },
                        EntryPoints = new List<EntryPoint>()
                        {
                            new EntryPoint()
                            {
                                EntryPointType = "video",
                            },
                        },
                    },
                    // Add in the invited attendees which will be Doctor and Patient emails
                    Attendees = new List<EventAttendee>()
                    {
                        new EventAttendee { Email = "doctor@email.com" },
                        new EventAttendee { Email = "patient@email.com"}
                    }
                };

                // create the request
                var createdEventRequest = service.Events.Insert(newEvent, "primary");

                // Mark sending notifications true
                createdEventRequest.SendNotifications = true;

                // Without this the request will ignore the ConferenceData object
                createdEventRequest.ConferenceDataVersion = 1;

                // Create the CalendarEvent with a google meet link!!
                Event createdEvent = createdEventRequest.Execute();

                // Get the Google Meet Link from the response
                string meetingLink = createdEvent.ConferenceData?.EntryPoints
                    .FirstOrDefault(x => x.EntryPointType == "video")?.Uri;

                if (meetingLink != null)
                {
                    Console.WriteLine($"Meeting Link: {meetingLink}");
                }
                else if (createdEvent != null)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(createdEvent));
                    Console.WriteLine(createdEvent.ConferenceData);
                }
                else
                {
                    Console.WriteLine("You suck!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
