using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using SimpleHttpServer;
using SimpleHttpServer.Models;
using SimpleHttpServer.RouteHandlers;
using System.Net;
using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json;

namespace JarvisExchangeAgent
{
    class Program
    {
        private static ExchangeService service10;
        private static ExchangeService service268;
        private static ExchangeService service37;
        private static ExchangeService service459;
        private static ExchangeService serviceCrestron;

        private static void SetUpServiceAccounts()
        {
            // EWS API Setup
            service10 = new ExchangeService(ExchangeVersion.Exchange2013);
            service268 = new ExchangeService(ExchangeVersion.Exchange2013);
            service37 = new ExchangeService(ExchangeVersion.Exchange2013);
            service459 = new ExchangeService(ExchangeVersion.Exchange2013);
            serviceCrestron = new ExchangeService(ExchangeVersion.Exchange2013);

            service10.Credentials = new WebCredentials("Ic.Jarvis10.ic@canada.ca", "c+Y6Q8g~%N4Y^k2e7n6q%Y$M4R!u9J");
            service268.Credentials = new WebCredentials("Ic.Jarvis268.ic@canada.ca", "6g3x~N#Tc4%c+T+M6G=t3w5Bt6n5E~");
            service37.Credentials = new WebCredentials("Ic.Jarvis37.ic@canada.ca", "#r@Y!m4R!Q7K3d9xB#H82d%c3J@R6h");
            service459.Credentials = new WebCredentials("Ic.Jarvis459.ic@canada.ca", "+g5D2f%NQ6h6z~%EF@6F^v5s@B^v4Y");
            serviceCrestron.Credentials = new WebCredentials("ic.crestron.ic@canada.ca", "2D5n%L+sg4$t!m@Y9R6T3p%F8Q!N4r");

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;
            service10.AutodiscoverUrl("Ic.Jarvis10.ic@canada.ca", RedirectionUrlValidationCallback);
            service268.AutodiscoverUrl("Ic.Jarvis268.ic@canada.ca", RedirectionUrlValidationCallback);
            service37.AutodiscoverUrl("Ic.Jarvis37.ic@canada.ca", RedirectionUrlValidationCallback);
            service459.AutodiscoverUrl("Ic.Jarvis459.ic@canada.ca", RedirectionUrlValidationCallback);
            service459.AutodiscoverUrl("ic.crestron.ic@canada.ca", RedirectionUrlValidationCallback);
            Console.WriteLine("url: " + service10.Url);
        }

        static void Main(string[] args)
        {
            // SetUpServiceAccounts();
            API api = new API();

            log4net.Config.XmlConfigurator.Configure();
            var route_config = new List<Route>() {
                new Route {
                    Name = "Hello Handler",
                    UrlRegex = @"^/$",
                    Method = "POST",
                    Callable = (HttpRequest request) => {
                        try{
                            Console.WriteLine(request.Content);
                            JarvisRecurrenceRequest booking = JsonConvert.DeserializeObject<JarvisRecurrenceRequest>(request.Content);
                            string roomEmail = "ic.conf-ott-235queen-399b-conf.ic@canada.ca";
                            DateTime start = DateTime.Parse("2019-04-14T02:00:00");
                            DateTime end = DateTime.Parse("2019-04-14T04:00:00");
                            string subject = "Jarvis";
                            string body = "daily stand-up";
                            List<string> attendees = null;

                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "Hello from SimpleHttpServer",
                                ReasonPhrase = "OK",
                                StatusCode = "200"
                            };
                        } catch (Exception ex)
                        {
                            Console.WriteLine(ex);

                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "Hello from SimpleHttpServer",
                                ReasonPhrase = "Boo",
                                StatusCode = "300"
                            };
                        }
                    }
                }, 
                new Route {
                    Name = "Test Post",
                    UrlRegex = @"^/booking$",
                    Method = "POST",
                    Callable = (HttpRequest req) => {
                        try {
                            Console.WriteLine(req.Content);
                            JarvisRequest booking = JsonConvert.DeserializeObject<JarvisRequest>(req.Content);

                            string roomEmail = booking.room;
                            DateTime start = DateTime.Parse(booking.start);
                            DateTime end = DateTime.Parse(booking.end);
                            string subject = booking.subject;
                            string body = booking.body;
                            int floor = booking.floor;
                            List<string> attendees = booking.attendees;

                            Console.WriteLine("booking room: " + booking.room);
                            Console.WriteLine("booking start: " + booking.start);
                            Console.WriteLine("booking end: " + booking.end);
                            Console.WriteLine("booking subject: " + booking.subject);
                            Console.WriteLine("booking body: " + booking.body);
                            Console.WriteLine("booking floor: " + booking.floor);

                            //ExchangeService service = GetService(floor);

                            //FindItemsResults<Appointment> appointments = api.GetAppointments(service, roomEmail, start.AddSeconds(1), end);
                            //Console.WriteLine(appointments.TotalCount);

                            if (0 == 0)
                            {
                                //ItemId newAppointment = api.CreateMeeting(service, roomEmail, subject, body, start, end, attendees, null);

                                JarvisResponse res = new JarvisResponse();
                                res.eventId = DateTime.Now.ToString();
                                string resJSON = JsonConvert.SerializeObject(res);
                                return new HttpResponse()
                                {
                                    ContentAsUTF8 = resJSON,
                                    ReasonPhrase = "OK",
                                    StatusCode = "200"
                                };
                            }

                            else
                            {
                                return new HttpResponse()
                                {
                                    ContentAsUTF8 = "conflict",
                                    ReasonPhrase = "Conflict",
                                    StatusCode = "444"
                                };
                            }
                        } catch (Exception e)
                        {
                            Console.WriteLine(e);
                            return new HttpResponse()
                                {
                                    ContentAsUTF8 = "error",
                                    ReasonPhrase = "Error",
                                    StatusCode = "300"
                                };
                        }
                    }
                },
                new Route()
                {
                    Name = "Availability",
                    UrlRegex = @"^/avail$",
                    Method = "POST",
                    Callable = (HttpRequest req) => {
                        JarvisRequest booking = JsonConvert.DeserializeObject<JarvisRequest>(req.Content);
                        DateTime start = DateTime.Parse(booking.start);
                        DateTime end = DateTime.Parse(booking.end);
                        string roomEmail = booking.room;
                        int floor = booking.floor;

                        Console.WriteLine("booking room: " + booking.room);
                        Console.WriteLine("booking start: " + booking.start);
                        Console.WriteLine("booking end: " + booking.end);
                        Console.WriteLine("booking floor: " + booking.floor);

                        ExchangeService service = GetService(floor);
                        //service.TraceEnabled = true;
                        //service.TraceFlags = TraceFlags.All;
                        FindItemsResults<Appointment> appointments = api.GetAppointments(service, roomEmail, start.AddSeconds(1), end);
                        
                        if (0 == appointments.TotalCount)
                        {
                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "free",
                                ReasonPhrase = "OK",
                                StatusCode = "200"
                            };
                        }
                        else
                        {
                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "busy",
                                ReasonPhrase = "OK",
                                StatusCode = "200"
                            };
                        }
                    }
                },
                new Route()
                {
                    Name = "Cancel",
                    UrlRegex = @"^/cancel$",
                    Method = "POST",
                    Callable = (HttpRequest req) => {
                        JarvisRequest booking = JsonConvert.DeserializeObject<JarvisRequest>(req.Content);
                        string roomEmail = booking.room;
                        string eventId = booking.eventID;
                        int floor = booking.floor;

                        Console.WriteLine("booking room: " + booking.room);
                        Console.WriteLine("booking eventId: " + booking.eventID);
                        Console.WriteLine("booking floor: " + booking.floor);

                        ExchangeService service = GetService(floor);
                        //service.TraceEnabled = true;
                        //service.TraceFlags = TraceFlags.All;
                        return new HttpResponse()
                            {
                                ContentAsUTF8 = "free",
                                ReasonPhrase = "OK",
                                StatusCode = "200"
                            };
                        /*
                        if (0 == appointments.TotalCount)
                        {
                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "free",
                                ReasonPhrase = "OK",
                                StatusCode = "200"
                            };
                        }
                        else
                        {
                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "busy",
                                ReasonPhrase = "OK",
                                StatusCode = "200"
                            };
                        }
                        */
                    }
                },  new Route()
                {
                    Name = "Recur",
                    UrlRegex = @"^/recur$",
                    Method = "POST",
                    Callable = (HttpRequest req) => {
                        try {
                            Console.WriteLine(req.Content);
                            JarvisRecurrenceRequest booking = JsonConvert.DeserializeObject<JarvisRecurrenceRequest>(req.Content);

                            Recurrence recurrence = null;
                            switch(booking.type)
                            {
                                case "weekly":
                                    DayOfTheWeek[] daysOfWeek = new DayOfTheWeek[booking.daysOfWeek.Count];
                                    int count = 0;
                                    foreach (int day in booking.daysOfWeek)
                                    {
                                        daysOfWeek[count++] = (DayOfTheWeek)day;
                                    }
                                    recurrence = new Recurrence.WeeklyPattern(DateTime.Parse(booking.start),
                                                                              booking.interval,
                                                                              daysOfWeek);
                                    break;
                                case "monthly":
                                    recurrence = new Recurrence.MonthlyPattern(DateTime.Parse(booking.start),
                                                                               booking.interval,
                                                                               booking.dayOfMonth);
                                break;
                                case "daily":
                                    recurrence = new Recurrence.DailyPattern(DateTime.Parse(booking.start),
                                                                             booking.interval);
                                    break;
                                default:

                                break;
                            }

                            if (booking.endDate != null)
                            {
                                recurrence.EndDate = DateTime.Parse(booking.endDate);
                            }
                            else if (booking.numberOfOccurrences != -1)
                            {
                                recurrence.NumberOfOccurrences = booking.numberOfOccurrences;
                            }

                            Console.WriteLine(recurrence.HasEnd);

                            // ExchangeService service = GetService(floor);
                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "free",
                                ReasonPhrase = "OK",
                                StatusCode = "200"
                            };
                        } catch (Exception e)
                        {
                            Console.WriteLine(e);
                            return new HttpResponse()
                            {
                                ContentAsUTF8 = "Error",
                                ReasonPhrase = "No",
                                StatusCode = "300"
                            };
                        }
                    }
                }
            };

            HttpServer httpServer = new HttpServer(3000, route_config);
            
            Thread thread = new Thread(new ThreadStart(httpServer.Listen));
            thread.Start();
        }

        private static bool CertificateValidationCallBack(object sender,
                                                          System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                                                          System.Security.Cryptography.X509Certificates.X509Chain chain,
                                                          System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                            (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
        private static ExchangeService GetService(int floor)
        {
            switch (floor)
            {
                case 0:
                case 1:
                    return service10;
                case 2:
                case 6:
                case 8:
                    return service268;
                case 3:
                case 7:
                    return service37;
                case 4:
                case 5:
                case 9:
                    return service459;
                default:
                    return serviceCrestron;
            }
        }
    }
} 

public class API
{
    /*
    private static void GetUserFreeBusy(ExchangeService service, List<AttendeeInfo> attendees, 
                                        int duration, TimeWindow timeWindow)
    {
        // Specify availability options.
        AvailabilityOptions myOptions = new AvailabilityOptions();
        myOptions.MeetingDuration = duration;
        myOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

        // Return a set of free/busy times.
        GetUserAvailabilityResults freeBusyResults = service.GetUserAvailability(attendees,
                                                                                    timeWindow,
                                                                                    AvailabilityData.FreeBusy,
                                                                                    myOptions);

        //XmlDocument xdoc = new XmlDocument();
        //xdoc.LoadXml(free)

        // Display available meeting times.
        Console.WriteLine("Availability for {0}", attendees[0].SmtpAddress);
        Console.WriteLine();

        foreach (AttendeeAvailability availability in freeBusyResults.AttendeesAvailability)
        {
            Console.WriteLine(availability.Result);
            Console.WriteLine();
            foreach (CalendarEvent calendarItem in availability.CalendarEvents)
            {
                Console.WriteLine("Free/busy status: " + calendarItem.FreeBusyStatus);
                Console.WriteLine("Start time: " + calendarItem.StartTime);
                Console.WriteLine("End time: " + calendarItem.EndTime);
                Console.WriteLine();
            }
        }
    }

    private static void GetSuggestedMeetingTimes(ExchangeService service)
    {
        // Create a list of attendees.
        List<AttendeeInfo> attendees = new List<AttendeeInfo>();

        attendees.Add(new AttendeeInfo()
        {
            SmtpAddress = "ic.conf-ott-235queen-169a-conf.ic@canada.ca",
            AttendeeType = MeetingAttendeeType.Organizer
        });

        // Specify suggested meeting time options.
        AvailabilityOptions myOptions = new AvailabilityOptions();
        myOptions.MeetingDuration = 90;
        myOptions.MaximumNonWorkHoursSuggestionsPerDay = 0;
        myOptions.GoodSuggestionThreshold = 49;
        myOptions.MaximumSuggestionsPerDay = 48;
        myOptions.MinimumSuggestionQuality = SuggestionQuality.Excellent;
        myOptions.DetailedSuggestionsWindow = new TimeWindow(DateTime.Now, DateTime.Now.AddDays(2));

        // Return a set of suggested meeting times.
        GetUserAvailabilityResults results = service.GetUserAvailability(attendees,
                                                                        new TimeWindow(DateTime.Now.AddHours(16), DateTime.Now.AddDays(2)),
                                                                            AvailabilityData.Suggestions,
                                                                            myOptions);
        // Display available meeting times.
        Console.WriteLine("Availability for {0}", attendees[0].SmtpAddress);
        Console.WriteLine();

        foreach (Suggestion suggestion in results.Suggestions)
        {
            Console.WriteLine(suggestion.Date);
            Console.WriteLine();
            foreach (TimeSuggestion timeSuggestion in suggestion.TimeSuggestions)
            {
                Console.WriteLine("Suggested meeting time:" + timeSuggestion.MeetingTime);
                Console.WriteLine();
            }
        }
    }
    */
    public ItemId CreateMeeting(ExchangeService service, string room, string subject, string body, DateTime start, DateTime end)
    {
        return CreateMeeting(service, room, subject, body, start, end, new List<string>(), null);
    }

    public ItemId CreateMeeting(ExchangeService service, string room, string subject, string body, DateTime start, DateTime end,
                                List<string> attendees, Recurrence recurrence)
    {
        Appointment appointment = new Appointment(service);
        appointment.Subject = subject;
        appointment.Body = body;
        appointment.Start = start;
        appointment.End = end;
        appointment.Recurrence = recurrence;
        if (attendees != null && attendees.Count != 0)
        {
            foreach (string attendee in attendees)
            {
                appointment.OptionalAttendees.Add(attendee);
            }
        }
        appointment.Save(new FolderId(WellKnownFolderName.Calendar, new Mailbox(room)), SendInvitationsMode.SendToAllAndSaveCopy);
        return appointment.Id;
    }

    public FindItemsResults<Appointment> GetAppointments(ExchangeService service, string emailAddress,
                                                         DateTime start, DateTime end)
    {
        // Room calendar
        CalendarFolder roomCalendar = CalendarFolder.Bind(service,
                                                      new FolderId(WellKnownFolderName.Calendar,
                                                               new Mailbox(emailAddress)));
        CalendarView roomCalendarView = new CalendarView(start, end);
        roomCalendarView.PropertySet = new PropertySet(AppointmentSchema.Start, AppointmentSchema.End); 
        return roomCalendar.FindAppointments(roomCalendarView);
    }
}

public class JarvisRequest
{
    public string room;
    public string start;
    public string end;
    public string subject = "Test Meeting";
    public string body = "Test Body";
    public int floor;
    public string eventID;
    public List<string> attendees = null;
}

public class JarvisRecurrenceRequest
{
    public string room;
    public string start;
    public string end;
    public string subject = "Test Meeting";
    public string body = "Test Body";
    public int floor;
    public string eventID;
    public List<string> attendees = null;

    public string startDate;
    public bool hasEnd = false;
    public string endDate = null;
    public int numberOfOccurrences = -1;
    public string type;
    public int interval;
    public List<int> daysOfWeek;
    public int dayOfMonth;
}

public class JarvisResponse {
    public ItemId eventId = null;
}