using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace TestEwsNtlm
{
    internal class Program
    {
        static void Main()
        {
            Console.Write("URL: ");
            string url = Console.ReadLine();
            Console.Write("Benutzername: ");
            string username = Console.ReadLine();
            Console.Write("Passwort: ");
            string password = ReadPassword();
            Console.Write("Domaine: ");
            string domain = Console.ReadLine();

            // Create the ExchangeService instance
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            // Set credentials for NTLM authentication
            service.Credentials = new NetworkCredential(username, password, domain);

            // Set the Exchange server URL
            service.Url = new Uri(url);

            // Configure NTLM authentication explicitly
            service.PreAuthenticate = true;
            service.UseDefaultCredentials = false;
            service.TraceEnabled = true;

            try
            {
                // Get calendar folder
                Folder calendar = Folder.Bind(service, WellKnownFolderName.Calendar);
                Console.WriteLine($"Calendar: {calendar.DisplayName}");

                // Find appointments in the next 7 days
                DateTime startDate = DateTime.Now;
                DateTime endDate = startDate.AddDays(7);

                CalendarView calView = new CalendarView(startDate, endDate);
                FindItemsResults<Appointment> appointments = service.FindAppointments(
                    WellKnownFolderName.Calendar, calView);

                Console.WriteLine($"\nUpcoming appointments ({appointments.Items.Count}):");
                foreach (Appointment appt in appointments)
                {
                    Console.WriteLine($"- {appt.Subject}");
                    Console.WriteLine($"  Start: {appt.Start}");
                    Console.WriteLine($"  End: {appt.End}");
                    Console.WriteLine($"  Location: {appt.Location}\n");
                }

                // Create a new appointment
                Appointment newAppt = new Appointment(service);
                newAppt.Subject = "Team Meeting";
                newAppt.Body = "Discuss project status";
                newAppt.Start = DateTime.Now.AddDays(1).Date.AddHours(10);
                newAppt.End = newAppt.Start.AddHours(1);
                newAppt.Location = "Conference Room A";
                newAppt.Save(SendInvitationsMode.SendToNone);

                Console.WriteLine("Appointment created successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static string ReadPassword()
        {
            string password = "";
            ConsoleKey key;

            do
            {
                var keyInfo = Console.ReadKey(intercept: true);
                key = keyInfo.Key;

                if (key == ConsoleKey.Backspace && password.Length > 0)
                {
                    // remove one character
                    password = password.Substring(0, password.Length - 1);

                    // remove * from console
                    Console.Write("\b \b");
                }
                else if (!char.IsControl(keyInfo.KeyChar))
                {
                    // add the char to the password
                    password += keyInfo.KeyChar;

                    // print *
                    Console.Write("*");
                }

            } while (key != ConsoleKey.Enter);

            Console.WriteLine(); // move to next line
            return password;
        }
    }
 }
