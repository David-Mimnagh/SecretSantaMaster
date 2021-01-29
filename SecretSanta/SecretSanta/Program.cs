using SecretSanta.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;


namespace SecretSanta
{
    class Program
    {
        static List<Participant> GetAllParticipants()
        {
            List<Participant> p = new List<Participant>();

            string[] info = File.ReadAllLines("../../Files/participantList.csv");

            foreach ( var line in info )
            {
                string[] infoIndiv = line.Split(',');
                if ( infoIndiv[0] != "id" )
                {
                    p.Add(new Participant { Id = Convert.ToInt32(infoIndiv[0]), Name = infoIndiv[1], SpouseId = Convert.ToInt32(infoIndiv[2]), EmailAddress = infoIndiv[3], Interests = infoIndiv[4], PreviousSSantaId = Convert.ToInt32(infoIndiv[5]), ShirtSize = infoIndiv[6] });
                }
            }

            return p;
        }

        static Random rnd = new Random();
        static List<Dictionary<Participant, Participant>> GetSecretSantas(List<Participant> p)
        {
            List<Dictionary<Participant, Participant>> pP = new List<Dictionary<Participant, Participant>>();


            var availablelist = new List<Participant>();

            foreach ( var pers in p )
            {
                availablelist.Add(pers);
            }
            List<int> addedList = new List<int>();


            foreach ( var person in p )
            {
                //repeat while person not in addedList
                while ( !addedList.Contains(person.Id) )
                {
                    //look at available list
                    int r = rnd.Next(availablelist.Count);

                    //get a random person
                    var possiblePair = availablelist.ElementAt(r);

                    //make sure its not spouse
                    if ( person.Id != possiblePair.Id )
                    {
                        if ( person.SpouseId != possiblePair.Id )
                        {
                            if (person.PreviousSSantaId != possiblePair.Id)
                            {
                                // add to pP
                                var newAddition = new Dictionary<Participant, Participant>();
                                newAddition.Add(person, possiblePair);
                                pP.Add(newAddition);
                                // remove from avail list
                                addedList.Add(person.Id);
                                availablelist.RemoveAt(r);
                            }
                        }
                    }
                }
            }


            return pP;

        }

        static string BuildInterests(Participant part)
        {
            string interests = "";
            var interestListBefore = part.Interests.Split('-').ToList();
            var interestListAfter = new List<string>();
            foreach ( var i in interestListBefore )
            {
                var interestString = "";
                for ( int j = 0; j < i.Length; j++ )
                {
                    if ( j == 0 )
                    {
                        interestString += i[j].ToString().ToUpper();
                    }
                    else
                    {
                        if ( i[j] == ' ' )
                        {
                            interestString += "+";
                        }
                        else
                        {
                            if ( interestString[j-1] == '+' )
                            interestString += i[j].ToString().ToUpper();
                            else
                            interestString += i[j].ToString();
                        }
                    }
                }
                interestListAfter.Add(interestString);
            }
            string baseAmazon = "https://www.amazon.co.uk/s/ref=nb_sb_noss_2?url=search-alias%3Daps&field-keywords=";
            foreach ( var i in interestListAfter )
            {
                string copyString = i.Replace('+', ' ');
                interests += "<li style = \"padding-top: 5px\"><a href = \"" + (baseAmazon + i) + "\"target=\"_blank\"style = \"color: white; \">" + copyString + "</a></li>";
            }

            return interests;
        }
        static string BuildEmailHTML(Participant p)
        {
            string[] info = File.ReadAllLines("../../Files/basehtml.txt");
            string baseHTMLString = "";
            foreach (var line in info)
            {
                baseHTMLString += line;
            }
            baseHTMLString = baseHTMLString.Replace("%%USERNAME%%", p.Name).Replace("%%TOPSIZE%%", p.ShirtSize); ;
            baseHTMLString += BuildInterests(p);
            baseHTMLString += "<p style = \"padding-top: 30px\"> This year, there is a <b style='padding-top: 10px; font-size:32px;'>£25</b> limit. Don't go over, and try not to be under by too much!<br/><br/>Merry Christmas!</p></ul></div></div></body></html>";
            return baseHTMLString;
        }

        static void SendOutEmails(List<Dictionary<Participant, Participant>> pP)
        {
            for ( int i = 0; i < pP.Count; i++ )
            {
                foreach ( var p in pP[i] )
                {
                    MailMessage mail = new MailMessage("secret_santa@family.com", p.Key.EmailAddress);
                   //MailMessage mail = new MailMessage("secret_santa@family.com", "davidmimnagh1@googlemail.com");
                    SmtpClient client = new SmtpClient();
                    client.Port = 587;
                    client.Host = "smtp.gmail.com";
                    client.EnableSsl = true;
                    client.Timeout = 10000;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(Properties.Settings.Default.UName, Properties.Settings.Default.Pass);

                    mail.Subject = p.Key.Name + ", your secret santa has been selected!";
                    mail.Body = BuildEmailHTML(p.Value).Replace("%%USER%%", p.Key.Name);
                    mail.IsBodyHtml = true;
                    mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                    client.Send(mail);
                }
            }



        }
        static void CreateNewCSV(List<Dictionary<Participant, Participant>> participantsPair)
        {
            string csv = File.ReadAllText("../../Files/basecsv.txt");
            for (int i = 0; i < participantsPair.Count; i++)
            {
                foreach (var pair in participantsPair[i])
                {
                    var participant = pair.Key;
                    var match = pair.Value;
                    csv += $"{participant.Id.ToString()},{participant.Name},{participant.SpouseId.ToString()},{participant.EmailAddress},{participant.Interests},{match.Id.ToString()},{participant.ShirtSize}{Environment.NewLine}";
                }
            }
            
            File.WriteAllText("../../Files/participantList_NEW.csv",csv);
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Getting Participants...");
            List<Participant> participants = GetAllParticipants();
            Console.WriteLine("Getting Pairs...");
            List<Dictionary<Participant, Participant>> partiPair = GetSecretSantas(participants);

            Console.WriteLine("Sending emails...");
            SendOutEmails(partiPair);
            Console.WriteLine("Finished sending emails");

            Console.WriteLine("Creating new CSV file...");
            CreateNewCSV(partiPair);

            Console.WriteLine("Secret Santa has completed...");
            Console.ReadLine();
        }
    }
}
