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
                    p.Add(new Participant { Id = Convert.ToInt32(infoIndiv[0]), Name = infoIndiv[1], SpouseId = Convert.ToInt32(infoIndiv[2]), EmailAddress = infoIndiv[3], Interests = infoIndiv[4], PreviousSSantaId = Convert.ToInt32(infoIndiv[5]) });
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
            baseHTMLString = baseHTMLString.Replace("%%USERNAME%%", p.Name);
            baseHTMLString += BuildInterests(p);
            baseHTMLString += "<p style = \"padding-top: 30px\"> REMEMBER EVERYONE - <br/>There is a £20 limit. Don't go over, and try not to be under by too much!<br/>Merry Christmas!</p></ul></div></div></body></html>";
            return baseHTMLString;
        }

        static void SendOutEmails(List<Dictionary<Participant, Participant>> pP)
        {
            for ( int i = 0; i < pP.Count; i++ )
            {
                foreach ( var p in pP[i] )
                {
                    //MailMessage mail = new MailMessage("secret_santa@family.com", p.Key.EmailAddress);
                   MailMessage mail = new MailMessage("secret_santa@family.com", "davidmimnagh1@googlemail.com");
                    SmtpClient client = new SmtpClient();
                    client.Port = 587;
                    client.Host = "smtp.gmail.com";
                    client.EnableSsl = true;
                    client.Timeout = 10000;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(SecretSanta.Properties.Settings.Default.UName, SecretSanta.Properties.Settings.Default.Pass);

                    mail.Subject = p.Key.Name + ", your secret santa has been selected!";
                    mail.Body = BuildEmailHTML(p.Value);
                    mail.IsBodyHtml = true;
                    mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                    client.Send(mail);
                }
            }



        }

        static void Main(string[] args)
        {
            List<Participant> participants = GetAllParticipants();

            List<Dictionary<Participant, Participant>> partiPair = GetSecretSantas(participants);

            SendOutEmails(partiPair);

            Console.WriteLine("Done sending email");
            Console.ReadLine();
        }
    }
}
