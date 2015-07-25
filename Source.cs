using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Mail;
using OpenPop.Mime;
using OpenPop.Pop3;
using System.Threading;

namespace RCM
{
    class Program
    {
        static void Main(string[] args)
        {
            //podatki o email racunu
            string username = "username@gmail.com";
            string password = "passwrd";

            //podatki za sprejem mailov
            string hostname = "pop.gmail.com";
            int port = 995;
            bool useSsl = true;
           
            //podatki za posiljanje mailov
            string outHost = "smtp.gmail.com";
            int outPort = 587;
            bool outSsl = true;

            //tajno geslo, ki mora biti v Subject, da program zazna vsebino
            string secret = "zadeva";

            //izbriši prebrano sporočilo iz serverja (na gmail ni potrebno, ker zazna mail samo enkrat)
            bool deleteMessages = false;

            //interval ponavljanja v ms
            int interval = 60 * 1000;
            /*
             * REPLAY SETTINGS
             * Če zelite, da vam program v odgovor pošlje output terminala,
             * je potrebno pred ukazom vpisati replay:
             * 
             * primer: "ping 192.168.1.1"
             * Če hočemo odgovor bo vsebina maila izgledala "replay:ping 192.168.1.1"
             */


            /*
             * Posiljanje:
             * -na siolu ne deluje,
             * -na gmail je potrebno onemogočiti varno prijavo
            
             * Branje:
             * Deluje na siol.net
             * Problem pri gmail.com -ko enkrat prebere mail, ga naslednič ne zazna več ?!? 
             * Na gmail.com brisanje sporočil ne deluje!
             */
            
            ///////////////////////////////////////////////////////////////////////////////////////////
            ///////////////////////////////////////////////////////////////////////////////////////////
            /*
             * Source... do not change!
             * 
             */
            Console.Beep();
            Console.WriteLine("Start listening on " + username);
            Console.WriteLine("Secret passphrase in Subject is set to: " + secret);
            Console.WriteLine();
            Console.Beep();
            while (true)
            {
                List<Message> sporocila = FetchAllMessages(hostname, port, useSsl, username, password);

                foreach (Message s in sporocila)
                {
                    //preberemo Subject sprorocila
                    string subject = s.Headers.Subject.ToString();

                    if (subject == secret) //ce se ujema z skrivno besedo
                    {
                        string mid = s.Headers.MessageId.ToString().Trim(); //id sporocila
                        string content = s.MessagePart.MessageParts[0].GetBodyAsText().Trim(); //potegnemo vsebino

                        //če obstaja vsebina
                        if (!String.IsNullOrEmpty(content))
                        {
                            bool replay = false;
                            //ugotovimo ali je zahtevan odgovor
                            if (content.Substring(0, 7) == "replay:")
                            {
                                replay = true;
                                content = content.Substring(7);
                            }

                            //izpis podatkov
                            Console.WriteLine("-----------------------------------------------------------------");
                            Console.WriteLine("Requested response: " + replay);
                            Console.WriteLine("Deleteing message: " + deleteMessages);
                            Console.WriteLine("Command to execute: " + content);
                            Console.WriteLine("Message id: " + mid);

                            if (replay)
                            {
                                //execute and save output, then send it back to sender

                                string sender = s.Headers.From.Address;
                                Console.WriteLine("Output sending to:" + sender);

                                string outBody = runCMD_Feedback(content);
                                Console.WriteLine();
                                Console.WriteLine("Output:" + outBody);
                                bool send_status = SendOutputBackToSender(username, password, sender, outHost, outPort, outBody, outSsl);
                                Console.WriteLine();
                                Console.WriteLine("Send status: " + send_status);
                            }
                            else
                            {
                                runCMD_NoFeedback(content);
                            }
                        }

                        if (deleteMessages)
                        {
                            bool delete = DeleteMessageByMessageId(hostname, port, useSsl, username, password, mid);
                            Console.WriteLine("Delete status: " + delete);
                        }
                    }
                }
                //Console.WriteLine("-------------------------------------------------");                
                Thread.Sleep(interval);
            }
        }

        public static List<Message> FetchAllMessages(string hostname, int port, bool useSsl, string username, string password)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();

                // We want to download all messages
                List<Message> allMessages = new List<Message>(messageCount);

                // Messages are numbered in the interval: [1, messageCount]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number
                for (int i = messageCount; i > 0; i--)
                {
                    allMessages.Add(client.GetMessage(i));
                }

                // Now return the fetched messages
                return allMessages;
            }
        }

        public static bool DeleteMessageByMessageId(string hostname, int port, bool useSsl, string username, string password, string messageId)
        {
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);


                // Get the number of messages on the POP3 server
                int messageCount = client.GetMessageCount();

                // Run trough each of these messages and download the headers
                for (int messageItem = messageCount; messageItem > 0; messageItem--)
                {
                    // If the Message ID of the current message is the same as the parameter given, delete that message
                    if (client.GetMessageHeaders(messageItem).MessageId == messageId)
                    {
                        // Delete
                        client.DeleteMessage(messageItem);
                        return true;
                    }
                }
                // We did not find any message with the given messageId, report this back
                return false;
            }
        }

        public static bool SendOutputBackToSender(string username, string password, string sender, string outHost, int outPort, string outBody,bool enableOutSsl)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(outHost);

                mail.From = new MailAddress(username);
                mail.To.Add(sender);
                mail.Subject = "AutoResponseRCM";
                mail.Body = outBody;

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(username, password);
                SmtpServer.EnableSsl = enableOutSsl;

                SmtpServer.Send(mail);
                return true;
            }
            catch //(Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                return false;
            }
        }

        public static string runCMD_Feedback(string command)
        {
            string output = "";

            ProcessStartInfo startInfo = new ProcessStartInfo("cmd", "/c " + command)
            {
                WindowStyle = ProcessWindowStyle.Hidden,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            Process process = Process.Start(startInfo);
            //process.OutputDataReceived += (sender, e) => Console.WriteLine(e.Data);
            process.OutputDataReceived += (sender, e) => output += e.Data;

            process.BeginOutputReadLine();
            process.Start();
            process.WaitForExit();

            return output;
        }

        public static void runCMD_NoFeedback(string command)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Arguments = "/C "+ command;
            process.StartInfo = startInfo;
            process.Start();
        }
    }
}
