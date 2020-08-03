using CommandLine;
using CommandLine.Text;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CleanupMyMailbox
{
    class Program
    {
        public class Options
        {
            [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
            public bool Verbose { get; set; }

            [Option('e', "email", Required = true, HelpText = "Email address")]
            public string Email { get; set; }

            [Option('m', "mailbox", Required = false, HelpText = "Path to the mailbox that needs to be cleaned. Use a forward slash path separator. For example: INBOX/Automated/Services")]
            public string MailboxPath { get; set; }

            [Option('d', "defaultcredentials", Required = false, HelpText = "Use default credentials. Use this option for hosted Exchange instances where the user is logged into the same domain.")]
            public bool UseDefaultCredentials { get; set; }

            [Option('u', "username", Required = false, HelpText = "Username for authenticating to the Exchange server. Leave blank if the same as your email address.")]
            public string Username { get; set; }

            [Option('s', "password", Required = false, HelpText = "Password for authenticating to the Exchange server. If required and left blank the program will prompt for a password.")]
            public string Password { get; set; }

            [Option('a', "age", Required = true, HelpText = "The age of emails after which they should be deleted. This should be a string that the .NET TimeSpan.Parse() method will recognize. For example: 1 = 1 day, 6:30 = 6 hours and 30 minutes")]
            public TimeSpan Age { get; set; }

            [Option('x', "deletemode", Required = false, HelpText = "Sets the Delete mode to Hard Delete. Emails will not go into Deleted Items folders")]
            public bool DeleteMode { get; set; }

            [Option('n', "noprompt", Required = false, HelpText = "Do not prompt before deleting emails")]
            public bool NoPrompt { get; set; }



            public override string ToString()
            {
                StringBuilder sb = new StringBuilder();
                string sLineFormat = "{0,-22}{1}\r\n";
                sb.AppendFormat(sLineFormat, "Verbose", Verbose);
                sb.AppendFormat(sLineFormat, "Email", Email);
                sb.AppendFormat(sLineFormat, "Mailbox", MailboxPath);
                sb.AppendFormat(sLineFormat, "UseDefaultCredentials", UseDefaultCredentials);
                sb.AppendFormat(sLineFormat, "Age", Age);
                sb.AppendFormat(sLineFormat, "No Prompt", NoPrompt);
                return sb.ToString();
            }
        }

        private Options _options;
        private ExchangeService _service;

        public Program(Options options)
        {
            _options = options;
        }

        public void Run()
        {
            Echo(ProgramHeader());
            Echo(string.Format("OPTIONS:\r\n{0}{1}\r\n", _options, "".PadLeft(80, '*')), true);

            _service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            _service.UseDefaultCredentials = _options.UseDefaultCredentials;
            if (!_options.UseDefaultCredentials)
            {
                string sUsername = (!string.IsNullOrWhiteSpace(_options.Username) ? _options.Username : _options.Email);
                string sPassword = _options.Password;
                if (string.IsNullOrWhiteSpace(_options.Password))
                {
                    sPassword = GetPassword();
                }
                _service.Credentials = new WebCredentials(sUsername, sPassword);
            }

            _service.AutodiscoverUrl(_options.Email, RedirectionUrlValidationCallback);
            Echo(string.Format("AutoDixcovered WS URL : {0}", _service.Url), true);

            Echo(string.Format("Cleaning up mailbox : {0}", _options.MailboxPath));

            List<string> pathParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(_options.MailboxPath))
            {
                pathParts.AddRange(_options.MailboxPath.Split(new char[] { '/' }));

            }
            if (0 == pathParts[0].ToUpper().CompareTo("INBOX"))
            {
                pathParts.RemoveAt(0);
            }

            Folder fCleanup = GetFolder(pathParts);
            Echo(string.Format("Found [{0}] items at path [{1}]", fCleanup.TotalCount, (string.IsNullOrWhiteSpace(_options.MailboxPath) ? "INBOX" : _options.MailboxPath)), true);

            Echo(string.Format("Finding items of minimum age [{0}]", _options.Age));
            FindItemsResults<Item> oldItems = fCleanup.FindItems(new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.Now.Subtract(_options.Age)), new ItemView(10000));
            Echo(string.Format("Found [{0}] old items", oldItems.TotalCount));

            bool bContinue = _options.NoPrompt;
            if (!_options.NoPrompt)
            {
                Console.Write("Continue deleting [{0}] emails? [Y/n] : ", oldItems.TotalCount);
                string sAnswer = Console.ReadLine();
                bContinue = (!string.IsNullOrWhiteSpace(sAnswer) && sAnswer.ToUpper().StartsWith("Y"));
            }

            if (bContinue)
            {
                Echo("Deleting: ");
                int i = 1;
                foreach (Item item in oldItems)
                {
                    if (_options.Verbose)
                    {
                        Echo(string.Format("    [{0}] [{1}] - {2}", i++, item.DateTimeReceived, item.Subject));
                    }
                    else
                    {
                        Console.Write(".");
                    }
                    if (_options.DeleteMode)
                    {
                        item.Delete(DeleteMode.HardDelete);
                    }
                    else
                    {
                        item.Delete(DeleteMode.MoveToDeletedItems);
                    }

                }

                Echo((_options.Verbose ? "Finished" : "\r\nFinished"));
            }
        }

        private Folder GetFolder(List<string> path)
        {
            FolderId parentId = new FolderId(WellKnownFolderName.Inbox, _options.Email);
            // First get the INBOX
            Echo("Binding to INBOX", true);

            Folder fInbox = Folder.Bind(_service, parentId);

            if (null != path && 0 < path.Count)
            {
               Folder result = null;

                foreach (string sPath in path)
                {
                    Echo(string.Format("Binding to folder [{0}]", sPath), true);
                    result = _service.FindFolders(parentId, new SearchFilter.IsEqualTo(FolderSchema.DisplayName, sPath), new FolderView(1)).FirstOrDefault();
                    parentId = result.Id;
                }

                // Return the folder
                return result;
            }

            // Just return the INBOX
            return fInbox;
        }

        private string GetPassword()
        {
            Console.Write("Please Enter Password for Email {0}: ", _options.Email);
            string pass = "";
            do
            {
                ConsoleKeyInfo key = Console.ReadKey(true);
                // Backspace Should Not Work
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    pass += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                    {
                        pass = pass.Substring(0, (pass.Length - 1));
                        Console.Write("\b \b");
                    }
                    else if (key.Key == ConsoleKey.Enter)
                    {
                        break;
                    }
                }
            } while (true);
            Console.WriteLine("");
            return pass;
        }

        private bool RedirectionUrlValidationCallback(string redirectionUrl)
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

        private void Echo(string message, bool verbose = false)
        {
            if (_options.Verbose || !verbose)
            {
                Console.WriteLine(message);
            }
        }

        static void Main(string[] args)
        {
            var parser = new CommandLine.Parser(with => with.HelpWriter = null);
            var parserResult = parser.ParseArguments<Options>(args);
            parserResult
                .WithParsed<Options>(options => new Program(options).Run())
                .WithNotParsed(errs => DisplayHelp(parserResult, errs));

#if DEBUG
            Console.Write("Press [Enter] to continue...");
            Console.ReadLine();
#endif
        }

        static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<Error> errs)
        {
            var helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                h.Heading = ProgramHeader();
                h.Copyright = string.Empty;
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
            Console.WriteLine(helpText);
        }

        static string ProgramHeader()
        {
            return string.Format(
                "\r\n{0}\r\n{1} {2}\r\n{3}\r\n{0}",
                "".PadLeft(80, '*'),
                "Cleanup My Mailbox",
                "1.0.0",
                string.Format("Copyright (c) {0} Jonathan Franzone", DateTime.Now.Year)
            );
        }
    }
}
