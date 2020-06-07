using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static OutlookConsole.GraphClient;

namespace OutlookConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var graphClient = new GraphClient();
            var Token = graphClient.GetAuthToken2();

            bool anotherRequest = true;
            Token.Wait();
            while (anotherRequest)
            {
                Console.WriteLine("Choose a number:");
                Console.WriteLine("1. Get your info from Microsoft");
                Console.WriteLine("2. Get your last 10 emails");
                Console.WriteLine("3. Write an email");
                ConsoleKeyInfo choice = Console.ReadKey();
                Console.WriteLine();
                Console.WriteLine();
                switch (choice.KeyChar)
                {
                    case '1':
                        Task<string> users = graphClient.GetThisUser();
                        users.Wait();
                        Console.WriteLine(users.Result);
                        Console.WriteLine();
                        Console.WriteLine();
                        break;
                    case '2': 
                        Task<List<SimpleEmail>> emails = graphClient.GetTenEmails();
                        emails.Wait();
                        foreach (SimpleEmail email in emails.Result)
                        {
                            Console.WriteLine("Subject:   " + email.Subject);
                            Console.WriteLine();
                            Console.WriteLine("From:      " + email.FromAddress);
                            Console.WriteLine();
                            Console.WriteLine("Date:      " + email.Date);
                            Console.WriteLine();
                            Console.WriteLine("Email ID:  " + email.Id);
                            Console.WriteLine();
                            Console.WriteLine(email.Body);
                            Console.WriteLine();
                            Console.WriteLine();
                        }

                        Console.ReadLine();
                        Console.WriteLine();
                        Console.WriteLine();
                        break;
                    case '3':
                        //Console.Write("To? (enter a valid email address?)");
                        //string toEmail = Console.ReadLine();
                        //Console.Write("Subject?");
                        //string subject = Console.ReadLine();
                        //Console.WriteLine("Message body?");
                        //string body = Console.ReadLine();
                        Task<bool> draftCreated = graphClient.SendEmail();
                        Console.WriteLine();
                        Console.WriteLine();
                        break;
                }
                
            }
        }
    }
}

//get user, after getting token
//var client = new RestClient("https://graph.microsoft.com/v1.0/me/");
//client.Timeout = -1;
//var request = new RestRequest(Method.GET);
//request.AddHeader("Authorization", "Bearer EwBgA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAbeXxyeJGh5SJcmBf+52EI5xj+Ve96BXDR5fqTIwqv63zp9Bv9LQAs2r+eofk6866yBNHv7qs95tMKnN7ypZfV3W8TfzrBbURuufNdhyj/Sg/u9r/NDLgj/LYhfXt7cmCa8VpT6TOM8wq9DYC+1KU8VwdzB+u0PCYt5t8muhg7Jd+n6u9aqqgIyjUkydle0BfEPfVEnYormF4xrhkEx3WFOVjkLqPL0k4pfKCyRSgf+QC9j6LYHygPZDeEkv9wlwoP0Rgb0L1lqGQwQ6y5sFeRVw907NeRWFEAlsg1RnpYJiruwXOhRsMmDPuCS+lsj4rz1Ju2zhf9hZLJZuxC5yb2oDZgAACO/9jqsiP0JAMAL9ZC9jvFAZSXC98pc7QUkCIXyBCxfqxPx9uhALQqOLIu3UXQXnr+SDqrSWEavkSGAPLDQiQ2spYkKkX85u2LsP8IVrC4qnSSEiVY9MeIiCyJmomu1mTRAKJJ8ojc5vnJnxpBQPMUUHcAW6gsqw0KQYD3/dgMBmCatRp5d2bn3rj+v1ekDcJmsaf2QnRwulnmz42a+jgccleCZmK3u3Mbj9K/MGMes4ZHnDmsIy+2bXtah6allgbSqTfLgzubFEnKRWGY6yEHvHbd3rTRy7zorWqkJQHHGHp+s/wob7tOUn9T3+mh5m6eyTuWjWwBUdxL93HN2xQMhfDkWfFGNXd+ADiK2saDEYy5oVNtbdOPXCe1OF+sXWIlxLFw2RtQwGsqXRIeaD0OVDkODuKMfo9IuGa1bBnRq2KVWzWmH2vBPPtyWgBZgw1WSdi/cCSvxoy5KETrI+Flpm/35oyqvTlUEQKad0sWtYwVJCFJTsBWmFrZd54bwD2OdHoUDvoCKU5snvElzpf20ig/K9muZZsb7C2oajPK7b7wmYK9j/gfjMuhRp8YA6C7sjlorAyRQSRasbyKK/uaB6pN6qu9ubmX1CYFpHSvV792wRsSCPDBJMdvFvgJa1xz1baP7zo30D/WX7nJJrlHYM4zZSHQahofM+FD+weaNCRGIQ5IRgh7MaTVLLjO9vmuRoVM9SIwoRZwnDxOzemvCyjXQNk8bMHMSeSCCm4dvZOagHi5eOXR8k9nAC");
//IRestResponse response = client.Execute(request);
//Console.WriteLine(response.Content);


//Get 10 messages
//var client = new RestClient("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages");
//client.Timeout = -1;
//var request = new RestRequest(Method.GET);
//request.AddHeader("Authorization", "Bearer EwBgA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAeDUcwVPwJ5Q5Elbp2Wt2bholzpKhB+uuYZlneGheJ+SJqL1pyo667vCKc2/1yGl/Y81V+LxlHPtjNnMU41kfQd4M3krXnBIUCj4jw51Rj3wNOOmma8reSKLELIHIdkZLQLRazxUNDk/BkOhJvs5kUC81eWA+UNrGa3MJml3DIKSjzr1wexaAAlw0jij0bg51l6YwurvlHM5VpHj6BRmNEI3NKQtVgl9JCQNYyw1485jW/gJsW5xUZyvyg84j87/b6BcUjlYeCcbQLimn1uEye3FZYNz0ffoF2zCdsNqCk29qNCqwDAAIrypHNDZ5hhdLYfCs10I7UXRrVpuj8N5J2QDZgAACKNo2YZONdfUMAKUxIoi8Qs8P67iEB0eIY6zM/fvIKZTI3tRN9WQiEZn8kNeY4vQn5xIeTdpjv5n9Vhh4OBB1w1y/ZvjYPlQVZEc6CiQZRHT8KKoaUyuLKNrTajqDmy1HNM5jt2mBBNBIcEMfZ4hcSwdQpOBvzNTDfcO7YXw07i4BBejjq2Z2P+lF0XhIn7ict+TNpqLxPm5/HdWySCAlkjHzDePeI80zd0AyRU1X5FeGxv+DkEgm3F777TqstPEcOFQ9RJajBB24dLXBtGfhz0g2r1D6iDMXuxmzG1CSUSlWpmuyW1Pvvo2QxzZNobd7VAQUlPTZGmxmw/LLG+Ml21m5RczVD6tc0iykKzGof2+R+QYlPoQCfo9S61W+56HwzFRzg8OkI80JfY+T2nE/drgFj1DAkXkh7mNUfmuzD1o9h+fDoRc6esKY62Tyqxm25ukuSmwZTFsT4iNek6X8jFylrG7H996KfCK1zgAZDH+z0k01skBwl7euCYCNL7t2ns7Q+f2DdPXM5zF4td07vUXOev1tL7QEZ/Z3Bukp4v3ORk9vWGq9JxOCWTmgWJL5gmMAaPYC2S2FwFAc797wUx8Ga9INCB93c1GZObYPFiX9O/ysFLZFAN4Ub+E/IiX8Et/S9YxzVSVekZhgiazL0HspT+fqpuDewIWqtx7umQzOjOn2Q9DVm0FKtSHxmhBP22in4Jr3uwlcPnB0XhlVbEbzPFLcMiP8cJUulmw0TaLkNPwqwpXw/fjR3QC");
//IRestResponse response = client.Execute(request);
//Console.WriteLine(response.Content);

//