using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace OrganizationStructure.Generator
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            if (args == null)
            {
                throw new ArgumentNullException(nameof(args));
            }
            if (args.Length != 2 || args.Length != 4)
            {
                throw new ArgumentOutOfRangeException(nameof(args));
            }

            string outputFolder;

            GraphServiceClient client;
            // create client from existing access token
            if (args.Length == 2)
            {
                var accessToken = args[0];
                outputFolder = args[1];
                client = GraphServiceClientFactory.CreateClientFromToken(accessToken);
            }
            else
            {
                var tenantId = args[0];
                var clientId = args[1];
                var clientSecret = args[2];
                outputFolder = args[3];
                client = GraphServiceClientFactory.CreateClientFromClientSecretCredential(tenantId, clientId, clientSecret);
            }

            var users = await GetUsersAsync(client);
            await GroupByManagerAsync(users, outputFolder);
            await GroupByDepartmentAsync(users, outputFolder);
            await GroupByJobTitleAsync(users, outputFolder);
            await GroupByOfficeLocationAsync(users, outputFolder);
        }

        private static async Task GroupByManagerAsync(List<User> users, string outputFolder)
        {
            var managerUsers = users.Where(x => x.Manager != null).GroupBy(x => x.Manager.Id);

            using var file = new StreamWriter(Path.Combine(outputFolder, "Users.md"));
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("markmap:");
            await file.WriteLineAsync("  colorFreezeLevel: 6");
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("# Organization");
            foreach (var manager in users.Where(x => x.Manager == null))
            {
                await file.WriteLineAsync($"## {manager.DisplayName}");
                await WriteUsersAsync(manager.Id, 0);
            }

            async Task WriteUsersAsync(string managerId, int level)
            {
                var mgUsers = managerUsers.FirstOrDefault(x => x.Key == managerId);
                if (mgUsers == null || !mgUsers.Any())
                {
                    return;
                }
                var indent = new string(' ', level);
                foreach (var item in mgUsers)
                {
                    await file.WriteLineAsync($"{indent}- {item.DisplayName}");
                    await file.WriteLineAsync($"{indent}  {item.JobTitle} ({item.Department})");
                    await WriteUsersAsync(item.Id, level + 2);
                }
            }
        }

        private static async Task GroupByDepartmentAsync(List<User> users, string outputFolder)
        {
            var departments = users.GroupBy(x => x.Department);

            using var file = new StreamWriter(Path.Combine(outputFolder, "Departments.md"));
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("markmap:");
            await file.WriteLineAsync("  colorFreezeLevel: 6");
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("# Departments");
            foreach (var department in departments)
            {
                await file.WriteLineAsync($"## {department.Key}");
                foreach (var user in department)
                {
                    await file.WriteLineAsync($"- {user.DisplayName}");
                    await file.WriteLineAsync($"  {user.JobTitle} ({user.OfficeLocation})");
                }
            }
        }

        private static async Task GroupByJobTitleAsync(List<User> users, string outputFolder)
        {
            var jobTitles = users.GroupBy(x => x.JobTitle);

            using var file = new StreamWriter(Path.Combine(outputFolder, "JobTitles.md"));
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("markmap:");
            await file.WriteLineAsync("  colorFreezeLevel: 6");
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("# Job titles");
            foreach (var jobTitle in jobTitles)
            {
                await file.WriteLineAsync($"## {jobTitle.Key}");
                foreach (var user in jobTitle)
                {
                    await file.WriteLineAsync($"- {user.DisplayName}");
                    await file.WriteLineAsync($"  {user.Department} ({user.OfficeLocation})");
                }
            }
        }

        private static async Task GroupByOfficeLocationAsync(List<User> users, string outputFolder)
        {
            var countries = users.GroupBy(x => x.Country);
            

            using var file = new StreamWriter(Path.Combine(outputFolder, "OfficeLocations.md"));
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("markmap:");
            await file.WriteLineAsync("  colorFreezeLevel: 6");
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("# Office locations");

            foreach (var country in countries)
            {
                await file.WriteLineAsync($"### {country.Key}");
                var cities = country.GroupBy(x => x.City);
                foreach (var city in cities)
                {
                    await file.WriteLineAsync($"#### {city.Key}");
                    var locations = city.GroupBy(x => x.OfficeLocation);
                    foreach (var location in locations)
                    {
                        await file.WriteLineAsync($"##### {location.Key}");
                        foreach (var user in location)
                        {
                            await file.WriteLineAsync($"- {user.DisplayName}");
                            await file.WriteLineAsync($"  {user.Department} ({user.JobTitle})");
                        }
                    }
                }
            }
        }

        private static async Task<List<User>> GetUsersAsync(GraphServiceClient graphClient)
        {
            // use filter if you need to exclude guest users, etc.
            var response = await graphClient.Users.GetAsync(x =>
            {
                x.QueryParameters.Select = new[] { "id", "displayName", "givenName", "mail", "jobTitle", "officeLocation", "city", "country", "department" };
                x.QueryParameters.Expand = new[] { "manager($select=id)" };
            });

            var users = new List<User>();
            var pageIterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(graphClient, response, (user) =>
            {
                users.Add(user);
                return true;
            });
            await pageIterator.IterateAsync();

            // filtering users
            //users = users.Where(x => x.Mail != null && x.Mail.EndsWith("contoso.com")).ToList();
            //users = users.Where(x => x.GivenName != null).ToList();

            //foreach (var name in new[] { "Tool" })
            //{
            //    users = users.Where(x => x.GivenName != name).ToList();
            //}
            //foreach (var email in new[] { "johndoec@contoso.com" })
            //{
            //    users = users.Where(x => x.Mail != email).ToList();
            //}

            return users;
        }
    }
}