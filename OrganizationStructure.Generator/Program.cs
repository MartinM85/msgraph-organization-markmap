using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.Security;
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
            if (args.Length != 4)
            {
                throw new ArgumentOutOfRangeException(nameof(args));
            }
            var tenantId = args[0];
            var clientId = args[1];
            var clientSecret = args[2];
            var outputFolder = args[3];

            var users = await GetUsersAsync(tenantId, clientId, clientSecret);
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
            var locations = users.GroupBy(x => x.OfficeLocation);

            using var file = new StreamWriter(Path.Combine(outputFolder, "OfficeLocations.md"));
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("markmap:");
            await file.WriteLineAsync("  colorFreezeLevel: 6");
            await file.WriteLineAsync("---");
            await file.WriteLineAsync("# Office locations");
            foreach (var location in locations)
            {
                await file.WriteLineAsync($"## {location.Key}");
                foreach (var user in location)
                {
                    await file.WriteLineAsync($"- {user.DisplayName}");
                    await file.WriteLineAsync($"  {user.Department} ({user.JobTitle})");
                }
            }
        }

        private static async Task<List<User>> GetUsersAsync(string tenantId, string clientId, string clientSecret)
        {
            var clientSecretCredentials = new ClientSecretCredential(tenantId, clientId, clientSecret);
            var graphClient = new GraphServiceClient(clientSecretCredentials);

            // use filter if you need to exclude guest users, etc.
            var response = await graphClient.Users.GetAsync(x =>
            {
                x.QueryParameters.Select = new[] { "id", "displayName", "givenName", "mail", "jobTitle", "officeLocation", "city", "department" };
                x.QueryParameters.Expand = new[] { "manager($select=id)" };
            });

            var users = new List<User>();
            var pageIterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(graphClient, response, (user) =>
            {
                users.Add(user);
                return true;
            });

            await pageIterator.IterateAsync();
            return users;
        }
    }
}