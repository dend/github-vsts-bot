// Docs Feedback Processor
// (c) Den Delimarsky, 2018. All rights reserved.

#r "Newtonsoft.Json"

using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;

public static async Task Run(dynamic payload, TraceWriter log)
{
    List<string> approvedUsers = new List<string>{"dend", "thedanfernandez", "powerhelmsman", "meganbradley"};

    log.Info("Received a payload.");

    // Process only if a new comment is created.
    if (payload.action != "created")
    {
        return;
    }

    if (payload.comment != null)
    {
        // Don't process your own comments, and check against an approve-list of users who can create customer feedback.
        if (payload.comment.user.login.ToString().ToLower() != "botcrane" && approvedUsers.Contains(payload.comment.user.login.ToString().ToLower()))
        {
            if (!string.IsNullOrWhiteSpace(payload.comment.body.ToString().ToLower()) && !payload.comment.body.ToString().ToLower().Contains("#log-suggestion"))
            {
                return;
            }
            else
            {
                var operationalBody = payload.comment.body.ToString().ToLower();
                string microsoftId = string.Empty;
                
                var regex = new Regex(@"[\@].\S*");
                var match = regex.Match(operationalBody);
                if (match != null)
                {
                    log.Info("Found a tagged GitHub ID: " + match.Value.ToString());
                    var cleanGitHubId = match.Value.ToString().Replace("@", "");
                    log.Info("Clean ID: " + cleanGitHubId);

                    microsoftId = await ResolveGitHubAliasToIdentity(cleanGitHubId, log);

                    log.Info("Discovered Microsoft ID: " + microsoftId);
                }

                log.Info("Task is executing further to create a VSTS item...");

                string comment = "{ \"body\": \"Failed to submit internal item.\" }";
                string label = "[ \"failed-logged-request\" ]";

                try
                {
                    var vstsItemUrl = await CreateVstsCustomerSuggestion(payload.issue.title.ToString(), payload.issue.body.ToString(), payload.issue.html_url.ToString(), microsoftId, log);
                    comment = "{ \"body\": \"🚀 **ATTENTION**: [Internal request](" + vstsItemUrl + ") logged.\" }";
                    label = "[ \"logged-request\" ]";
                }
                catch (Exception ex)
                {
                    log.Info("Failed to insert issue.");
                    log.Info(ex.Message);
                }

                if (payload.issue != null)
                {
                    log.Info($"{payload.issue.user.login} posted an issue #{payload.issue.number}:{payload.issue.title}");

                    //Post a comment 
                    await SendGitHubRequest(payload.issue.comments_url.ToString(), comment);

                    //Add a label
                    await SendGitHubRequest($"{payload.issue.url.ToString()}/labels", label);
                }
            }
        }
    }
}

public static async Task SendGitHubRequest(string url, string requestBody)
{
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("username", "version"));

        // Add the GITHUB_CREDENTIALS as an app setting, Value is the "PersonalAccessToken"
        // Please follow the link https://developer.github.com/v3/oauth/ to get more information on GitHub authentication 
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", Environment.GetEnvironmentVariable("GITHUB_CREDENTIALS"));
        var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
        await client.PostAsync(url, content);
    }
}

public static async Task<string> CreateVstsCustomerSuggestion(string title, string description, string linkToIssue, string microsoftId, TraceWriter log)
{
    string url = "";
    string complexDescription = $"{description}<br/><br/>Original GitHub Issue: <a href='{linkToIssue}'>{linkToIssue}</a>";

    var jsonizedTitle = JsonConvert.ToString(title);
    var jsonizedDescription = JsonConvert.ToString(complexDescription);
    var jsonizedId = JsonConvert.ToString(microsoftId);

    string baseString = $@"[
        {{
            ""op"": ""add"",
            ""value"": {jsonizedTitle},
            ""from"": null,
            ""path"":""/fields/System.Title""
        }},
        {{
            ""op"": ""add"",
            ""value"": {jsonizedDescription},
            ""from"": null,
            ""path"":""/fields/System.Description""
        }},
        {{
            ""op"": ""add"",
            ""value"": {jsonizedId},
            ""from"": null,
            ""path"":""/fields/System.AssignedTo""
        }}
    ]";

    log.Info("Creating item...");
    using (var client = new HttpClient())
    {
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(":" + Environment.GetEnvironmentVariable("VSTS_CREDENTIALS"))));
        var content = new StringContent(baseString, Encoding.UTF8, "application/json-patch+json");
        var response = await client.PostAsync(url, content);
        string contents = await response.Content.ReadAsStringAsync();

        //log.Info(contents);

        var json = JsonConvert.DeserializeObject<dynamic>(contents);
        return json._links.html.href.ToString();
    }
}

public static async Task<string> ResolveGitHubAliasToIdentity(string gitHubId, TraceWriter log)
{
    string url = $"";

    using (var client = new HttpClient())
    {
        try
        {
            //log.Info("Using API version: " + Environment.GetEnvironmentVariable("OSPO_API_VERSION"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(":" + Environment.GetEnvironmentVariable("VSTS_OSPO_CREDENTIALS"))));
            client.DefaultRequestHeaders.TryAddWithoutValidation("api-version", Environment.GetEnvironmentVariable("OSPO_API_VERSION"));

            var response = await client.GetAsync(url);

            string contents = await response.Content.ReadAsStringAsync();

            //log.Info(contents);

            var json = JsonConvert.DeserializeObject<dynamic>(contents);
            //log.Info("Used email: " + json.aad.emailAddress.ToString());

            return json.aad.userPrincipalName.ToString();
        }
        catch(Exception ex)
        {
            //log.Info(ex.Message);
            return string.Empty;
        }
    }
}