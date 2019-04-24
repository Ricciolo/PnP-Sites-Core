using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Diagnostics;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Linq;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model.Teams;
using OfficeDevPnP.Core.Utilities;
using System.Net;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Web;
using System.Net.Http;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Object Handler to manage Microsoft Teams stuff
    /// </summary>
    internal class ObjectTeams : ObjectHandlerBase
    {
        public override string Name => "Teams";
        public override string InternalName => "Teams";

        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#if !ONPREMISES
            using (var scope = new PnPMonitoredScope(Name))
            {
                var accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.ReadWrite.All User.ReadBasic.All");

                // - Teams Templates
                var teamTemplates = template.ParentHierarchy.Teams?.TeamTemplates;
                if (teamTemplates != null && teamTemplates.Any())
                {
                    foreach (var teamTemplate in teamTemplates)
                    {
                        var team = CreateByTeamTemplate(scope, parser, teamTemplate, accessToken);
                        // possible further processing...
                    }
                }

                // - Teams
                var teams = template.ParentHierarchy.Teams?.Teams;
                if (teams != null && teams.Any())
                {
                    foreach (var team in teams)
                    {
                        CreateByTeam(scope, parser, team, accessToken);
                        // possible further processing...
                    }
                }

                // - Apps
            }
#endif

            return parser;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            // So far, no extraction
            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template, ProvisioningTemplateApplyingInformation applyingInformation)
        {
#if !ONPREMISES
            if (!_willProvision.HasValue)
            {
                _willProvision = template.ParentHierarchy.Teams?.TeamTemplates?.Any() |
                    template.ParentHierarchy.Teams?.Teams?.Any();
            }
#else
            if (!_willProvision.HasValue)
            {
                _willProvision = false;
            }
#endif
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = false;
            }
            return _willExtract.Value;
        }

        private static JToken CreateByTeam(PnPMonitoredScope scope, TokenParser parser, Team team, string accessToken)
        {
            if (!String.IsNullOrWhiteSpace(team.CloneFrom))
            {
                // TODO: handle cloning
                scope.LogError("Cloning not supported yet");
                return null;
            }

            string groupId = CreateGroup(scope, team, accessToken);
            if (groupId == null) return null;

            string teamId = CreateGroupTeam(scope, team, groupId, accessToken);

            if (!SetGroupSecurity(scope, team, teamId, groupId, accessToken)) return null;
            if (!SetTeamChannels(scope, parser, team, teamId, groupId, accessToken)) return null;
            if (!SetTeamApps(scope, team, teamId, groupId, accessToken)) return null;

            try
            {

                return JToken.Parse(HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/beta/teams/{teamId}", accessToken));
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingError, ex.Message);
            }

            return null;
        }

        private static string CreateGroup(PnPMonitoredScope scope, Team team, string accessToken)
        {
            // Group id is specified
            if (!String.IsNullOrWhiteSpace(team.GroupId)) return team.GroupId;

            var content = new
            {
                displayName = team.DisplayName,
                mailEnabled = true,
                groupTypes = new[] { "Unified" },
                mailNickname = team.MailNickname,
                securityEnabled = false
            };

            return new HttpRequestAddOrUpdateConfig(
                scope,
                "https://graph.microsoft.com/beta/groups",
                content,
                "ObjectConflict",
                "mailNickname",
                team.MailNickname,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_CreatingGroupError,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_GroupAlreadyExists,
                accessToken).Execute();
        }

        private static bool SetGroupSecurity(PnPMonitoredScope scope, Team team, string teamId, string groupId, string accessToken)
        {
            string[] desideredOwnerIds;
            string[] desideredMemberIds;
            try
            {
                var userIdsByUPN = team.Security.Owners
                    .Select(o => o.UserPrincipalName)
                    .Concat(team.Security.Members.Select(m => m.UserPrincipalName))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(k => k, k =>
                    {
                        var jsonUser = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/beta/users/{WebUtility.UrlEncode(k)}?$select=id", accessToken);
                        return JToken.Parse(jsonUser).Value<string>("id");
                    });

                desideredOwnerIds = team.Security.Owners.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
                desideredMemberIds = team.Security.Members.Select(o => userIdsByUPN[o.UserPrincipalName]).ToArray();
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_FetchingUserError, ex.Message);
                return false;
            }

            string[] ownerIdsToAdd;
            string[] ownerIdsToRemove;
            try
            {
                // Get current group owners
                var jsonOwners = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/beta/groups/{groupId}/owners?$select=id", accessToken);

                string[] currentOwnerIds = GetIdsFromList(jsonOwners);

                // Exclude owners already into the group
                ownerIdsToAdd = desideredOwnerIds.Except(currentOwnerIds).ToArray();

                if (team.Security.ClearExistingOwners)
                {
                    ownerIdsToRemove = currentOwnerIds.Except(desideredOwnerIds).ToArray();
                }
                else
                {
                    ownerIdsToRemove = new string[0];
                }
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_ListingOwnersError, ex.Message);
                return false;
            }

            // Add new owners
            foreach (string ownerId in ownerIdsToAdd)
            {
                try
                {
                    object content = new JObject
                    {
                        ["@odata.id"] = $"https://graph.microsoft.com/beta/users/{ownerId}"
                    };
                    HttpHelper.MakePostRequest($"https://graph.microsoft.com/beta/groups/{groupId}/owners/$ref", content, "application/json", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_AddingOwnerError, ex.Message);
                    return false;
                }
            }

            // Remove exceeding owners
            foreach (string ownerId in ownerIdsToRemove)
            {
                try
                {
                    HttpHelper.MakeDeleteRequest($"https://graph.microsoft.com/beta/groups/{groupId}/owners/{ownerId}/$ref", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingOwnerError, ex.Message);
                    return false;
                }
            }

            string[] memberIdsToAdd;
            string[] memberIdsToRemove;
            try
            {
                // Get current group members
                var jsonOwners = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/beta/groups/{groupId}/members?$select=id", accessToken);

                string[] currentMemberIds = GetIdsFromList(jsonOwners);

                // Exclude members already into the group
                memberIdsToAdd = desideredMemberIds.Except(currentMemberIds).ToArray();

                if (team.Security.ClearExistingMembers)
                {
                    memberIdsToRemove = currentMemberIds.Except(desideredMemberIds).ToArray();
                }
                else
                {
                    memberIdsToRemove = new string[0];
                }
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_ListingMembersError, ex.Message);
                return false;
            }

            // Add new members
            foreach (string ownerId in memberIdsToAdd)
            {
                try
                {
                    object content = new JObject
                    {
                        ["@odata.id"] = $"https://graph.microsoft.com/beta/users/{ownerId}"
                    };
                    HttpHelper.MakePostRequest($"https://graph.microsoft.com/beta/groups/{groupId}/members/$ref", content, "application/json", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_AddingMemberError, ex.Message);
                    return false;
                }
            }

            // Remove exceeding members
            foreach (string memberId in memberIdsToRemove)
            {
                try
                {
                    HttpHelper.MakeDeleteRequest($"https://graph.microsoft.com/beta/groups/{groupId}/members/{memberId}/$ref", accessToken);
                }
                catch (Exception ex)
                {
                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_RemovingMemberError, ex.Message);
                    return false;
                }
            }

            return true;
        }

        private static bool SetTeamChannels(PnPMonitoredScope scope, TokenParser parser, Team team, string teamId, string groupId, string accessToken)
        {
            // TODO: create resource strings for exceptions

            if (team.Channels != null && team.Channels.Any())
            {
                foreach (var channel in team.Channels)
                {
                    // Create the channel object for the API call
                    var channelToCreate = new
                    {
                        channel.Description,
                        channel.DisplayName,
                        channel.IsFavoriteByDefault
                    };

                    string channelId = new HttpRequestAddOrUpdateConfig(
                        scope,
                        $"https://graph.microsoft.com/beta/teams/{teamId}/channels",
                        channelToCreate,
                        "NameAlreadyExists",
                        "displayName",
                        channel.DisplayName,
                        "todo error",
                        "todo warning",
                        accessToken)
                    {
                        CanPatch = false
                    }.Execute();
                    if (channelId == null) return false;

                    // If there are any Tabs for the current channel
                    if (channel.Tabs == null || !channel.Tabs.Any()) continue;

                    foreach (var tab in channel.Tabs)
                    {
                        // Create the object for the API call
                        var tabToCreate = new
                        {
                            tab.DisplayName,
                            tab.TeamsAppId,
                            configuration = tab.Configuration != null ? new
                            {
                                tab.Configuration.EntityId,
                                tab.Configuration.ContentUrl,
                                tab.Configuration.RemoveUrl,
                                tab.Configuration.WebsiteUrl
                            } : null
                        };

                        string tabId = new HttpRequestAddOrUpdateConfig(
                            scope,
                            $"https://graph.microsoft.com/beta/teams/{teamId}/channels/{channelId}/tabs",
                            tabToCreate,
                            "NameAlreadyExists",
                            "displayName",
                            tab.DisplayName,
                            "TODO: error",
                            "TODO: warning",
                            accessToken).Execute();
                        if (tabId == null) return false;
                    }

                    // TODO: Handle TabResources

                    // If there are any messages for the current channel
                    if (channel.Messages == null || !channel.Messages.Any()) continue;

                    foreach (var message in channel.Messages)
                    {
                        // Get and parse the CData
                        var messageString = parser.ParseString(message.Message);
                        var messageJson = JToken.Parse(messageString);

                        new HttpRequestAddConfig(scope,
                            $"https://graph.microsoft.com/beta/teams/{teamId}/channels/{channelId}/messages",
                            messageJson,
                            "TODO: error",
                            accessToken).Execute();
                    }
                }
            }

            return true;
        }

        private static bool SetTeamApps(PnPMonitoredScope scope, Team team, string teamId, string groupId, string accessToken)
        {
            foreach (var app in team.Apps)
            {
                object content = new JObject
                {
                    ["teamsApp@odata.bind"] = app.AppId
                };

                string id = new HttpRequestAddConfig(scope,
                    $"https://graph.microsoft.com/beta/teams/{teamId}/installedApps",
                    content,
                    "TODO: error",
                    accessToken).Execute();
            }

            return true;
        }

        private static string CreateGroupTeam(PnPMonitoredScope scope, Team team, string groupId, string accessToken)
        {
            var content = new
            {
                team.DisplayName,
                team.Description,
                team.Classification,
                team.Specialization,
                team.Visibility,
                funSettings = new
                {
                    team.FunSettings.AllowGiphy,
                    team.FunSettings.GiphyContentRating,
                    team.FunSettings.AllowStickersAndMemes,
                    team.FunSettings.AllowCustomMemes,
                },
                guestSettings = new
                {
                    team.GuestSettings.AllowCreateUpdateChannels,
                    team.GuestSettings.AllowDeleteChannels,
                },
                memberSettings = new
                {
                    team.MemberSettings.AllowCreateUpdateChannels,
                    team.MemberSettings.AllowAddRemoveApps,
                    team.MemberSettings.AllowDeleteChannels,
                    team.MemberSettings.AllowCreateUpdateRemoveTabs,
                    team.MemberSettings.AllowCreateUpdateRemoveConnectors
                },
                messagingSettings = new
                {
                    team.MessagingSettings.AllowUserEditMessages,
                    team.MessagingSettings.AllowUserDeleteMessages,
                    team.MessagingSettings.AllowOwnerDeleteMessages,
                    team.MessagingSettings.AllowTeamMentions,
                    team.MessagingSettings.AllowChannelMentions
                }
            };

            new HttpRequestAddOrUpdateConfig(
                scope,
                $"https://graph.microsoft.com/beta/groups/{groupId}/team",
                content,
                "Conflict",
                "id",
                groupId,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError,
                CoreResources.Provisioning_ObjectHandlers_Teams_Team_AlreadyExists,
                accessToken)
            {
                PutInsteadOfPost = true
            }.Execute();

            return groupId;
        }

        private static JToken CreateByTeamTemplate(PnPMonitoredScope scope, TokenParser parser, TeamTemplate teamTemplate, string accessToken)
        {
            HttpResponseHeaders responseHeaders;
            try
            {
                var content = OverwriteJsonTemplateProperties(parser, teamTemplate);
                responseHeaders = HttpHelper.MakePostRequestForHeaders("https://graph.microsoft.com/beta/teams", content, "application/json", accessToken);
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_ProvisioningError, ex.Message);
                return null;
            }

            try
            {
                var teamId = responseHeaders.Location.ToString().Split('\'')[1];
                var team = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/v1.0/groups/{teamId}", accessToken);
                return JToken.Parse(team);
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_FetchingError, ex.Message);
            }

            return null;
        }

        private static string[] GetIdsFromList(string json)
        {
            return JsonConvert.DeserializeAnonymousType(json, new { value = new[] { new { id = "" } } }).value.Select(v => v.id).ToArray();
        }

        private static string OverwriteJsonTemplateProperties(TokenParser parser, TeamTemplate teamTemplate)
        {
            var jsonTemplate = parser.ParseString(teamTemplate.JsonTemplate);
            var team = JToken.Parse(jsonTemplate);

            if (teamTemplate.DisplayName != null) team["displayName"] = teamTemplate.DisplayName;
            if (teamTemplate.Description != null) team["description"] = teamTemplate.Description;
            if (teamTemplate.Classification != null) team["classification"] = teamTemplate.Classification;
            if (teamTemplate.Visibility != null) team["visibility"] = teamTemplate.Visibility.ToString();

            return team.ToString();
        }

        private class HttpRequestAddConfig
        {
            public HttpRequestAddConfig(PnPMonitoredScope scope,
                                     string uri,
                                     object content,
                                     string errorMessage,
                                     string accessToken)
            {
                Scope = scope;
                Uri = uri;
                Content = content;
                ErrorMessage = errorMessage;
                AccessToken = accessToken;
            }

            public PnPMonitoredScope Scope { get; }

            public string Uri { get; }

            public object Content { get; }

            public string ErrorMessage { get; }

            public string AccessToken { get; }

            public bool PutInsteadOfPost { get; set; }

            public virtual string Execute()
            {
                try
                {
                    // Create item
                    string json = PutInsteadOfPost ?
                        HttpHelper.MakePutRequestForString(Uri, Content, "application/json", AccessToken)
                        : HttpHelper.MakePostRequestForString(Uri, Content, "application/json", AccessToken);

                    return JToken.Parse(json).Value<string>("id");
                }
                catch (Exception ex)
                {
                    return HandleError(ex);
                }
            }

            protected virtual string HandleError(Exception ex)
            {
                Scope.LogError(ErrorMessage, ex.Message);
                return null;
            }
        }

        private class HttpRequestAddOrUpdateConfig : HttpRequestAddConfig
        {
            public HttpRequestAddOrUpdateConfig(PnPMonitoredScope scope,
                                     string uri,
                                     object content,
                                     string conflictMessage,
                                     string conflictFieldName,
                                     string conflictFieldValue,
                                     string errorMessage,
                                     string warningMessage,
                                     string accessToken) : base(scope, uri, content, errorMessage, accessToken)
            {
                ConflictMessage = conflictMessage;
                ConflictFieldName = conflictFieldName;
                ConflictFieldValue = conflictFieldValue;
                WarningMessage = warningMessage;
            }

            public string ConflictMessage { get; }

            public string ConflictFieldName { get; }

            public string ConflictFieldValue { get; }

            public string WarningMessage { get; }

            public bool CanPatch { get; set; }

            protected override string HandleError(Exception ex)
            {
                // Group already exists
                if (ex.InnerException.Message.Contains(ConflictMessage))
                {
                    try
                    {
                        Scope.LogWarning(WarningMessage);

                        // If it's a POST we need to look for any existing item
                        string id = null;
                        string uri = Uri;
                        // In case of PUT we already have the id
                        if (!PutInsteadOfPost)
                        {
                            // Filter by field and value specified
                            string json = HttpHelper.MakeGetRequestForString($"{uri}?$select=id&$filter={ConflictFieldName}%20eq%20'{WebUtility.UrlEncode(ConflictFieldValue)}'");
                            id = GetIdsFromList(json)[0];
                            uri = $"{Uri}/{id}";
                        }

                        // Path the item
                        if (CanPatch)
                        {
                            HttpHelper.MakePatchRequestForString(uri, Content, "application/json", AccessToken);
                        }

                        return id;
                    }
                    catch (Exception ex2)
                    {
                        Scope.LogError(ErrorMessage, ex2.Message);
                        return null;
                    }
                }

                return base.HandleError(ex);
            }
        }
    }
}
