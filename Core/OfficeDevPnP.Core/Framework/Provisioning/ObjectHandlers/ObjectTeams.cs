﻿using Microsoft.SharePoint.Client;
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
            string teamId = CreateGroupTeam(scope, team, groupId, accessToken);

            SetGroupSecurity(scope, team, teamId, groupId, accessToken);
            SetTeamChannels(scope, parser, team, teamId, groupId, accessToken);
            SetTeamApps(scope, team, teamId, groupId, accessToken);

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

            try
            {
                var content = new
                {
                    displayName = "pnp test",
                    mailEnabled = true,
                    groupTypes = new[] { "Unified" },
                    mailNickname = "pnptest",
                    securityEnabled = false
                };

                // Create group
                var json = HttpHelper.MakePostRequestForString($"https://graph.microsoft.com/beta/groups", content, "application/json", accessToken);

                return JToken.Parse(json).Value<string>("id");
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_CreatingGroupError, ex.Message);
                return null;
            }
        }

        private static void SetGroupSecurity(PnPMonitoredScope scope, Team team, string teamId, string groupId, string accessToken)
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
                return;
            }

            string[] ownerIdsToAdd;
            string[] ownerIdsToRemove;
            try
            {
                // Get current group owners
                var jsonOwners = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/beta/groups/{groupId}/owners?$select=id", accessToken);

                string[] currentOwnerIds = JsonConvert.DeserializeAnonymousType(jsonOwners, new { value = new[] { new { id = "" } } })
                    .value.Select(i => i.id)
                    .ToArray();

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
                return;
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
                    return;
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
                    return;
                }
            }

            string[] memberIdsToAdd;
            string[] memberIdsToRemove;
            try
            {
                // Get current group members
                var jsonOwners = HttpHelper.MakeGetRequestForString($"https://graph.microsoft.com/beta/groups/{groupId}/members?$select=id", accessToken);

                string[] currentMemberIds = JsonConvert.DeserializeAnonymousType(jsonOwners, new { value = new[] { new { id = "" } } })
                    .value.Select(i => i.id)
                    .ToArray();

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
                return;
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
                    return;
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
                    return;
                }
            }
        }

        private static void SetTeamChannels(PnPMonitoredScope scope, TokenParser parser, Team team, string teamId, string groupId, string accessToken)
        {
            // TODO: create resource strings for exceptions

            if (team.Channels != null && team.Channels.Any())
            {
                foreach (var channel in team.Channels)
                {
                    string channelId;

                    // Create the channel object for the API call
                    var channelToCreate = new
                    {
                        channel.Description,
                        channel.DisplayName,
                        channel.IsFavoriteByDefault
                    };

                    try
                    {
                        // POST an API request to create the channel
                        var responseHeaders = HttpHelper.MakePostRequestForHeaders($"https://graph.microsoft.com/beta/teams/{teamId}/channels", channelToCreate, "application/json", accessToken);
                        channelId = responseHeaders.Location.ToString().Split('\'')[1];
                    }
                    catch (Exception ex)
                    {
                        scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError, ex.Message);
                        return;
                    }

                    // If there are any Tabs for the current channel
                    if (channel.Tabs != null && channel.Tabs.Any())
                    {
                        try
                        {
                            foreach (var tab in channel.Tabs)
                            {
                                // Create the object for the API call
                                var tabToCreate = new
                                {
                                    tab.DisplayName,
                                    tab.TeamsAppId,
                                    configuration = new
                                    {
                                        tab.Configuration.EntityId,
                                        tab.Configuration.ContentUrl,
                                        tab.Configuration.RemoveUrl,
                                        tab.Configuration.WebsiteUrl
                                    }
                                };

                                try
                                {
                                    // POST an API request to create the tab
                                    HttpHelper.MakePostRequestForHeaders($"https://graph.microsoft.com/beta/teams/{teamId}/channels/{channelId}/tabs", tabToCreate, "application/json", accessToken);
                                }
                                catch (Exception ex)
                                {
                                    scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_FetchingError, ex.Message);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_FetchingError, ex.Message);
                        }
                    }

                    // TODO: Handle TabResources

                    // If there are any messages for the current channel
                    if (channel.Messages != null && channel.Messages.Any())
                    {
                        foreach (var message in channel.Messages)
                        {
                            try
                            {
                                // Get and parse the CData
                                var messageString = parser.ParseString(message.Message);
                                var messageJson = JToken.Parse(messageString);

                                // POST the message to the API
                                HttpHelper.MakePostRequest($"https://graph.microsoft.com/beta/teams/{teamId}/channels/{channelId}/messages", messageJson, "application/json", accessToken);
                            }
                            catch (Exception ex)
                            {
                                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_TeamTemplate_FetchingError, ex.Message);
                            }
                        }
                    }
                }
            }
        }

        private static void SetTeamApps(PnPMonitoredScope scope, Team team, string teamId, string groupId, string accessToken)
        {
            throw new NotImplementedException();
        }

        private static string CreateGroupTeam(PnPMonitoredScope scope, Team team, string groupId, string accessToken)
        {
            HttpResponseHeaders responseHeaders;
            try
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

                string json = HttpHelper.MakePutRequestForString($"https://graph.microsoft.com/beta/groups/{groupId}/team", content, "application/json", accessToken);

                return JToken.Parse(json).Value<string>("id");
                //return responseHeaders.Location.ToString().Split('\'')[1];
            }
            catch (Exception ex)
            {
                scope.LogError(CoreResources.Provisioning_ObjectHandlers_Teams_Team_ProvisioningError, ex.Message);
                return null;
            }
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
    }
}
