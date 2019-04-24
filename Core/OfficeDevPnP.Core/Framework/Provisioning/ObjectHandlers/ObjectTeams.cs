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
                var accessToken = PnPProvisioningContext.Current.AcquireToken("https://graph.microsoft.com/", "Group.ReadWrite.All");

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

            string groupId = WebUtility.UrlEncode(team.GroupId);
            string teamId = CreateGroupTeam(scope, team, groupId, accessToken);

            SetGroupSecurity(scope, team, teamId, groupId, accessToken);
            SetTeamChannels(scope, team, teamId, groupId, accessToken);
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

        private static void SetGroupSecurity(PnPMonitoredScope scope, Team team, string teamId, string groupId, string accessToken)
        {
            throw new NotImplementedException();
        }

        private static void SetTeamChannels(PnPMonitoredScope scope, Team team, string teamId, string groupId, string accessToken)
        {
            throw new NotImplementedException();
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
                    isArchived = team.Archived,
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
                
                responseHeaders = HttpHelper.MakePostRequestForHeaders($"https://graph.microsoft.com/beta/groups/{groupId}/team", content, "application/json", accessToken);

                return responseHeaders.Location.ToString().Split('\'')[1];
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
