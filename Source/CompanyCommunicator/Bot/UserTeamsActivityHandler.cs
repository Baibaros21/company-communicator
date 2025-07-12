// <copyright file="UserTeamsActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Bot
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using Azure;
    using Microsoft.Azure.Documents.SystemFunctions;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Configuration.UserSecrets;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using  Microsoft.Teams.Apps.CompanyCommunicator.Common;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Company Communicator User Bot.
    /// Captures user data, team data.
    /// </summary>
    public class UserTeamsActivityHandler : TeamsActivityHandler
    {

        private static readonly string TeamRenamedEventType = "teamRenamed";

        private readonly TeamsDataCapture teamsDataCapture;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly INotificationDataRepository notificationDataRepository;
        /// <summary>
        /// Initializes a new instance of the <see cref="UserTeamsActivityHandler"/> class.
        /// </summary>
        /// <param name="teamsDataCapture">Teams data capture service.</param>
        public UserTeamsActivityHandler(TeamsDataCapture teamsDataCapture, ISentNotificationDataRepository sentNotificationDataRepository, INotificationDataRepository notificationDataRepository)
        {
            this.teamsDataCapture = teamsDataCapture ?? throw new ArgumentNullException(nameof(teamsDataCapture));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));

        }

        /// <summary>
        /// Handles task module fetch requests.
        /// </summary>
        /// <param name="turnContext">The turn context.</param>
        /// <param name="taskModuleRequest">The task module request.</param>
        /// <param name="cancellationToken">The cancellation token.</param>
        /// <returns>A task module response.</returns>
        protected override Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                var activityId = turnContext.Activity.ReplyToId;
                var data = taskModuleRequest.Data as JObject;
                // Access the "type" field inside the "msteams" object
                var msteamsType = data["type"]?.ToString();
                Console.WriteLine($"msteams type: {msteamsType}");

                if (data != null &&
                    data["type"] != null &&
                    data["type"].ToString() == "task/fetch")
                {
                    // Ensure URL uses HTTPS
                    var secureUrl = Common.Constants.BaseUrl + "/videoplayer?activityId=" + activityId;

                    var taskInfo = new TaskModuleTaskInfo
                    {
                        Url = secureUrl,
                        FallbackUrl = secureUrl,
                        Height = 700,
                        Width =800,
                        Title = "Videoplayer Task Module",
                    };

                    var response = new TaskModuleResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = taskInfo,
                        },
                    };
                    Console.WriteLine($"Task module response: {JsonConvert.SerializeObject(data)}");

                    return Task.FromResult(response);
                }
                else
                {
                    // Log when data is null
                    Console.WriteLine("Task module request data was null");
                    return Task.FromResult(new TaskModuleResponse
                    {
                        Task = new TaskModuleMessageResponse
                        {
                            Value = "There was an error loading the video player. Please try again.",
                            Type = "message",
                        },
                    });
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                Console.WriteLine($"Error in OnTeamsTaskModuleFetchAsync: {ex.Message}");

                // Return a friendly error message to the user
                return Task.FromResult(new TaskModuleResponse
                {
                    Task = new TaskModuleMessageResponse
                    {
                        Value = "An error occurred while loading the video player. Please try again later.",
                        Type = "message",
                    },
                });
            }
        }

        /// <summary>
        /// Invoked when a conversation update activity is received from the channel.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnConversationUpdateActivityAsync(
            ITurnContext<IConversationUpdateActivity> turnContext,
            CancellationToken cancellationToken)
        {
            // base.OnConversationUpdateActivityAsync is useful when it comes to responding to users being added to or removed from the conversation.
            // For example, a bot could respond to a user being added by greeting the user.
            // By default, base.OnConversationUpdateActivityAsync will call <see cref="OnMembersAddedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been added or <see cref="OnMembersRemovedAsync(IList{ChannelAccount}, ITurnContext{IConversationUpdateActivity}, CancellationToken)"/>
            // if any users have been removed. base.OnConversationUpdateActivityAsync checks the member ID so that it only responds to updates regarding members other than the bot itself.
            await base.OnConversationUpdateActivityAsync(turnContext, cancellationToken);

            var activity = turnContext.Activity;

            var isTeamRenamed = this.IsTeamInformationUpdated(activity);
            if (isTeamRenamed)
            {
                await this.teamsDataCapture.OnTeamInformationUpdatedAsync(activity);
            }

            if (activity.MembersAdded != null)
            {
                await this.teamsDataCapture.OnBotAddedAsync(turnContext, activity, cancellationToken);
            }

            if (activity.MembersRemoved != null)
            {
                await this.teamsDataCapture.OnBotRemovedAsync(activity);
            }
        }

        /// <inheritdoc/>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {

            await turnContext.SendActivityAsync(MessageFactory.Text($"Echo: {turnContext.Activity.Text}"), cancellationToken);
        }



        /// <inheritdoc/>
        protected override async Task OnReactionsAddedAsync(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {
            await this.OnReactionChanged(messageReactions, turnContext, true);
            await base.OnReactionsAddedAsync(messageReactions, turnContext, cancellationToken);
        }

        /// <inheritdoc/>
        protected override async Task OnTeamsReadReceiptAsync(ReadReceiptInfo readReceiptInfo, ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
        {
            // Check if the user has read the message
            if (!readReceiptInfo.IsNull())
            {
                //User has read the message
                //Add your logic here to handle this scenario
                //turnContext.Activity.from.Id will get the user who recived the messages
                //use msgraph.user to retrive user name 
                //store the name in SendNotifcationS
                string messageId = readReceiptInfo.LastReadMessageId;
                string activityId = turnContext.Activity.Id;
                var result = await this.sentNotificationDataRepository.GetNotificationByColumnFilter("ActivityId", activityId);
                var resultList = new List<SentNotificationDataEntity>(result);
                NotificationDataEntity notification = await this.notificationDataRepository.GetAsync("SentNotifications", resultList[0].PartitionKey);
                notification.Seen++;
                //await this.sentNotificationDataRepository.CreateOrUpdateAsync(resultList[0]);
                await this.notificationDataRepository.CreateOrUpdateAsync(notification);

            }
            else
            {
                // User has not read the message
                // Add your logic here to handle this scenario
            }

        }

        /// <inheritdoc/>
        protected override async Task OnReactionsRemovedAsync(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, CancellationToken cancellationToken)
        {

            await this.OnReactionChanged(messageReactions, turnContext, false);
            await base.OnReactionsRemovedAsync(messageReactions, turnContext, cancellationToken);   
        }

        private async Task OnReactionChanged(IList<MessageReaction> messageReactions, ITurnContext<IMessageReactionActivity> turnContext, bool add)
        {
            string activityId = turnContext.Activity.Id;

            var result = await this.sentNotificationDataRepository.GetNotificationByColumnFilter("ActivityId", activityId);
            var resultList = new List<SentNotificationDataEntity>(result);
            NotificationDataEntity notification = await this.notificationDataRepository.GetAsync("SentNotifications", resultList[0].PartitionKey);
            if (add)
            {
                foreach (var reaction in messageReactions)
                {

                    var reactionType = reaction.Type.ToLower();
                    switch (reactionType)
                    {
                        case "like":
                            notification.Like++;
                            break;
                        case "heart":
                            notification.Heart++;
                            break;
                        case "laugh":
                            notification.Laugh++;
                            break;
                        case "surprised":
                            notification.Surprise++;
                            break;
                    }

                }
            }
            else
            {
                foreach (var reaction in messageReactions)
                {

                    var reactionType = reaction.Type.ToLower();
                    switch (reactionType)
                    {
                        case "like":
                            notification.Like--;
                            break;
                        case "heart":
                            notification.Heart--;
                            break;
                        case "laugh":
                            notification.Laugh--;
                            break;
                        case "surprised":
                            notification.Surprise--;
                            break;
                    }
                }
            }
            await this.notificationDataRepository.CreateOrUpdateAsync(notification);
        }

        private bool IsTeamInformationUpdated(IConversationUpdateActivity activity)
        {
            if (activity == null)
            {
                return false;
            }

            var channelData = activity.GetChannelData<TeamsChannelData>();
            if (channelData == null)
            {
                return false;
            }

            return UserTeamsActivityHandler.TeamRenamedEventType.Equals(channelData.EventType, StringComparison.OrdinalIgnoreCase);
        }



    }
}