// <copyright file="MessageService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams
{
    using System;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Adapter;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;
    using Newtonsoft.Json;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Teams message service.
    /// </summary>
    public class MessageService : IMessageService
    {
        private readonly string microsoftAppId;
        private readonly CCBotAdapterBase botAdapter;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageService"/> class.
        /// </summary>
        /// <param name="botOptions">The bot options.</param>
        /// <param name="botAdapter">The bot adapter.</param>
        public MessageService(
            IOptions<BotOptions> botOptions,
            CCBotAdapterBase botAdapter)
        {
            this.microsoftAppId = botOptions?.Value?.UserAppId ?? throw new ArgumentNullException(nameof(botOptions));
            this.botAdapter = botAdapter ?? throw new ArgumentNullException(nameof(botAdapter));
        }

        /// <inheritdoc/>
        //public async Task<SendMessageResponse> SendMessageAsync(
        //    IMessageActivity message,
        //    string conversationId,
        //    string serviceUrl,
        //    int maxAttempts,
        //    ILogger log)
        //{
        //    if (message is null)
        //    {
        //        throw new ArgumentNullException(nameof(message));
        //    }

        //    if (string.IsNullOrEmpty(conversationId))
        //    {
        //        throw new ArgumentException($"'{nameof(conversationId)}' cannot be null or empty", nameof(conversationId));
        //    }

        //    if (string.IsNullOrEmpty(serviceUrl))
        //    {
        //        throw new ArgumentException($"'{nameof(serviceUrl)}' cannot be null or empty", nameof(serviceUrl));
        //    }

        //    if (log is null)
        //    {
        //        throw new ArgumentNullException(nameof(log));
        //    }

        //    var conversationReference = new ConversationReference
        //    {
        //        ServiceUrl = serviceUrl,
        //        Conversation = new ConversationAccount
        //        {
        //            Id = conversationId,
        //        },
        //    };

        //    var response = new SendMessageResponse
        //    {
        //        TotalNumberOfSendThrottles = 0,
        //        AllSendStatusCodes = string.Empty,
        //    };

        //    await this.botAdapter.ContinueConversationAsync(
        //        botAppId: this.microsoftAppId,
        //        reference: conversationReference,
        //        callback: async (turnContext, cancellationToken) =>
        //        {
        //            var policy = this.GetRetryPolicy(maxAttempts, log);
        //            try
        //            {
        //                /*Start RND Code*/

        //                //message.Text = "**📢 Important Announcement**\nYour actual message content here.";
        //                log.LogInformation($"Attachments count: {message.Attachments?.Count ?? 0}");



        //                /*End the RND Code*/


        //                // Send message.
        //                await policy.ExecuteAsync(async () => await turnContext.SendActivityAsync(message));

        //                // Success.
        //                response.ResultType = SendMessageResult.Succeeded;
        //                response.StatusCode = (int)HttpStatusCode.Created;
        //                response.AllSendStatusCodes += $"{(int)HttpStatusCode.Created},";
        //            }
        //            catch (ErrorResponseException exception)
        //            {
        //                var errorMessage = $"{exception.GetType()}: {exception.Message}";
        //                log.LogError(exception, $"Failed to send message. Exception message: {errorMessage}");

        //                response.StatusCode = (int)exception.Response.StatusCode;
        //                response.AllSendStatusCodes += $"{(int)exception.Response.StatusCode},";
        //                response.ErrorMessage = exception.ToString();
        //                switch (exception.Response.StatusCode)
        //                {
        //                    case HttpStatusCode.TooManyRequests:
        //                        response.ResultType = SendMessageResult.Throttled;
        //                        response.TotalNumberOfSendThrottles = maxAttempts;
        //                        break;

        //                    case HttpStatusCode.NotFound:
        //                        response.ResultType = SendMessageResult.RecipientNotFound;
        //                        break;

        //                    default:
        //                        response.ResultType = SendMessageResult.Failed;
        //                        break;
        //                }
        //            }
        //        },
        //        cancellationToken: CancellationToken.None);

        //    return response;
        //}

        public async Task<SendMessageResponse> SendMessageAsync(
          IMessageActivity message,
          string conversationId,
          string serviceUrl,
          int maxAttempts,
          ILogger log)
        {
            if (message is null)
            {
                throw new ArgumentNullException(nameof(message));
            }

            if (string.IsNullOrEmpty(conversationId))
            {
                throw new ArgumentException($"'{nameof(conversationId)}' cannot be null or empty", nameof(conversationId));
            }

            if (string.IsNullOrEmpty(serviceUrl))
            {
                throw new ArgumentException($"'{nameof(serviceUrl)}' cannot be null or empty", nameof(serviceUrl));
            }

            if (log is null)
            {
                throw new ArgumentNullException(nameof(log));
            }

            var conversationReference = new ConversationReference
            {
                ServiceUrl = serviceUrl,
                Conversation = new ConversationAccount
                {
                    Id = conversationId,
                },
            };

            var response = new SendMessageResponse
            {
                TotalNumberOfSendThrottles = 0,
                AllSendStatusCodes = string.Empty,
            };

            await this.botAdapter.ContinueConversationAsync(
                botAppId: this.microsoftAppId,
                reference: conversationReference,
                callback: async (turnContext, cancellationToken) =>
                {
                    var policy = this.GetRetryPolicy(maxAttempts, log);
                    try
                    {
                        /*Start RND Code*/
                        ////message.Text = "**📢 Important Announcement**\nYour actual message content here.";
                        //log.LogInformation($"Attachments count: {message.Attachments?.Count ?? 0}");
                        ///*End the RND Code*/

                        //// Send message.
                        //await policy.ExecuteAsync(async () => await turnContext.SendActivityAsync(message));

                        //// Success.
                        //response.ResultType = SendMessageResult.Succeeded;
                        //response.StatusCode = (int)HttpStatusCode.Created;
                        //response.AllSendStatusCodes += $"{(int)HttpStatusCode.Created},";

                        log.LogInformation($"Attachments count: {message.Attachments?.Count ?? 0}");

                        string notificationTitle = "📢 New Notification"; // Default title

                        if (message.Attachments != null && message.Attachments.Count > 0)
                        {
                            // Extract Adaptive Card JSON
                            var adaptiveCardJson = JsonConvert.SerializeObject(message.Attachments[0].Content);
                            log.LogInformation($"Adaptive Card JSON: {adaptiveCardJson}");

                            dynamic adaptiveCard = JsonConvert.DeserializeObject(adaptiveCardJson);

                            // Extract Notification Title
                            if (adaptiveCard?.title != null)
                            {
                                notificationTitle = $"📢 {adaptiveCard.title}";
                            }
                            else if (adaptiveCard?.body != null && adaptiveCard.body.Count > 0)
                            {
                                notificationTitle = $"📢 {adaptiveCard.body[0].text}";
                            }
                        }

                        // ✅ 1. Send only the title first (to appear in the right-side popup)
                        var notificationMessage = new Activity
                        {
                            Type = ActivityTypes.Message,
                            Text = notificationTitle,
                            Summary = notificationTitle, // Ensures the title appears in the notification
                            ChannelData = new
                            {
                                Notification = new
                                {
                                    Alert = true, // Forces the popup to appear
                                    Text = notificationTitle // This should be shown in the popup
                                }
                            }
                        };

                        await policy.ExecuteAsync(async () => await turnContext.SendActivityAsync(notificationMessage));

                        // ✅ 2. Send the Adaptive Card separately (without overriding the pop-up)
                        if (message.Attachments != null && message.Attachments.Count > 0)
                        {
                            var adaptiveCardMessage = MessageFactory.Attachment(message.Attachments[0]);
                            adaptiveCardMessage.Summary = null; // Prevents "Sent a card" from appearing
                            adaptiveCardMessage.ChannelData = new
                            {
                                Notification = new
                                {
                                    Alert = false // Prevents an additional pop-up
                                }
                            };

                            await policy.ExecuteAsync(async () => await turnContext.SendActivityAsync(adaptiveCardMessage));
                        }

                        // ✅ Success response
                        response.ResultType = SendMessageResult.Succeeded;
                        response.StatusCode = (int)HttpStatusCode.Created;
                        response.AllSendStatusCodes += $"{(int)HttpStatusCode.Created},";
                    }
                    catch (ErrorResponseException exception)
                    {
                        var errorMessage = $"{exception.GetType()}: {exception.Message}";
                        log.LogError(exception, $"Failed to send message. Exception message: {errorMessage}");

                        response.StatusCode = (int)exception.Response.StatusCode;
                        response.AllSendStatusCodes += $"{(int)exception.Response.StatusCode},";
                        response.ErrorMessage = exception.ToString();
                        switch (exception.Response.StatusCode)
                        {
                            case HttpStatusCode.TooManyRequests:
                                response.ResultType = SendMessageResult.Throttled;
                                response.TotalNumberOfSendThrottles = maxAttempts;
                                break;

                            case HttpStatusCode.NotFound:
                                response.ResultType = SendMessageResult.RecipientNotFound;
                                break;

                            default:
                                response.ResultType = SendMessageResult.Failed;
                                break;
                        }
                    }
                },
                cancellationToken: CancellationToken.None);

            return response;
        }

        private AsyncRetryPolicy GetRetryPolicy(int maxAttempts, ILogger log)
        {
            var delay = Backoff.DecorrelatedJitterBackoffV2(medianFirstRetryDelay: TimeSpan.FromSeconds(1), retryCount: maxAttempts);
            return Policy
                .Handle<ErrorResponseException>(e =>
                {
                    var errorMessage = $"{e.GetType()}: {e.Message}";
                    log.LogError(e, $"Exception thrown: {errorMessage}");

                    // Handle throttling and internal server errors.
                    var statusCode = e.Response.StatusCode;
                    return statusCode == HttpStatusCode.TooManyRequests || ((int)statusCode >= 500 && (int)statusCode < 600);
                })
                .WaitAndRetryAsync(delay);
        }
    }
}
