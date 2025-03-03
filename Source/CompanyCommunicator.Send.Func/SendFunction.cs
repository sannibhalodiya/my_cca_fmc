// <copyright file="SendFunction.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services;
    using Newtonsoft.Json;

    /// <summary>
    /// Azure Function App triggered by messages from a Service Bus queue
    /// Used for sending messages from the bot.
    /// </summary>
    public class SendFunction
    {

        /// <summary>
        /// This is set to 10 because the default maximum delivery count from the service bus
        /// message queue before the service bus will automatically put the message in the Dead Letter
        /// Queue is 10.
        /// </summary>
        private static readonly int MaxDeliveryCountForDeadLetter = 10;
        private static readonly string AdaptiveCardContentType = "application/vnd.microsoft.card.adaptive";
        private static readonly string CachePrefixSentCards = "sentcard_";

        private readonly int maxNumberOfAttempts;
        private readonly double sendRetryDelayNumberOfSeconds;
        private readonly INotificationService notificationService;
        private readonly ISendingNotificationDataRepository notificationRepo;
        private readonly IMessageService messageService;
        private readonly ISendQueue sendQueue;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly IMemoryCache memoryCache;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendFunction"/> class.
        /// </summary>
        /// <param name="options">Send function options.</param>
        /// <param name="notificationService">The service to precheck and determine if the queue message should be processed.</param>
        /// <param name="messageService">Message service.</param>
        /// <param name="notificationRepo">Notification repository.</param>
        /// <param name="sendQueue">The send queue.</param>
        /// <param name="localizer">Localization service.</param>
        public SendFunction(
            IOptions<SendFunctionOptions> options,
            INotificationService notificationService,
            IMessageService messageService,
            ISendingNotificationDataRepository notificationRepo,
            ISendQueue sendQueue,
            IStringLocalizer<Strings> localizer,
            IMemoryCache memoryCache)
        {
            if (options is null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            this.maxNumberOfAttempts = options.Value.MaxNumberOfAttempts;
            this.sendRetryDelayNumberOfSeconds = options.Value.SendRetryDelayNumberOfSeconds;

            this.notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            this.messageService = messageService ?? throw new ArgumentNullException(nameof(messageService));
            this.notificationRepo = notificationRepo ?? throw new ArgumentNullException(nameof(notificationRepo));
            this.sendQueue = sendQueue ?? throw new ArgumentNullException(nameof(sendQueue));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.memoryCache = memoryCache ?? throw new ArgumentNullException(nameof(memoryCache));
        }

        /// <summary>
        /// Azure Function App triggered by messages from a Service Bus queue
        /// Used for sending messages from the bot.
        /// </summary>
        /// <param name="myQueueItem">The Service Bus queue item.</param>
        /// <param name="deliveryCount">The deliver count.</param>
        /// <param name="enqueuedTimeUtc">The enqueued time.</param>
        /// <param name="messageId">The message ID.</param>
        /// <param name="log">The logger.</param>
        /// <param name="context">The execution context.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName("SendMessageFunction")]
        public async Task Run(
            [ServiceBusTrigger(
                SendQueue.QueueName,
                Connection = SendQueue.ServiceBusConnectionConfigurationKey)]
            string myQueueItem,
            int deliveryCount,
            DateTime enqueuedTimeUtc,
            string messageId,
            ILogger log,
            ExecutionContext context)
        {
            log.LogInformation($"C# ServiceBus queue trigger function processed message: {myQueueItem}");

            var messageContent = JsonConvert.DeserializeObject<SendQueueMessageContent>(myQueueItem);

            try
            {
                // Check if notification is canceled.
                var isCanceled = await this.notificationService.IsNotificationCanceled(messageContent);
                if (isCanceled)
                {
                    // No-op in case notification is canceled.
                    return;
                }

                // Check if recipient is a guest user.
                if (messageContent.IsRecipientGuestUser())
                {
                    await this.notificationService.UpdateSentNotification(
                        notificationId: messageContent.NotificationId,
                        recipientId: messageContent.RecipientData.RecipientId,
                        totalNumberOfSendThrottles: 0,
                        statusCode: SentNotificationDataEntity.NotSupportedStatusCode,
                        allSendStatusCodes: $"{SentNotificationDataEntity.NotSupportedStatusCode},",
                        errorMessage: this.localizer.GetString("GuestUserNotSupported"));
                    return;
                }

                // Check if notification is pending.
                var isPending = await this.notificationService.IsPendingNotification(messageContent);
                if (!isPending)
                {
                    // Notification is either already sent or failed and shouldn't be retried.
                    return;
                }

                // Check if conversationId is set to send message.
                if (string.IsNullOrWhiteSpace(messageContent.GetConversationId()))
                {
                    await this.notificationService.UpdateSentNotification(
                        notificationId: messageContent.NotificationId,
                        recipientId: messageContent.RecipientData.RecipientId,
                        totalNumberOfSendThrottles: 0,
                        statusCode: SentNotificationDataEntity.FinalFaultedStatusCode,
                        allSendStatusCodes: $"{SentNotificationDataEntity.FinalFaultedStatusCode},",
                        errorMessage: this.localizer.GetString("AppNotInstalled"));
                    return;
                }

                // Check if the system is throttled.
                var isThrottled = await this.notificationService.IsSendNotificationThrottled();
                if (isThrottled)
                {
                    // Re-Queue with delay.
                    await this.sendQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                    return;
                }

                // Send message.

                //var custome = await this.GetMessageActivity_Custom(messageContent, log);
                //var response1 = await this.messageService.SendMessageAsync(
                //    message: custome,
                //    serviceUrl: messageContent.GetServiceUrl(),
                //    conversationId: messageContent.GetConversationId(),
                //    maxAttempts: this.maxNumberOfAttempts,
                //    logger: log);


                var messageActivity = await this.GetMessageActivity(messageContent, log);
                var response = await this.messageService.SendMessageAsync(
                    message: messageActivity,
                    serviceUrl: messageContent.GetServiceUrl(),
                    conversationId: messageContent.GetConversationId(),
                    maxAttempts: this.maxNumberOfAttempts,
                    logger: log);

                // Process response.
                await this.ProcessResponseAsync(messageContent, response, log);
            }
            catch (InvalidOperationException exception)
            {
                // Bad message shouldn't be requeued.
                log.LogError(exception, $"InvalidOperationException thrown. Error message: {exception.Message}");
            }
            catch (Exception exception)
            {
                var exceptionMessage = $"{exception.GetType()}: {exception.Message}";
                log.LogError(exception, $"Failed to send message. ErrorMessage: {exceptionMessage}");

                // Update status code depending on delivery count.
                var statusCode = SentNotificationDataEntity.FaultedAndRetryingStatusCode;
                if (deliveryCount >= SendFunction.MaxDeliveryCountForDeadLetter)
                {
                    // Max deliveries attempted. No further retries.
                    statusCode = SentNotificationDataEntity.FinalFaultedStatusCode;
                }

                // Update sent notification table.
                await this.notificationService.UpdateSentNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: 0,
                    statusCode: statusCode,
                    allSendStatusCodes: $"{statusCode},",
                    errorMessage: this.localizer.GetString("Failed"),
                    exception: exception.ToString());

                throw;
            }
        }

        /// <summary>
        /// Process send notification response.
        /// </summary>
        /// <param name="messageContent">Message content.</param>
        /// <param name="sendMessageResponse">Send notification response.</param>
        /// <param name="log">Logger.</param>
        private async Task ProcessResponseAsync(
            SendQueueMessageContent messageContent,
            SendMessageResponse sendMessageResponse,
            ILogger log)
        {
            var statusReason = string.Empty;
            if (sendMessageResponse.ResultType == SendMessageResult.Succeeded)
            {
                log.LogInformation($"Successfully sent the message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}");
            }
            else
            {
                log.LogError($"Failed to send message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}" +
                    $"\nResult: {sendMessageResponse.ResultType}." +
                    $"\nErrorMessage: {sendMessageResponse.ErrorMessage}.");

                statusReason = this.localizer.GetString("Failed");
            }

            await this.notificationService.UpdateSentNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: sendMessageResponse.TotalNumberOfSendThrottles,
                    statusCode: sendMessageResponse.StatusCode,
                    allSendStatusCodes: sendMessageResponse.AllSendStatusCodes,
                    errorMessage: statusReason,
                    exception: sendMessageResponse.ErrorMessage);

            // Throttled
            if (sendMessageResponse.ResultType == SendMessageResult.Throttled)
            {
                // Set send function throttled.
                await this.notificationService.SetSendNotificationThrottled(this.sendRetryDelayNumberOfSeconds);

                // Requeue.
                await this.sendQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                return;
            }
        }

        private async Task<IMessageActivity> GetMessageActivity(SendQueueMessageContent message, ILogger log)
        {
            var cacheKeySentCard = CachePrefixSentCards + message.NotificationId;
            bool isCacheEntryExists = this.memoryCache.TryGetValue(cacheKeySentCard, out string jsonAC);

            if (!isCacheEntryExists)
            {
                // Download serialized AC from blob storage.
                jsonAC = await this.notificationRepo.GetAdaptiveCardAsync(message.NotificationId);
                this.memoryCache.Set(cacheKeySentCard, jsonAC, TimeSpan.FromHours(Constants.CacheDurationInHours));

                log.LogInformation($"Successfully cached the sent card data." +
                                $"\nNotificationId Id: {message.NotificationId}");
            }

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCardContentType,
                Content = JsonConvert.DeserializeObject(jsonAC),
            };

            return MessageFactory.Attachment(adaptiveCardAttachment);

            //var messageActivity = MessageFactory.Attachment(adaptiveCardAttachment);
            //messageActivity.Text = "Notification Title....";

            //return messageActivity;
        }


        private async Task<IMessageActivity> GetMessageActivity_Custom(SendQueueMessageContent message, ILogger log)
        {
            var cacheKeySentCard = CachePrefixSentCards + message.NotificationId;
            bool isCacheEntryExists = this.memoryCache.TryGetValue(cacheKeySentCard, out string jsonAC);

            if (!isCacheEntryExists)
            {
                // Download serialized AC from blob storage.
                jsonAC = await this.notificationRepo.GetAdaptiveCardAsync(message.NotificationId);
                this.memoryCache.Set(cacheKeySentCard, jsonAC, TimeSpan.FromHours(Constants.CacheDurationInHours));

                log.LogInformation($"Successfully cached the sent card data." +
                                $"\nNotificationId Id: {message.NotificationId}");
            }

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCardContentType,
                Content = JsonConvert.DeserializeObject(jsonAC),
            };

            return MessageFactory.Text("Custome Notification Title.....");

            //var messageActivity = MessageFactory.Attachment(adaptiveCardAttachment);
            //messageActivity.Text = "Notification Title....";

            //return messageActivity;
        }
    }
}