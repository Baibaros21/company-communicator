// <copyright file="NotificationRepositoryExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Extensions for the repository of the notification data.
    /// </summary>
    public static class NotificationRepositoryExtensions
    {
        /// <summary>
        /// Create a new draft notification.
        /// </summary>
        /// <param name="notificationRepository">The notification repository.</param>
        /// <param name="notification">Draft Notification model class instance passed in from Web API.</param>
        /// <param name="userName">Name of the user who is running the application.</param>
        /// <returns>The newly created notification's id.</returns>
        public static async Task<string> CreateDraftNotificationAsync(
            this INotificationDataRepository notificationRepository,
            DraftNotification notification,
            string userName)
        {
            var newId = notificationRepository.TableRowKeyGenerator.CreateNewKeyOrderingOldestToMostRecent();
            var imageId = newId + "-Image";
            var posterId = newId + "-poster";
            var notificationEntity = new NotificationDataEntity
            {
                PartitionKey = NotificationDataTableNames.DraftNotificationsPartition,
                RowKey = newId,
                Id = newId,
                Title = notification.Title,
                Department = notification.Department,
                ImageLink = notification.ImageLink,
                ImageEmbedLink = notification.ImageEmbedLink,
                PosterLink = notification.PosterLink,
                VideoLink = notification.VideoLink,
                Summary = notification.Summary,
                Author = notification.Author,
                ButtonTitle = notification.ButtonTitle,
                ButtonLink = notification.ButtonLink,
                CreatedBy = userName,
                CreatedDate = DateTime.UtcNow,
                IsDraft = true,
                Card = notification.Card,
                Teams = notification.Teams,
                Rosters = notification.Rosters,
                Groups = notification.Groups,
                AllUsers = notification.AllUsers,
                Template = notification.Template,
            };

            if (!string.IsNullOrEmpty(notification.ImageLink) && notification.ImageLink.StartsWith(Constants.ImageBase64Format))
            {
                notificationEntity.ImageLink = await notificationRepository.SaveImageAsync(imageId, notification.ImageLink);
                notificationEntity.ImageBase64BlobName = imageId;
            }
            else
            {
                notificationEntity.ImageLink = notification.ImageLink;
            }
            if (!string.IsNullOrEmpty(notification.PosterLink) && notification.PosterLink.StartsWith(Constants.ImageBase64Format))
            {
                notificationEntity.PosterLink = await notificationRepository.SaveImageAsync(posterId, notification.PosterLink);
                notificationEntity.PosterBase64BlobName = posterId;
            }
            else
            {
                notificationEntity.PosterLink = notification.PosterLink;
            }

            await notificationRepository.CreateOrUpdateAsync(notificationEntity);

            return newId;
        }
    }
}
