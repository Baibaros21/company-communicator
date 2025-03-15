using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Configuration;
using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
using Microsoft.Teams.Apps.CompanyCommunicator.Common;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
using Microsoft.Teams.Apps.CompanyCommunicator.DraftNotificationPreview;
using Microsoft.Teams.Apps.CompanyCommunicator.Models;
using Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob;


namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    /// <summary>
    /// Controller for the draft notification data.
    /// </summary>
    [Route("api/defaultdata")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class DefaultsDataController : ControllerBase
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ITeamDataRepository teamDataRepository;
        private readonly IDraftNotificationPreviewService draftNotificationPreviewService;
        private readonly IGroupsService groupsService;
        private readonly IAppSettingsService appSettingsService;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly IBlobStorageProvider blobStorageProvider;
        private readonly string DEFAULT_LOGO_BLOB_NAME = "DEFAULT_LOGO";
        private readonly string DEFAULT_BANNER_BLOB_NAME = "DEFAULT_BANNER";
        private readonly string DEFAULT_HEADER_LOGO = "DEFAULT_HEADER_LOGO";

        /// <summary>
        /// Gets the IConfiguration instance.
        /// </summary>
        public IConfiguration Configuration { get; }

        public DefaultsDataController(
            INotificationDataRepository notificationDataRepository,
            ITeamDataRepository teamDataRepository,
            IDraftNotificationPreviewService draftNotificationPreviewService,
            IAppSettingsService appSettingsService,
            IStringLocalizer<Strings> localizer,
            IGroupsService groupsService,
            IBlobStorageProvider blobStorageProvider,
            IConfiguration configuration)
        {
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
            this.draftNotificationPreviewService = draftNotificationPreviewService ?? throw new ArgumentNullException(nameof(draftNotificationPreviewService));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.groupsService = groupsService ?? throw new ArgumentNullException(nameof(groupsService));
            this.appSettingsService = appSettingsService ?? throw new ArgumentNullException(nameof(appSettingsService));
            this.blobStorageProvider = blobStorageProvider ?? throw new ArgumentException(nameof(blobStorageProvider));
            this.Configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

        }

        [HttpPut]
        public async Task<IActionResult> UpdateDefaultsDataAsync([FromBody] DefaultsData defaults)
        {

            if (!string.IsNullOrEmpty(defaults.HeaderLogoLink) && defaults.HeaderLogoLink.StartsWith(Constants.ImageBase64Format))
            {
                await this.notificationDataRepository.SaveImageAsync(this.DEFAULT_HEADER_LOGO, defaults.HeaderLogoLink);
            }else if (!string.IsNullOrEmpty(defaults.HeaderLogoLink) && defaults.HeaderLogoLink.Equals("DELETE")) {

                await this.notificationDataRepository.DeleteImageAsync(this.DEFAULT_HEADER_LOGO);
            }

            if (!string.IsNullOrEmpty(defaults.LogoLink) && defaults.LogoLink.StartsWith(Constants.ImageBase64Format))
            {
                await this.notificationDataRepository.SaveImageAsync(this.DEFAULT_LOGO_BLOB_NAME, defaults.LogoLink);

            } else if (!string.IsNullOrEmpty(defaults.LogoLink) && defaults.LogoLink.Equals("DELETE"))
            {
                await this.notificationDataRepository.DeleteImageAsync(this.DEFAULT_LOGO_BLOB_NAME);
            }


            if (!string.IsNullOrEmpty(defaults.BannerLink) && defaults.BannerLink.StartsWith(Constants.ImageBase64Format))
            {
                await this.notificationDataRepository.SaveImageAsync(this.DEFAULT_BANNER_BLOB_NAME, defaults.BannerLink);

            }
            else if (!string.IsNullOrEmpty(defaults.BannerLink) && defaults.BannerLink.Equals("DELETE"))
            {
                await this.notificationDataRepository.DeleteImageAsync(this.DEFAULT_BANNER_BLOB_NAME);
            }

            if (!string.IsNullOrEmpty(defaults.HeaderText))
            {
                await this.appSettingsService.SetHeaderText(defaults.HeaderText);
            }
            return this.Ok();
        }

        [HttpGet]

        public async Task<ActionResult<DefaultsData>> GetDefaultDataAsync()
        {
            var logoImage = await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_LOGO_BLOB_NAME);
            var logoLink = logoImage != string.Empty ? "data:image/jpeg;base64," + logoImage : null;

            var bannerImage = await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_BANNER_BLOB_NAME);
            var bannerLink = bannerImage != string.Empty ? "data:image/jpeg;base64," + bannerImage : null;

            var headerLogoImage = await this.blobStorageProvider.DownloadBase64ImageAsync(this.DEFAULT_HEADER_LOGO);
            var headerLogoLink = headerLogoImage != string.Empty ? "data:image/jpeg;base64," + headerLogoImage : null;


            var headerText = await this.appSettingsService.GetHeaderText() ?? this.Configuration.GetValue<string>("REACT_APP_HEADERTEXT");
            var result = new DefaultsData
            {
                LogoFileName = this.DEFAULT_LOGO_BLOB_NAME,
                LogoLink = logoLink,
                BannerFileName = this.DEFAULT_BANNER_BLOB_NAME,
                BannerLink = bannerLink,
                HeaderLogoFileName = this.DEFAULT_HEADER_LOGO,
                HeaderLogoLink = headerLogoLink,
                HeaderText = headerText,

            };



            return this.Ok(result);
        }

       


    }
}