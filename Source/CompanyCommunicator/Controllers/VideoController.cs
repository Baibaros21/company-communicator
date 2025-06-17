// VideoController.cs
using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Blob;
using Azure.Storage.Blobs;
using Azure.Storage.Sas;
using Azure.Storage.Blobs.Models;
using Microsoft.Teams.Apps.CompanyCommunicator.Common.Clients;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    [Route("api/video")]
    
    public class VideoController : ControllerBase
    {
        private readonly IBlobStorageProvider _blobStorageProvider;
        private readonly ILogger<VideoController> _logger;

        public VideoController(
            IBlobStorageProvider blobStorageProvider,
            ILogger<VideoController> logger)
        {
            _blobStorageProvider = blobStorageProvider ?? throw new ArgumentNullException(nameof(blobStorageProvider));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadVideo(IFormFile file, string title)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file provided");
            }

            title = title ?? file.FileName;

            // Validate file type
            string contentType = file.ContentType.ToLower();
            if (!contentType.StartsWith("video/"))
            {
                return BadRequest("Only video files are allowed");
            }

            try
            {

                var videoSasUri = await this._blobStorageProvider.GetUploadBlobSASUriAsync(title, file);

                return Ok(new { url = videoSasUri, name = title });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading video");
                return StatusCode(500, $"An error occurred while uploading the video :{ex.Message}");
            }
        }


    }
}
