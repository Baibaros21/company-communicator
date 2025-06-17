import * as React from "react";
import "./videoPlayer.scss";
import { useTranslation } from "react-i18next";
import { useLocation } from "react-router-dom";
import { Spinner } from "@fluentui/react-components";
import { getSentNotificationByActivityId } from "../../apis/messageListApi";

export const VideoPlayer = () => {
  const { t } = useTranslation();

  // Get URL and activity ID from both route params and query params
  const location = useLocation();
  const queryParams = new URLSearchParams(location.search);
  const urlQuery = queryParams.get("url") || "";
  const posterQuery = queryParams.get("poster") || "";
  const activityIdQuery = queryParams.get("activityId") || "";

  const [useIfrmae, setUseIframe] = React.useState(false);
  const [loader, setLoader] = React.useState(true);
  const [error, setError] = React.useState("");
  const [videoUrl, setVideoUrl] = React.useState(urlQuery || "");
  const [posterUrl, setPosterUrl] = React.useState(
    posterQuery || "https://adaptivecards.io/content/poster-video.png"
    );
  const [activityId, setActivityId] = React.useState(activityIdQuery || "");



    React.useEffect(() => {
        setActivityId(activityIdQuery);
    }, [activityIdQuery])

  React.useEffect(() => {
    // If activity ID is provided, fetch notification details
      if (activityId) {
          GetSentNotification(activityId);
    } else if (!videoUrl) {
      console.warn("No video URL or activity ID provided");
      setLoader(false);
    } else {
      // If video URL is already provided, just set loader to false
      setLoader(false);
      }
  }, [activityId, videoUrl]); // Remove videoUrl from dependencies to avoid infinite loops


    const getYouTubeEmbedUrl = (url: string): string => {
        if (!url) return url;

        // Regular expression to match YouTube URL patterns
        const youtubeRegex = /^(https?:\/\/)?(www\.)?(youtube\.com\/watch\?v=|youtu\.be\/)([a-zA-Z0-9_-]{11})(?:\S*)?$/;
        const match = url.match(youtubeRegex);

        // If it's a valid YouTube URL, extract the video ID and return the embed URL
        if (match && match[4]) {
            const videoId = match[4];
            return `https://www.youtube.com/embed/${videoId}`;
        }

        // Return the original URL if it's not a valid YouTube URL
        return url;
    };
    const GetSentNotification = async (activityId: string) => {
        await getSentNotificationByActivityId(activityId)
            .then((response) => {
                if (response && response.data) {
                    const notification = response.data;
                    // Set video URL if available
                    if (notification.videoLink) {
                        // Process the URL to handle YouTube links
                        const processedUrl = getYouTubeEmbedUrl(notification.videoLink);
                        setVideoUrl(processedUrl);

                        const isYouTubeEmbed = processedUrl.includes("youtube.com/embed/");

                        // Set iframe mode for YouTube embeds and non-MP4 videos
                        setUseIframe(isYouTubeEmbed);
                    }

                    // Set poster URL if available
                    if (notification.posterLink) {
                        setPosterUrl(notification.posterLink);
                    }
                }
            })
            .catch((error) => {
                console.error("Failed to fetch notification details:", error);
                setError("Failed to fetch video information from notification");
            })
            .finally(() => {
                setLoader(false);
            });
    }
  const handleVideoError = () => {
      setLoader(false);
      setUseIframe(true);
    };
  
    const handleIframeError = () => {
        setLoader(false);
        setError("Failed to load video. Please check the URL and try again.");
    };
  return (
    <>
      <div className="taskModule">
        {loader && (
          <div className="spinner-container">
            <Spinner />
          </div>
        )}

              {!loader && !error && videoUrl && (useIfrmae ?
                   (<div className="video-container">
                      <iframe
                          src={videoUrl}
                          frameBorder="0"
                          allowFullScreen
                          title="Video Player"
                          onError={handleIframeError}
                          allow="autoplay; encrypted-media"
                          
                      />
                  </div> )
              : ( <div className="video-container">
                  <video
                          controls
                      poster={posterUrl}
                      onLoadedData={() => setLoader(false)}
                      onError={handleVideoError}
                  >
                      <source src={videoUrl} type="video/mp4" />
                      {t("videoNotSupported")}
                  </video>
              </div>
              ))}
         {!loader && !error && !videoUrl && (
          <div className="error-message">No video URL available</div>
        )}
      </div>
    </>
  );
};
