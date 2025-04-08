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


    const GetSentNotification = async (activityId: string) => {
        await getSentNotificationByActivityId(activityId)
            .then((response) => {
                if (response && response.data) {
                    const notification = response.data;
                    // Set video URL if available
                    if (notification.videoLink) {
                        setVideoUrl(notification.videoLink);
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
        {error && <div className="error-message">{error}</div>}
              {!loader && !error && videoUrl &&(
          <div className="video-container">
            <video
              controls
              width="100%"
              height="auto"
              poster={posterUrl}
              onLoadedData={() => setLoader(false)}
              onError={handleVideoError}
            >
              <source src={videoUrl} type="video/mp4" />
              {t("videoNotSupported")}
            </video>
          </div>
              )}
         {!loader && !error && !videoUrl && (
          <div className="error-message">No video URL available</div>
        )}
      </div>
    </>
  );
};
