// VideoUploadTask.tsx - Create this file in CompanyCommunicator/ClientApp/src/components/VideoUploadTask/
import * as React from "react";
import "./VideoUplaodTask.scss";
import { useTranslation } from "react-i18next";
import {
  Button,
  Input,
  Spinner,
  Text,
  makeStyles,
  tokens,
  Field,
  Label
} from "@fluentui/react-components";
import { ArrowUpload24Regular, VideoClip24Regular } from "@fluentui/react-icons";
import * as microsoftTeams from "@microsoft/teams-js";
import axios from "../../apis/axiosJWTDecorator";
import i18n from "../../i18n";
import { getBaseUrl } from "../../configVariables";

const useStyles = makeStyles({
  container: {
    maxWidth: "800px",

  },
  fileInput: {
    display: "none",
  },
  uploadButton: {
    marginTop: "16px",
  },
  previewContainer: {
    marginTop: "24px",

    backgroundColor: tokens.colorNeutralBackground1,
  },
  videoContainer: {
    marginTop: "16px",
    width: "100%",

    boxShadow: "0 2px 4px rgba(0, 0, 0, 0.1)",
  },
  progressOuter: {
    height: "8px",
    backgroundColor: tokens.colorNeutralBackground3,

  },
  progressInner: {
    height: "100%",
    backgroundColor: tokens.colorBrandBackground,

  },
  urlField: {
    marginTop: "24px",
  },
  buttonContainer: {
    display: "flex",
    justifyContent: "flex-end",
    marginTop: "24px",

  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: "8px",
  },
});

export const VideoUploadTask = () => {
  const { t } = useTranslation();
  const styles = useStyles();
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  
  const [uploadProgress, setUploadProgress] = React.useState(0);
  const [isUploading, setIsUploading] = React.useState(false);
  const [videoFile, setVideoFile] = React.useState<File | null>(null);
  const [videoUrl, setVideoUrl] = React.useState("");
  const [previewUrl, setPreviewUrl] = React.useState("");
  const [error, setError] = React.useState("");
  const [videoTitle, setVideoTitle] = React.useState("");
  
  const handleFileButtonClick = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };
  
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    // Validate file type
    if (!file.type.startsWith("video/")) {
      setError(t("OnlyVideoFilesAllowed"));
      return;
    }
    
    setVideoFile(file);
    setVideoTitle(file.name);
    setError("");
    
    // Create local preview URL
    const objectUrl = URL.createObjectURL(file);
    setPreviewUrl(objectUrl);
    
    // Clean up preview URL when component unmounts
    return () => URL.revokeObjectURL(objectUrl);
  };
  
  const handleUpload = async () => {
    if (!videoFile) {
      setError(t("PleaseSelectVideoFirst"));
      return;
    }
    
    setError("");
    setIsUploading(true);
    setUploadProgress(0);
    
    try {
      // Create FormData
      const formData = new FormData();
      formData.append("file", videoFile);
        formData.append("filename", videoFile.name);
        let baseAxiosUrl = getBaseUrl() + "/api";
        let url = baseAxiosUrl + "/video/upload";

        // Get auth token from Teams
        const token = await new Promise<string>((resolve, reject) => {
            const authTokenRequest = {
                successCallback: (token: string) => {
                    resolve(token);
                },
                failureCallback: (error: string) => {
                    console.error("Error getting auth token:", error);
                    reject(new Error("Failed to get authentication token"));
                },
                resources: []
            };

            microsoftTeams.authentication.getAuthToken(authTokenRequest);
        });

        // Create headers
        const headers: HeadersInit = {};
        if (token) {
            headers['Authorization'] = `Bearer ${token}`;
        }
        headers['Accept-Language'] = i18n.language || 'en-US';

        // Make the fetch request
        const response = await fetch(url, {
            method: 'POST',
            body: formData,
            headers: headers,
            credentials: 'include',
        });

        // Check if response is successful
        if (!response.ok) {
            throw new Error(`Server responded with ${response.status}: ${response.statusText}`);
        }

        // Parse the response
        const data = await response.json();
 
      // Set the video URL from the response
        setVideoUrl(data.url);
      
      // Wait a moment to show the completed progress bar
      setTimeout(() => {
        setIsUploading(false);
      }, 500);
    } catch (error) {
        const err = error as any; // Type assertion after catching the error
        console.error("Error uploading video:", err);
        setError(err.response?.data || t("FailedToUploadVideo"));
        setIsUploading(false);
    }
  };
  
  const handleCancel = () => {
    microsoftTeams.tasks.submitTask();
  };
  
  const handleDone = () => {
    if (videoUrl) {
      microsoftTeams.tasks.submitTask({ videoUrl: videoUrl });
    } else {
      setError(t("PleaseUploadVideoFirst"));
    }
  };
  
  const handleUrlChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setVideoTitle(e.target.value);
  };
  

  
  return (
    <div className={styles.container}>
      <Text  weight="semibold">{t("UploadVideo")}</Text>
      
      <Field label={t("VideoTitle")} required>
        <Input 
          value={videoTitle}
          onChange={handleUrlChange}
          placeholder={t("EnterVideoTitle")}
        />
      </Field>
      
      <div>
        <Label>{t("ChooseVideoFile")}</Label>
        <input
          type="file"
          className={styles.fileInput}
          ref={fileInputRef}
          accept="video/*"
          onChange={handleFileChange}
        />
        <Button
          icon={<ArrowUpload24Regular />}
          onClick={handleFileButtonClick}
          disabled={isUploading}
        >
          {t("SelectVideoFile")}
        </Button>
        
        {videoFile && (
          <Text >
            {t("SelectedFile")}: {videoFile.name} ({Math.round(videoFile.size / 1024 / 1024 * 100) / 100} MB)
          </Text>
        )}
      </div>
      
      {isUploading && (
        <>
          <div className={styles.progressOuter}>
            <div 
              className={styles.progressInner} 
              style={{ width: `${uploadProgress}%` }}
            />
          </div>
          <Text >{uploadProgress}% {t("Uploaded")}</Text>
        </>
      )}
      
      <Button
        className={styles.uploadButton}
        appearance="primary"
        icon={<ArrowUpload24Regular />}
        disabled={!videoFile || isUploading}
        onClick={handleUpload}
      >
        {isUploading ? t("Uploading") : t("UploadVideo")}
      </Button>
     
      
      {error && (
        <Text className={styles.errorText} >
          {error}
        </Text>
      )}
      
      {previewUrl && !isUploading && (
        <div className={styles.previewContainer}>
          <Text weight="semibold">{t("VideoPreview")}</Text>
          <div className={styles.videoContainer}>
            <video
              controls
              width="100%"
              height="auto"
              src={previewUrl}
            />
          </div>
        </div>
      )}
      
      {videoUrl && !previewUrl && (
        <div className={styles.previewContainer}>
          <Text weight="semibold">{t("UploadedVideoPreview")}</Text>
          <div className={styles.videoContainer}>
              <video
                controls
                width="auto"
                height="400px"
                src={videoUrl}
              />
          </div>
        </div>
      )}
      
      <div className={styles.buttonContainer}>
        <Button onClick={handleCancel}>{t("Cancel")}</Button>
        <Button 
          appearance="primary" 
          onClick={handleDone}
          disabled={!videoUrl && !previewUrl}
        >
          {t("Done")}
        </Button>
      </div>
    </div>
  );
};

export default VideoUploadTask;
