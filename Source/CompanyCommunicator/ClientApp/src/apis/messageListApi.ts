// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { getBaseUrl } from "../configVariables";
import axios from "./axiosJWTDecorator";

let baseAxiosUrl = getBaseUrl() + "/api";

export const getSentNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications";
    return await axios.get(url);
};

export const getDraftNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.get(url);
};

export const getCardTemplates = async (): Promise<any> => {
    let url = baseAxiosUrl + "/cardtemplates";
    return await axios.get(url);
}

export const getCardTemplate = async (name: any): Promise<any> => {
    let url = baseAxiosUrl + "/cardtemplates/" + name;
    return await axios.get(url);
}

export const createCardTemplate = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/cardtemplates";
    return await axios.put(url, payload);
}

export const updateCardTemplate = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/cardtemplates";
    return await axios.put(url, payload);
}

export const getDefaultData = async (): Promise<any> => {
    let url = baseAxiosUrl + "/defaultdata";
    return await axios.get(url);
}

export const updateDefaultData = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/defaultdata";
    return await axios.put(url, payload);
}

export const verifyGroupAccess = async (): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/verifyaccess";
    return await axios.get(url, false);
};

export const getGroups = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/" + id;
    return await axios.get(url);
};

export const searchGroups = async (query: string): Promise<any> => {
    let url = baseAxiosUrl + "/groupdata/search/" + query;
    return await axios.get(url);
};

export const exportNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/exportnotification/export";
    return await axios.put(url, payload);
};

export const getSentNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications/" + id;
    return await axios.get(url);
};

export const getDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.get(url);
};

export const deleteDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.delete(url);
};

export const duplicateDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/duplicates/" + id;
    return await axios.post(url);
};

export const sendDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications";
    return await axios.post(url, payload);
};

export const updateDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.put(url, payload);
};

export const createDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.post(url, payload);
};

export const getTeams = async (): Promise<any> => {
    let url = baseAxiosUrl + "/teamdata";
    return await axios.get(url);
};

export const cancelSentNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications/cancel/" + id;
    return await axios.post(url);
};

export const getConsentSummaries = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/consentSummaries/" + id;
    return await axios.get(url);
};

export const sendPreview = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/previews";
    return await axios.post(url, payload);
};

export const getAppId = async (): Promise<any> => {
    let url = baseAxiosUrl + "/sentNotifications/appid";
    return await axios.get(url);
}
// Get sent notification by activity ID
export const getSentNotificationByActivityId = async (activityId: string): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications/byactivity/" + activityId;
    var res = await axios.get(url);
    return res;
};

export const getAuthenticationConsentMetadata = async (
    windowLocationOriginDomain: string,
    login_hint: string
): Promise<any> => {
    let url = `${baseAxiosUrl}/authenticationMetadata/consentUrl?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${login_hint}`;
    return await axios.get(url, undefined, false);
};
