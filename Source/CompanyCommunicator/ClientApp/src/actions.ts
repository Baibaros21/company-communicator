// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import {
    getDraftNotifications,
    getGroups,
    getSentNotifications,
    getTeams,
    searchGroups,
    verifyGroupAccess,
    getCardTemplates
} from "./apis/messageListApi";
import { formatDate } from "./i18n";
import {
    draftMessages,
    groups,
    isDraftMessagesFetchOn,
    isSentMessagesFetchOn,
    isCardTemplatesFetchOn,
    queryGroups,
    sentMessages,
    teamsData,
    verifyGroup,
    cardTemplates

} from "./messagesSlice";
import { store } from "./store";

type ITemplates = {
    name: string;
    card: any;
}
type Notification = {
    createdDateTime: string;
    failed: number;
    id: string;
    isCompleted: boolean;
    sentDate: string;
    sendingStartedDate: string;
    sendingDuration: string;
    succeeded: number;
    throttled: number;
    title: string;
    seen: number;
    like: number;
    heart: number;
    surprise: number;
    laugh: number;
    totalMessageCount: number;
    createdBy: string;
};
export const GetAllCardTemplatesAction = (dispatch: typeof store.dispatch) => {
    CardTemplatesFetchStatusAction(dispatch, true);
    getCardTemplates().then((response) => {

        const cardTemplatesList: ITemplates = response.data;
        dispatch(cardTemplates({ type: "FETCH_TEMPLATES", payload: cardTemplatesList }));

    }).finally(() => {
        CardTemplatesFetchStatusAction(dispatch, false);
    });
}

export const GetSentMessagesAction = (dispatch: typeof store.dispatch) => {
    SentMessageFetchStatusAction(dispatch, true);
    getSentNotifications()
        .then((response) => {
            const notificationList: Notification[] = response?.data || [];
            notificationList.forEach((notification) => {
                notification.sendingStartedDate = formatDate(notification.sendingStartedDate);
                notification.sentDate = formatDate(notification.sentDate);
            });
            dispatch(sentMessages({ type: "FETCH_MESSAGES", payload: notificationList || [] }));
        })
        .finally(() => {
            SentMessageFetchStatusAction(dispatch, false);
        });
};

export const GetSentMessagesSilentAction = (dispatch: typeof store.dispatch) => {
    getSentNotifications().then((response) => {
        const notificationList: Notification[] = response?.data || [];
        notificationList.forEach((notification) => {
            notification.sendingStartedDate = formatDate(notification.sendingStartedDate);
            notification.sentDate = formatDate(notification.sentDate);
        });
        dispatch(sentMessages({ type: "FETCH_MESSAGES", payload: notificationList || [] }));
    });
};

export const GetDraftMessagesAction = (dispatch: typeof store.dispatch) => {
    DraftMessageFetchStatusAction(dispatch, true);
    getDraftNotifications()
        .then((response) => {
            dispatch(draftMessages({ type: "FETCH_DRAFT_MESSAGES", payload: response?.data || [] }));
        })
        .finally(() => {
            DraftMessageFetchStatusAction(dispatch, false);
        });
};

export const GetDraftMessagesSilentAction = (dispatch: typeof store.dispatch) => {
    getDraftNotifications().then((response) => {
        dispatch(draftMessages({ type: "FETCH_DRAFT_MESSAGES", payload: response?.data || [] }));
    });
};

export const GetTeamsDataAction = (dispatch: typeof store.dispatch) => {
    getTeams().then((response) => {
        dispatch(teamsData({ type: "GET_TEAMS_DATA", payload: response?.data || [] }));
    });
};

export const GetGroupsAction = (dispatch: typeof store.dispatch, payload: { id: number }) => {
    getGroups(payload.id).then((response) => {
        dispatch(groups({ type: "GET_GROUPS", payload: response?.data || [] }));
    });
};

export const SearchGroupsAction = (dispatch: typeof store.dispatch, payload: { query: string }) => {
    searchGroups(payload.query).then((response) => {
        dispatch(queryGroups({ type: "SEARCH_GROUPS", payload: response?.data || [] }));
    });
};

export const VerifyGroupAccessAction = (dispatch: typeof store.dispatch) => {
    verifyGroupAccess()
        .then((response) => {
            dispatch(verifyGroup({ type: "VERIFY_GROUP_ACCESS", payload: true }));
        })
        .catch((error) => {
            const errorStatus = error.response.status;
            if (errorStatus === 403) {
                dispatch(verifyGroup({ type: "VERIFY_GROUP_ACCESS", payload: false }));
            } else {
                throw error;
            }
        });
};

export const DraftMessageFetchStatusAction = (dispatch: typeof store.dispatch, payload: boolean) => {
    dispatch(isDraftMessagesFetchOn({ type: "DRAFT_MESSAGES_FETCH_STATUS", payload }));
};

export const SentMessageFetchStatusAction = (dispatch: typeof store.dispatch, payload: boolean) => {
    dispatch(isSentMessagesFetchOn({ type: "SENT_MESSAGES_FETCH_STATUS", payload }));
};
export const CardTemplatesFetchStatusAction = (dispatch: typeof store.dispatch, payload: boolean) => {
    dispatch(isCardTemplatesFetchOn({ type: "CARD_TEMPLATES_FETCH_STATUS", payload }));
};
