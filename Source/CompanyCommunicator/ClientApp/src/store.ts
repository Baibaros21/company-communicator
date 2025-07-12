// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TypedUseSelectorHook, useDispatch, useSelector } from "react-redux";
import { configureStore } from "@reduxjs/toolkit";
import messagesReducer from "./messagesSlice";

export const store = configureStore({
    reducer: { messages: messagesReducer },
    devTools: process.env.NODE_ENV !== "production",
});

// Infer the `RootState` and `AppDispatch` types from the store itself
export type RootState = ReturnType<typeof store.getState>;

// Inferred type: {posts: PostsState, comments: CommentsState, users: UsersState}
export type AppDispatch = typeof store.dispatch;

export const useAppDispatch = () => useDispatch<AppDispatch>();
export const useAppSelector: TypedUseSelectorHook<RootState> = useSelector;
export enum TemplateSelection {
    Default = 'Default',
    Default_ar = 'Default Arabic',

    department = 'Department Message',
    department_ar = 'Department Message Arabic',


    departmentPoster_ar = 'department message with poster Arabic',
    departmentPoster = 'department message with poster',

    departmentVideo = 'department message with video',
    departmentVideo_ar = 'department message with video Arabic',

    infromational = 'Informational',

    infoPoster = 'Informational with Poster',
    infoPoster_ar = 'Informational with Poster Arabic',

    infoVideo = 'Informational with Video',
    infoVideo_ar = 'Informational with Video Arabic',

    video = "Video",
    
    
}
interface TemplateSelectionItem {
    key: TemplateSelection;
    displayName: string;
    author?: boolean;
    summary?: boolean;
    poster?: boolean;
    video?: boolean;
    department?: boolean;
}

export const TemplateItems: TemplateSelectionItem [] = [
        // General templates
        {
            key: TemplateSelection.Default,
            displayName: 'General Announcement',
            author: true,
            summary: true,
            poster: true,
            video: false,
            department: false
        },

        // Department templates (English)
        {
            key: TemplateSelection.department,
            displayName: 'Department Announcement',
            author: true,
            summary: true,
            poster: false,
            video: false,
            department: true
        },
        {
            key: TemplateSelection.departmentPoster,
            displayName: 'Department Announcement with Poster',
            author: true,
            summary: true,
            poster: true,
            video: false,
            department: true
        },
        {
            key: TemplateSelection.departmentVideo,
            displayName: 'Department Announcement with Video',
            author: true,
            summary: true,
            poster: false,
            video: true,
            department: true
        },

        // Informational templates (English)
      {
            key: TemplateSelection.infromational,
            displayName: 'Poster Only',
            author: false,
            summary: false,
            poster: true,
            video: false,
            department: false
    },
        /*
        {
            key: TemplateSelection.infoPoster,
            displayName: 'Informational Announcement with Poster',
            author: false,
            summary: true,
            poster: true,
            video: false,
            department: false
        },
        {
            key: TemplateSelection.infoVideo,
            displayName: 'Informational Announcement with Video',
            author: false,
            summary: true,
            poster: false,
            video: true,
            department: false
        },*/

        // Video template
        {
            key: TemplateSelection.video,
            displayName: 'Video Only',
            author: false,
            summary: false,
            poster: false,
            video: true,
            department: false
        },

        // General Arabic
        {
            key: TemplateSelection.Default_ar,
            displayName: 'General Arabic Announcement',
            author: true,
            summary: true,
            poster: true,
            video: false,
            department: false
        },

        // Department templates (Arabic)
        {
            key: TemplateSelection.department_ar,
            displayName: 'Department Arabic Announcement',
            author: true,
            summary: true,
            poster: false,
            video: false,
            department: true
        },
        {
            key: TemplateSelection.departmentPoster_ar,
            displayName: 'Department Arabic Announcement with Poster',
            author: true,
            summary: true,
            poster: true,
            video: false,
            department: true
        },
        {
            key: TemplateSelection.departmentVideo_ar,
            displayName: 'Department Arabic Announcement with Video',
            author: true,
            summary: true,
            poster: false,
            video: true,
            department: true
        },

        // Informational templates (Arabic)
/*        {
            key: TemplateSelection.infoPoster_ar,
            displayName: 'Informational Arabic Announcement with Poster',
            author: false,
            summary: true,
            poster: true,
            video: false,
            department: false
        },
        {
            key: TemplateSelection.infoVideo_ar,
            displayName: 'Informational Arabic Announcement with Video',
            author: false,
            summary: true,
            poster: false,
            video: true,
            department: false
        },*/
];

export interface IMessageState {
    title: string;
    id?: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
    succeeded?: string;
    template:  TemplateSelection;
    failed?: string;
    unknown?: string;
    canceled?: string;
    sentDate?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    teamNames?: string[];
    rosterNames?: string[];
    groupNames?: string[];
    allUsers?: boolean;
    sendingStartedDate?: string;
    sendingDuration?: string;
    errorMessage?: string;
    warningMessage?: string;
    canDownload?: boolean;
    sendingCompleted?: boolean;
    createdBy?: string;
    seen?: number;
    department?: string;
    posterLink?: string;
    videoLink?: string;
    isMsgDataUpdated?: boolean;
    card?: string;
    teams?: any[] | undefined;
    rosters?: any[];
    groups?: any[];
    throttled?: number;
    isDraftMsgUpdated?: boolean;
}