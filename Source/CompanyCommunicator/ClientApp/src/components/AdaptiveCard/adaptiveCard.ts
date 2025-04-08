// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.


import * as AdaptiveCards from 'adaptivecards';
import MarkdownIt from 'markdown-it';
import { getBaseUrl } from '../../configVariables';
import { TemplateSelection } from "../../store";
import { getCardTemplates, getCardTemplate, updateCardTemplate } from "../../apis/messageListApi";


AdaptiveCards.AdaptiveCard.onProcessMarkdown = function (text, result) {
    result.outputHtml = new MarkdownIt().render(text);
    result.didProcess = true;
};

interface ITemplateState {
    template: TemplateSelection,
    card: string
}

export const getInitAdaptiveCard = (titleText: string = "title", type: string = TemplateSelection.Default) => {

    switch (type) {

        case "videoPlayer": {

            return {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.6",
                "body": [
                    {
                        "type": "TextBlock",
                        "weight": "Bolder",
                        "spacing": "None",
                        "text": "Title",
                        "size": "Large",
                        "wrap": true,
                        "horizontalAlignment": "Center",
                        "name": "title"
                    },
                    {
                        "type": "Media",
                        "poster": "https://adaptivecards.io/content/poster-video.png",
                        "name": "video",
                        "sources": [
                            {
                                "mimeType": "video/mp4",
                                "url": `https://samplelib.com/lib/preview/mp4/sample-5s.mp4`
                            }
                        ]

                    },
                ]

            };
        }

        case "viewDefaults": {
            return {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "Image",
                        "url": getBaseUrl() + "/image/Logo.png",
                        "altText": "Image",
                        "horizontalAlignment": "Center",
                        "name": "logo",
                        "height": "50px",
                        "separator": true,
                        "size": "Stretch"
                    },
                    {
                        "type": "Image",
                        "url": getBaseUrl() + "/image/banner.png",
                        "spacing": "Small",
                        "horizontalAlignment": "Center",
                        "height": "50px",
                        "name": "banner",
                        "separator": true
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.5"
            };

        }


    };
}

export const saveAdaptiveCard = async (card: any, template: TemplateSelection) => {
    const payload: ITemplateState = {
        template: template,
        card: JSON.stringify(card)
    }
    await updateCardTemplate(payload);

}


var getProperty = (body: any, value: string): any => {
    return body.filter((prop: any) =>
        prop.name === value
    );
}

export const getCardTitle = (card: any) => {

    var filteredcomp = getProperty(card.body, "title");
    if (filteredcomp.length > 0) {
        return filteredcomp.text
    }
};


export const setCardTitle = (card: any, title: string) => {
    var filteredcomp = getProperty(card.body, "title");
    console.log("In title");
    if (filteredcomp.length > 0) {

        filteredcomp[0].text = title

    }

};


export const getDeptTitle = (card: any) => {

    var filteredcomp = getProperty(card.body, "department");
    if (filteredcomp.length > 0) {
        return filteredcomp[0].text
    }
};

export const setCardDeptTitle = (card: any, title?: string) => {
    var filteredcomp = getProperty(card.body, "department");
    if (filteredcomp.length > 0) {

        filteredcomp[0].text = title

    }

};

export const getCardImageLink = (card: any) => {
    const filteredcomp = getProperty(card.body, "image");
    if (filteredcomp.length > 0) {
        return filteredcomp[0].url
    }
};

export const setCardImageLink = (card: any, imageLink?: string) => {
    const filteredcomp = getProperty(card.body, "image");
    if (filteredcomp.length > 0) {
        filteredcomp[0].url = imageLink
    }
};
export const setCardVideoPlayerUrl = (card: any, videoLink?: string) => {

    const filteredcomp = getProperty(card.body, "video");
    if (filteredcomp.length > 0) {
        /*filteredcomp[0].selectAction.url = videoLink;*/
        try {
            filteredcomp[0].selectAction.data.videoId = videoLink;
        } catch (error) {
            console.log("Error in setting video player task module url");
        }
    }
};

export const setCardVideoUrl = (card: any, videoLink?: string) => {

    const filteredcomp = getProperty(card.body, "video");
    if (filteredcomp.length > 0) {
        filteredcomp[0].sources[0].url = videoLink;
    }
};

export const setCardVideoPoster = (card: any, imageLink?: string) => {

    const filteredcomp = getProperty(card.body, "video");
    if (filteredcomp.length > 0) {
        filteredcomp[0].poster = imageLink;
    }
};

export const setCardVideoPlayerPoster = (card: any, imageLink?: string) => {

    const filteredcomp = getProperty(card.body, "video");
    if (filteredcomp.length > 0) {
        filteredcomp[0].url = imageLink;
    }
}

export const getCardSummary = (card: any) => {
    const filteredcomp = getProperty(card.body, "summary");
    if (filteredcomp.length > 0) {
        return filteredcomp[0].text
    }
};

export const setCardSummary = (card: any, summary?: string) => {
    const filteredcomp = getProperty(card.body, "summary");
    if (filteredcomp.length > 0) {
        filteredcomp[0].text = summary
    }
};

export const getCardAuthor = (card: any) => {
    const filteredcomp = getProperty(card.body, "author");
    if (filteredcomp.length > 0) {
        return filteredcomp[0].text
    }
};

export const setCardAuthor = (card: any, author?: string) => {
    const filteredcomp = getProperty(card.body, "author");
    if (filteredcomp.length > 0) {
        filteredcomp[0].text = author
    }
};


export const setCardLogo = (card: any, imageLink?: string) => {
    const filteredcomp = getProperty(card.body, "logo");
    if (filteredcomp.length > 0) {
        filteredcomp[0].url = imageLink
    }
};

export const setCardBanner = (card: any, imageLink?: string) => {
    const filteredcomp = getProperty(card.body, "banner");
    if (filteredcomp.length > 0) {
        filteredcomp[0].url = imageLink
    }
};

export const getCardBtnTitle = (card: any) => {
    return card.actions[0].title;
};

export const getCardBtnLink = (card: any) => {
    return card.actions[0].url;
};

export const setCardBtn = (card: any, buttonTitle?: string, buttonLink?: string) => {
    if (buttonTitle && buttonLink) {
        card.actions = [
            {
                type: 'Action.OpenUrl',
                title: buttonTitle,
                url: buttonLink,
            },
        ];
    } else {
        delete card.actions;
    }
};
