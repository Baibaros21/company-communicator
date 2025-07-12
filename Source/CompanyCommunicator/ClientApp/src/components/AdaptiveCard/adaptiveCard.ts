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
export const getCardImageEmbedLink = (card: any): string => {
    // Validate input parameters
    if (!card || !card.body) {
        return '';
    }

    // Find image components in the card body
    const imageComponents = getProperty(card.body, "image");

    // Check if we found any image components and if the first one has a selectAction with a URL
    if (imageComponents?.length > 0 && imageComponents[0]?.selectAction?.url) {
        return imageComponents[0].selectAction.url;
    }

    // Return empty string if no image with embed link found
    return '';
};
export const setCardImageEmbedLink = (card: any, url?: string): void => {
    // Validate input parameters
    if (!card || !card.body || !url) {
        return;
    }

    // Find image components in the card body
    const imageComponents = getProperty(card.body, "image");

    // Check if we found any image components with a selectAction
    if (imageComponents?.length > 0) {
        const imageComponent = imageComponents[0];

        // Create selectAction if it doesn't exist
        if (!imageComponent.selectAction && url) {
            imageComponent.selectAction = {
                type: 'Action.OpenUrl',
                url: url
            };
        }
        // Update existing selectAction url
        else if (imageComponent.selectAction) {
            imageComponent.selectAction.url = url;
        }
    }
};


export const updateCardVideoElement = (card: any, videoLink?: string, posterLink?: string) => {

    const videoIndex = card.body.findIndex((item: any) => item.name === "video" || item.type === "Video");


    if (videoIndex > -1) {
        if(card.body[videoIndex].type !== "Image"){
            card.body[videoIndex] = {
                "type": "Image",
                "name": "video",
                "url": "/image/imagePlaceholder.png",
                "horizontalAlignment": "Center",
                "selectAction": {
                    "type": "Action.Submit",
                    "data": {
                        "msteams": {
                            "type": "task/fetch",
                            "taskModule": {
                                "title": "Your Task Module Title",
                                "height": "800px",
                                "width": "900px"
                            }
                        },
                        "videoId": "videoId"
                    }
                },
                "separator": true,
                "size": "Stretch"
            }
        }
        if (videoLink) {
                setCardVideoUrl(card, videoLink);
            }
        if (posterLink) {
            setCardVideoPoster(card, posterLink);
            }
        
    }
}


//export const setCardVideoPlayerUrl = (card: any, videoLink?: string) => {

//    const filteredcomp = getProperty(card.body, "video");
//    if (filteredcomp.length > 0) {
//        /*filteredcomp[0].selectAction.url = videoLink;*/
//        try {
//            filteredcomp[0].selectAction.data.videoId = videoLink;
//        } catch (error) {
//            console.log("Error in setting video player task module url");
//        }
//    }

//const filteredcomp = getProperty(card.body, "video");
//if (filteredcomp.length > 0) {
//    filteredcomp[0].sources[0].url = videoLink;
//}
//};
//export const setCardVideoPlayerPoster = (card: any, imageLink?: string) => {

//    const filteredcomp = getProperty(card.body, "video");
//    if (filteredcomp.length > 0) {
//        filteredcomp[0].url = imageLink;
//    }
//const filteredcomp = getProperty(card.body, "video");
//if (filteredcomp.length > 0) {
//    filteredcomp[0].poster = imageLink;
//}
//}
export const setCardVideoUrl = (card: any, videoLink?: string) => {

        const filteredcomp = getProperty(card.body, "video");
        if (filteredcomp.length > 0) {
          
            try {
                filteredcomp[0].selectAction.data.videoId = videoLink;
            } catch (error) {
                console.log("Error in setting video player task module url");
            }
        }
};

export const setCardVideoPoster = (card: any, imageLink?: string) => {

        const filteredcomp = getProperty(card.body, "video");
        if (filteredcomp.length > 0) {
            filteredcomp[0].url = imageLink;
        }
};



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
    // Check if imageLink is valid
    const isValidImageLink = imageLink && imageLink.trim() !== '';

    // Find the index of any existing logo component
    const logoIndex = card.body.findIndex((item: any) =>
        item.name === "logo" ||
        (item.type === "Image" && item.name === "logo"));

    // Case 1: Component exists
    if (logoIndex !== -1) {
        // If imageLink is empty/null, remove the component
        if (!isValidImageLink) {
            card.body.splice(logoIndex, 1);
            console.log("Logo component removed from card");
        } else {
            // Create completely new logo component
            const logoComponent = {
                "type": "Image",
                "url": imageLink,
                "altText": "Logo",
                "horizontalAlignment": "Center",
                "name": "logo",
                "spacing":"Small",
                "separator": true
            };

            // Replace existing component with the new one
            card.body[logoIndex] = logoComponent;
            console.log("Logo component replaced in card");
        }
    }
    // Case 2: Component doesn't exist but we have a valid link - add component
    else if (isValidImageLink) {
        // Create new logo component
        const logoComponent = {
            "type": "Image",
            "url": imageLink,
            "altText": "Logo",
            "horizontalAlignment": "Center",
            "spacing": "Small",
            "name": "logo",
            "separator": true
        };

        // Add to the beginning of the body array
        card.body.unshift(logoComponent);
        console.log("Logo component added to card");
    }
};

export const setCardBanner = (card: any, imageLink?: string) => {
    // Check if imageLink is valid
    const isValidImageLink = imageLink && imageLink.trim() !== '';

    // Find the index of any existing banner component
    const bannerIndex = card.body.findIndex((item: any) =>
        item.name === "banner" ||
        (item.type === "Image" && item.name === "banner"));

    // Case 1: Component exists
    if (bannerIndex !== -1) {
        // If imageLink is empty/null, remove the component
        if (!isValidImageLink) {
            card.body.splice(bannerIndex, 1);
            console.log("Banner component removed from card");
        } else {
            // Create completely new banner component
            const bannerComponent = {
                "type": "Image",
                "url": imageLink,
                "altText": "Banner",
                "horizontalAlignment": "Center",
                "name": "banner",
                "spacing": "Small"
            };

            // Replace existing component with the new one
            card.body[bannerIndex] = bannerComponent;
            console.log("Banner component replaced in card");
        }
    }
    // Case 2: Component doesn't exist but we have a valid link - add component
    else if (isValidImageLink) {
        // Create new banner component
        const bannerComponent = {
            "type": "Image",
            "url": imageLink,
            "altText": "Banner",
            "horizontalAlignment": "Center",
            "name": "banner",
            "spacing": "Small"
        };

        // Add to the end of the body array
        card.body.push(bannerComponent);
        console.log("Banner component added to card");
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

var getProperty = (body: any, value: string): any => {
    // Perform quick check for non-array bodies
    if (!body || !Array.isArray(body)) return [];

    // First try simple filter for better performance
    const simpleResult = body.filter((prop: any) => prop.name === value);

    // If we found results with the simple filter, return them
    if (simpleResult && simpleResult.length > 0) {
        return simpleResult;
    }

    // Otherwise, do a recursive search
    const results: any[] = [];

    const searchRecursively = (item: any): boolean => {
        // Skip if not an object or array, or if null/undefined
        if (!item || typeof item !== 'object') return false;

        // Check if this item has the name property we're looking for
        if (item.name === value) {
            results.push(item);
            return true; // Found what we're looking for
        }

        // If this is an array, search each element
        if (Array.isArray(item)) {
            for (const element of item) {
                // Exit early if found in this branch
                if (searchRecursively(element)) {
                    return true;
                }
            }
        } else {
            // Search object properties
            for (const prop of Object.values(item)) {
                if (prop && typeof prop === 'object') {
                    // Exit early if found in this branch
                    if (searchRecursively(prop)) {
                        return true;
                    }
                }
            }
        }

        return false; // Not found in this branch
    };

    // Try to find the property in each top-level item
    for (const item of body) {
        // If found in this branch, no need to check the rest
        if (searchRecursively(item)) {
            break;
        }
    }

    return results;


}