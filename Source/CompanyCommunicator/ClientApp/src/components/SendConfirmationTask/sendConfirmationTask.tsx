// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as AdaptiveCards from 'adaptivecards';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { useParams } from 'react-router-dom';
import { Button, Field, Label, Persona, Spinner, Text } from '@fluentui/react-components';
import * as microsoftTeams from '@microsoft/teams-js';
import {
    getConsentSummaries, getDraftNotification, getDefaultData, sendDraftNotification
} from '../../apis/messageListApi';
import {
    setCardAuthor, setCardDeptTitle,
    setCardBtn, setCardImageLink, setCardSummary,
    setCardTitle, updateCardVideoElement, 
    setCardLogo, setCardBanner, setCardImageEmbedLink
} from '../AdaptiveCard/adaptiveCard';
import { AvatarShape } from '@fluentui/react-avatar';
import { RootState, useAppDispatch, useAppSelector, TemplateSelection } from '../../store';
import * as ACData from 'adaptivecards-templating';
import { GetAllCardTemplatesAction } from "../../actions";

export interface IMessageState {
    id: string;
    title: string;
    acknowledgements?: number;
    reactions?: number;
    responses?: number;
    succeeded?: number;
    failed?: number;
    throttled?: number;
    sentDate?: string;
    imageLink?: string;
    imageEmbedLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    createdBy?: string;
    isDraftMsgUpdated: boolean;
    department?: string;
    posterLink?: string;
    videoLink?: string;
    template: TemplateSelection;
    card: string,
}

interface ITemplates {

    name: string;
    card: any;
}
interface IDefaults {

    logoFileName: string;
    logoLink: string;
    bannerFileName: string;
    bannerLink: string;
}

export interface IConsentState {
    teamNames: string[];
    rosterNames: string[];
    groupNames: string[];
    allUsers: boolean;
    messageId: number;
    isConsentsUpdated: boolean;
}

let card: any;

export const SendConfirmationTask = () => {
    const { t } = useTranslation();
    const { id } = useParams() as any;
    const dispatch = useAppDispatch();
    const Templates: any = useAppSelector((state: RootState) => state.messages).cardTemplates.payload;
    const [loader, setLoader] = React.useState(true);
    const [isCardReady, setIsCardReady] = React.useState(false);
    const [isTemplateReady, setIsTemplateReady] = React.useState(false);
    const [isDefaultsReady, setIsDefaultsReady] = React.useState(false);
    const [disableSendButton, setDisableSendButton] = React.useState(false);
    const [cardAreaBorderClass, setCardAreaBorderClass] = React.useState('');

    const [messageState, setMessageState] = React.useState<IMessageState>({
        id: "",
        title: "",
        isDraftMsgUpdated: false,
        template: TemplateSelection.Default,
        card: "",
    });
    const [defaultsState, setDefaultState] = React.useState<IDefaults>({
        logoFileName: "",
        logoLink: "",
        bannerLink: "",
        bannerFileName: ""
    });
    const [consentState, setConsentState] = React.useState<IConsentState>({
        teamNames: [],
        rosterNames: [],
        groupNames: [],
        allUsers: false,
        messageId: 0,
        isConsentsUpdated: false,
    });
    React.useEffect(() => {
        GetAllCardTemplatesAction(dispatch);
    }, []);

    React.useEffect(() => {
        console.log("template checking");
        if (Templates && Templates.length > 0) {
            console.log("template ready");
            setIsTemplateReady(true);
        }

    }, [Templates])

    React.useEffect(() => {
        console.log("start");
        if (id && isTemplateReady) {
            getDefaults();
            getDraftMessage(id);
            getConsents(id);
        }
    }, [id, isTemplateReady]);


    React.useEffect(() => {
        if (isCardReady && consentState.isConsentsUpdated && messageState.isDraftMsgUpdated && isDefaultsReady && isTemplateReady) {
            var adaptiveCard = new AdaptiveCards.AdaptiveCard();
            setCardLogo(card, defaultsState.logoLink);
            setCardBanner(card, defaultsState.bannerLink);
            const cardJsonString: string = JSON.stringify(card);
            setMessageState({ ...messageState, card: cardJsonString, });
            adaptiveCard.parse(card);
            const renderCard = adaptiveCard.render();
            if (renderCard) {
                document.getElementsByClassName("card-area-1")[0].appendChild(renderCard);
                setCardAreaBorderClass('card-area-border');
            }
            adaptiveCard.onExecuteAction = function (action: any) {
                window.open(action.url, "_blank");
            };

            setLoader(false);
        }
    }, [isCardReady, consentState.isConsentsUpdated, messageState.isDraftMsgUpdated, isDefaultsReady, isTemplateReady]);

    const getCurrentCardTemplate = (cardtemplate: TemplateSelection) => {
        var currentTemplate = Templates?.find((template: ITemplates) => template.name === cardtemplate)?.card;


        var cardTemplate = new ACData.Template(JSON.parse(currentTemplate));
        card = cardTemplate.expand({
            $root: {


            }
        });
        setCardLogo(card, defaultsState.logoLink);
        setCardBanner(card, defaultsState.bannerLink);


    };


    const updateCardData = (msg: IMessageState) => {
        getCurrentCardTemplate(msg.template);

        setCardTitle(card, msg.title);
        setCardImageLink(card, msg.imageLink);
        setCardImageEmbedLink(card, msg.imageEmbedLink);
        setCardSummary(card, msg.summary);
        setCardAuthor(card, msg.author);
        setCardDeptTitle(card, msg.department);
        updateCardVideoElement(card, msg.videoLink, msg.posterLink);
        if (msg.buttonTitle && msg.buttonLink) {
            setCardBtn(card, msg.buttonTitle, msg.buttonLink);
        }
        setIsCardReady(true);
    };

    const getDefaults = async () => {
        await getDefaultData().then((response) => {
            const defaultImages = response.data;
            setDefaultState({
                logoFileName: defaultImages.logoFileName,
                logoLink: defaultImages.logoLink,
                bannerFileName: defaultImages.bannerFileName,
                bannerLink: defaultImages.bannerLink,
            });

            setIsDefaultsReady(true);
        });

    };

    const getDraftMessage = async (id: number) => {

        console.log("Getting message : ");
        try {
            await getDraftNotification(id).then((response) => {
                console.log("Response : ", response.data);
                setMessageState({
                    ...response.data,
                    isDraftMsgUpdated: true,
                });
                updateCardData(response.data);



            });
        } catch (error) {
            console.log("error Getting message : ");
            return error;
        }


    };

    const getConsents = async (id: number) => {
        try {
            await getConsentSummaries(id).then((response) => {
                setConsentState({
                    ...consentState,
                    teamNames: response.data.teamNames.sort(),
                    rosterNames: response.data.rosterNames.sort(),
                    groupNames: response.data.groupNames.sort(),
                    allUsers: response.data.allUsers,
                    messageId: id,
                    isConsentsUpdated: true,
                });
            });
        } catch (error) {
            return error;
        }
    };

    const onSendMessage = () => {

        console.log("Message card : ", messageState.card);
        console.log("Message card State: ", messageState);
        console.log("Defaults state: ", defaultsState);
        setDisableSendButton(true);
        sendDraftNotification(messageState)
            .then(() => {
                microsoftTeams.tasks.submitTask();
            })
            .finally(() => {
                setDisableSendButton(false);
            });
    };

    const getItemList = (items: string[], secondaryText: string, shape: AvatarShape) => {
        let resultedTeams: any[] = [];
        if (items) {
            items.map((element) => {
                resultedTeams.push(
                    <li key={element + "key"}>
                        <Persona name={element} secondaryText={secondaryText} avatar={{ shape, color: "colorful" }} />
                    </li>
                );
            });
        }
        return resultedTeams;
    };

    const renderAudienceSelection = () => {
        if (consentState.teamNames && consentState.teamNames.length > 0) {
            return (
                <div key="teamNames" style={{ paddingBottom: "16px" }}>
                    <Label>{t("TeamsLabel")}</Label>
                    <ul className="ul-no-bullets">{getItemList(consentState.teamNames, "Team", "square")}</ul>
                </div>
            );
        } else if (consentState.rosterNames && consentState.rosterNames.length > 0) {
            return (
                <div key="rosterNames" style={{ paddingBottom: "16px" }}>
                    <Label>{t("TeamsMembersLabel")}</Label>
                    <ul className="ul-no-bullets">{getItemList(consentState.rosterNames, "Team", "square")}</ul>
                </div>
            );
        } else if (consentState.groupNames && consentState.groupNames.length > 0) {
            return (
                <div key="groupNames" style={{ paddingBottom: "16px" }}>
                    <Label>{t("GroupsMembersLabel")}</Label>
                    <ul className="ul-no-bullets">{getItemList(consentState.groupNames, "Group", "circular")}</ul>
                </div>
            );
        } else if (consentState.allUsers) {
            return (
                <div key="allUsers" style={{ paddingBottom: "16px" }}>
                    <Label>{t("AllUsersLabel")}</Label>
                    <div>
                        <Text className="info-text">{t("SendToAllUsersNote")}</Text>
                    </div>
                </div>
            );
        } else {
            return <div></div>;
        }
    };

    return (
        <>
            {loader && <Spinner />}
            <>
                <div className='adaptive-task-grid'>
                    <div className='form-area'>
                        {!loader && (
                            <>
                                <div style={{ paddingBottom: '16px' }}>
                                    <Field size='large' label={t('ConfirmToSend')}>
                                        <Text>{t('SendToRecipientsLabel')}</Text>
                                    </Field>
                                </div>
                                <div>{renderAudienceSelection()}</div>
                            </>
                        )}
                    </div>
                    <div className='card-area'>
                        <div className={cardAreaBorderClass}>
                            <div className='card-area-1'></div>
                        </div>
                    </div>
                </div>
                <div className='fixed-footer'>
                    <div className='footer-action-right'>
                        <div className='footer-actions-flex'>
                            {disableSendButton && <Spinner role='alert' id='sendLoader' label={t('PreparingMessageLabel')} size='small' labelPosition='after' />}
                            <Button disabled={loader || disableSendButton} style={{ marginLeft: '16px' }} onClick={onSendMessage} appearance='primary'>
                                {t('Send')}
                            </Button>
                        </div>
                    </div>
                </div>
            </>
        </>
    );
};
