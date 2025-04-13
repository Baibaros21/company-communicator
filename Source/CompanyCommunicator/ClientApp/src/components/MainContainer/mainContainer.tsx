// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import "./mainContainer.scss";
import * as React from "react";
import { useTranslation } from "react-i18next";
import {
    Accordion,
    AccordionHeader,
    AccordionItem,
    AccordionPanel,
    Button,
    Divider,
    Link,
    teamsLightTheme,
    Theme,
    Menu,
    MenuPopover,
    MenuTrigger,
    MenuList,
    MenuItem
} from "@fluentui/react-components";
import { Status24Regular, PersonFeedback24Regular, QuestionCircle24Regular, Settings20Regular, Edit20Regular, BookTemplate20Regular } from "@fluentui/react-icons";
import * as microsoftTeams from "@microsoft/teams-js";
import { GetDraftMessagesSilentAction, GetAllCardTemplatesAction } from "../../actions";
import { RootState, useAppDispatch, useAppSelector } from "../../store";
import { getBaseUrl } from "../../configVariables";
import { ROUTE_PARTS, ROUTE_QUERY_PARAMS } from "../../routes";
import { DraftMessages } from "../DraftMessages/draftMessages";
import { SentMessages } from "../SentMessages/sentMessages";
import { getDefaultData } from '../../apis/messageListApi';



interface IMainContainer {
    theme: Theme;
    themeName: string;
}



export const MainContainer = (props: IMainContainer) => {
    const url = getBaseUrl() + `/${ROUTE_PARTS.NEW_MESSAGE}?${ROUTE_QUERY_PARAMS.LOCALE}={locale}`;
    const { t } = useTranslation();
    const dispatch = useAppDispatch();
    const loader = useAppSelector((state: RootState) => state.messages).isCardTemplatesFetchOn.payload;
    const Templates: any = useAppSelector((state: RootState) => state.messages).cardTemplates.payload;
    const [customHeaderImagePath, setCustomeHeaderImagepath] = React.useState<any>(null);
    const [customHeaderText, setCustomHeaderText] = React.useState<string>("");


    React.useEffect(() => {
        GetAllCardTemplatesAction(dispatch);
        getDefaultsItem();

    }, []);

    const getDefaultsItem = async () => {

        try {
            await getDefaultData().then((response) => {

                const defaultImages = response.data;
                setCustomeHeaderImagepath(defaultImages.headerLogoLink);
                setCustomHeaderText(defaultImages.headerText);

            });

        }
        catch (error) {

        }
    }

    const onOpenTaskModule = (event: any, url: string, title: string) => {
        let taskInfo: microsoftTeams.TaskInfo = {
            url: url,
            title: title,
            height: microsoftTeams.TaskModuleDimension.Large,
            width: microsoftTeams.TaskModuleDimension.Large,
            fallbackUrl: url,
        };
        let submitHandler = (err: any, result: any) => { };
        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    };

    const onNewMessage = () => {
        let taskInfo: microsoftTeams.TaskInfo = {
            url,
            title: t("NewMessage"),
            height: microsoftTeams.TaskModuleDimension.Large,
            width: microsoftTeams.TaskModuleDimension.Large,
            fallbackUrl: url,
        };

        let submitHandler = (err: any, result: any) => {
            if (result === null) {
                document.getElementById("newMessageButtonId")?.focus();
            } else {
                GetDraftMessagesSilentAction(dispatch);
            }
        };

        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    };
    const editDefaultsUrl = () =>
        getBaseUrl() + `/${ROUTE_PARTS.MODIFYDEFAULTS}`;
    const editTemplatesUrl = () =>
        getBaseUrl() + `/${ROUTE_PARTS.MODIFY_TEMPLATES}`;


    return (
        <>
            <div className={props.themeName === "default" ? "cc-header-light" : "cc-header"}>
                <div className="cc-main-left">
                    {(customHeaderImagePath!==null|| customHeaderImagePath !=="") && < img
                        src={customHeaderImagePath}
                        alt=""
                        className="cc-logo"
                        title={customHeaderText}
                    />}
                    <span className="cc-title" title={customHeaderText}>
                        { customHeaderText}
                    </span>
                </div>
                <div className="cc-main-right">
                    <span className="cc-icon-holder">
                        <Menu >
                            <MenuTrigger disableButtonEnhancement>
                                <Button className="cc-icon-button" aria-label='Actions menu' icon={<Settings20Regular className=" cc-icon-button-icon" />} />
                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList className="cc-menu-list">
                                    <MenuItem
                                        icon={<Edit20Regular />}
                                        key={'modifyDefaultsKey'}
                                        onClick={() => onOpenTaskModule(null, editDefaultsUrl(), t('modifyDefaults'))}
                                    >
                                        {t('modifyDefaults')}
                                    </MenuItem>
                                    <MenuItem
                                        icon={<BookTemplate20Regular />}
                                        key={'modifytemplatesKey'}
                                        onClick={() => onOpenTaskModule(null, editTemplatesUrl(), t('modifyTemplates'))}
                                    >
                                        {t('modifyTemplates')}
                                    </MenuItem>
                                </MenuList>
                            </MenuPopover>
                        </Menu>
                    </span>
                    <span className="cc-icon-holder">
                        <Link title={t("Support")} className="cc-icon-link" target="_blank" href="https://aka.ms/M365CCIssues">
                            <QuestionCircle24Regular className="cc-icon" />
                        </Link>
                    </span>
                    <span className="cc-icon-holder">
                        <Link title={t("Feedback")} className="cc-icon-link" target="_blank" href="https://aka.ms/M365CCFeedback">
                            <PersonFeedback24Regular className="cc-icon" />
                        </Link>
                    </span>
                </div>
            </div>
            <Divider />
            {Templates && Templates.length > 0 && !loader &&
                <>

                    <div className="cc-new-message">
                        <Button
                            id="newMessageButtonId"
                            icon={<Status24Regular />}
                            appearance="primary"
                            onClick={onNewMessage}
                        >
                            {t("NewMessage")}
                        </Button>
                    </div>
                    <Accordion defaultOpenItems={["1", "2"]} multiple collapsible>
                        <AccordionItem value="1" key="draftMessagesKey">
                            <AccordionHeader>{t("DraftMessagesSectionTitle")}</AccordionHeader>
                            <AccordionPanel className="cc-accordion-panel">
                                <DraftMessages />
                            </AccordionPanel>
                        </AccordionItem>
                        <AccordionItem value="2" key="sentMessagesKey">
                            <AccordionHeader>{t("SentMessagesSectionTitle")}</AccordionHeader>
                            <AccordionPanel className="cc-accordion-panel">
                                <SentMessages />
                            </AccordionPanel>
                        </AccordionItem>
                    </Accordion>
                </>}
        </>
    );
};
