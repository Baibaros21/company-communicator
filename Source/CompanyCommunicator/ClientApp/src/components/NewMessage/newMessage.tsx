// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import './newMessages.scss';
import {
    ChevronDownRegular
 } from '@fluentui/react-icons';
import * as AdaptiveCards from 'adaptivecards';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { useParams } from 'react-router-dom';
import validator from 'validator';
import {
    Button,
    Combobox,
    ComboboxProps,
    Field,
    Input,
    Label,
    LabelProps,
    makeStyles,
    Option,
    Persona,
    Radio,
    RadioGroup,
    RadioGroupOnChangeData,
    shorthands,
    Spinner,
    Text,
    Textarea,
    tokens,
    useId,
} from '@fluentui/react-components';
import { InfoLabel } from '@fluentui/react-components/unstable';
import { ArrowUpload24Regular, Dismiss12Regular } from '@fluentui/react-icons';
import * as microsoftTeams from '@microsoft/teams-js';
import * as ACData from 'adaptivecards-templating';

import { GetDraftMessagesSilentAction, GetGroupsAction, GetTeamsDataAction, SearchGroupsAction, VerifyGroupAccessAction, GetAllCardTemplatesAction } from '../../actions';
import { createDraftNotification, getDraftNotification, updateDraftNotification, getDefaultData, getAppId } from '../../apis/messageListApi';
import { getBaseUrl } from '../../configVariables';
import { RootState, useAppDispatch, useAppSelector, TemplateSelection } from '../../store';
import {
    setCardAuthor, setCardDeptTitle,
    setCardBtn, setCardImageLink, setCardSummary,
    setCardTitle, setCardVideoPlayerUrl, setCardVideoPlayerPoster,
    setCardLogo, setCardBanner
} from '../AdaptiveCard/adaptiveCard';

const validImageTypes = ['image/gif', 'image/jpeg', 'image/png', 'image/jpg'];

interface IMessageState {
    id?: string;
    title: string;
    department?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonTitle?: string;
    buttonLink?: string;
    posterLink?: string;
    videoLink?: string;
    template: TemplateSelection;
    teams: any[];
    rosters: any[];
    groups: any[];
    allUsers: boolean;

}

interface ITeamTemplate {
    id: string;
    name: string;
}

interface IDefaults {

    logoFileName: string;
    logoLink: string;
    bannerFileName: string;
    bannerLink: string;
}

const useComboboxStyles = makeStyles({
    root: {
        display: 'grid',
        gridTemplateRows: 'repeat(1fr)',
        justifyItems: 'start',
        ...shorthands.gap('2px'),
        paddingLeft: '36px',
    },
    combobox: {
        position: 'relative',
        width: '100%',
        backgroundColor: tokens.colorNeutralBackground3,
    },
    comboboxInputContainer: {
        display: 'flex',
        alignItems: 'center',
        width: '100%',
        borderBlockColor: tokens.colorTransparentStroke,
        borderInlineColor: tokens.colorTransparentStroke,

        ...shorthands.borderRadius("4px"),
        ...shorthands.borderWidth('0px'),
        ...shorthands.borderStyle('solid'),
        ...shorthands.padding('1px', '4px'),
    },
    comboboxInput: {
        width: '100%',
        backgroundColor: tokens.colorNeutralBackground3,
        fontSize: tokens.fontSizeBase400,
        fontWeight: tokens.fontWeightRegular,
        color: tokens.colorNeutralForeground1,
        height: '38px',
        paddingLeft: tokens.spacingHorizontalM,

        borderBlockColor: tokens.colorTransparentStroke,
        borderInlineColor: tokens.colorTransparentStroke,

        ...shorthands.borderRadius("4px"),
        ...shorthands.borderWidth('0px'),
        ...shorthands.borderStyle('solid'),
        ...shorthands.padding('1px', '2px'),
    },
    comboboxIcon: {
        ...shorthands.padding('10px'),
        cursor: 'pointer',
    },
    comboboxOptions: {
        listStyleType: 'none',
        backgroundColor: tokens.colorTransparentBackground,

        position: 'absolute',
        width: '100%',
        ...shorthands.borderRadius(tokens.borderRadiusMedium),

        maxHeight: '200px',
        overflowY: 'auto',
        zIndex: 1000,
        display: 'none', // Initially hidden
    },
    comboboxOption: {

        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        backgroundColor: tokens.colorNeutralBackground1,
    },
    comboboxOptionHover: {
        backgroundColor: tokens.colorNeutralBackground1Hover,
    },
    tagsList: {
        listStyleType: 'none',
        marginBottom: tokens.spacingVerticalXXS,
        marginTop: 0,
        paddingLeft: 0,
        // display: "flex",
        gridGap: tokens.spacingHorizontalXXS,
    },
});



const useFieldStyles = makeStyles({
    styles: {
        marginBottom: tokens.spacingVerticalM,
        gridGap: tokens.spacingHorizontalXXS,
    },
});

interface ITemplates {

    name: string;
    card: any;
}
enum AudienceSelection {
    Teams = 'Teams',
    Rosters = 'Rosters',
    Groups = 'Groups',
    AllUsers = 'AllUsers',
    None = 'None',
}

enum CurrentPageSelection {
    TemplateCreation = "TemplateCreation",
    CardCreation = 'CardCreation',
    AudienceSelection = 'AudienceSelection',
}

let card: any;

const MAX_SELECTED_TEAMS_NUM: number = 20;

export const NewMessage = () => {
    let fileInput = React.createRef<any>();
    let posterFileInput = React.createRef<any>();
    const { t } = useTranslation();
    const { id } = useParams() as any;
    const dispatch = useAppDispatch();
    const Templates: any = useAppSelector((state: RootState) => state.messages).cardTemplates.payload;
    const teams = useAppSelector((state: RootState) => state.messages).teamsData.payload;
    const groups = useAppSelector((state: RootState) => state.messages).groups.payload;
    const queryGroups = useAppSelector((state: RootState) => state.messages).queryGroups.payload;
    const canAccessGroups = useAppSelector((state: RootState) => state.messages).verifyGroup.payload;
    const [selectedRadioButton, setSelectedRadioButton] = React.useState(AudienceSelection.None);
    const [selectedTemplate, setSelectedTemplate] = React.useState(TemplateSelection.Default);
    const [pageSelection, setPageSelection] = React.useState(CurrentPageSelection.TemplateCreation);
    const [allUsersState, setAllUsersState] = React.useState(false);
    const [imageFileName, setImageFileName] = React.useState('');
    const [videoFileName, setVideoFileName] = React.useState('');
    const [InternalAppId, setInternalAppId] = React.useState('');
    const [posterFileName, setPosterFileName] = React.useState('');
    const [imageUploadErrorMessage, setImageUploadErrorMessage] = React.useState('');
    const [titleErrorMessage, setTitleErrorMessage] = React.useState('');
    const [btnLinkErrorMessage, setBtnLinkErrorMessage] = React.useState('');
    const [showMsgDraftingSpinner, setShowMsgDraftingSpinner] = React.useState(false);
    const [isCardReady, setIsCardReady] = React.useState(false);

    const [allUsersAria, setAllUserAria] = React.useState('none');
    const [groupsAria, setGroupsAria] = React.useState('none');
    const [cardAreaBorderClass, setCardAreaBorderClass] = React.useState('');
    const [defaultsState, setDefaultState] = React.useState<IDefaults>({
        logoFileName: "",
        logoLink: "",
        bannerLink: "",
        bannerFileName: ""
    });


    const [messageState, setMessageState] = React.useState<IMessageState>({
        title: '',
        template: TemplateSelection.Default,
        teams: [],
        rosters: [],
        groups: [],
        allUsers: false,
    });

    // Handle selectedOptions both when an option is selected or deselected in the Combobox,
    // and when an option is removed by clicking on a tag
    const [teamsSelectedOptions, setTeamsSelectedOptions] = React.useState<ITeamTemplate[]>([]);
    const [rostersSelectedOptions, setRostersSelectedOptions] = React.useState<ITeamTemplate[]>([]);
    const [searchSelectedOptions, setSearchSelectedOptions] = React.useState<ITeamTemplate[]>([]);
    //const [cardTemplates, setCardTemplates] = React.useState<[ITemplates]>();

    React.useEffect(() => {
        GetTeamsDataAction(dispatch);
        VerifyGroupAccessAction(dispatch);
        GetAllCardTemplatesAction(dispatch);
        getDefaultsItem();
    }, []);

    React.useEffect(() => {
        if (Templates && Templates.length > 0) {
            getCurrentCardTemplate(selectedTemplate);
/*            setDefaultCard(card);
*/            updateAdaptiveCard();
            setIsCardReady(true);
        }

    }, [Templates])


    React.useEffect(() => {
        if (isCardReady) {
            updateAdaptiveCard();
        }
    }, [pageSelection]);
    React.useEffect(
        () => {
            if (isCardReady) {
                if (messageState.title !== "") setCardTitle(card, messageState.title);
                if (messageState.imageLink !== "") setCardImageLink(card, messageState.imageLink);
                if (messageState.department !== "") setCardDeptTitle(card, messageState.department);
                if (messageState.posterLink !== "") setCardVideoPlayerPoster(card, messageState.posterLink);
                if (messageState.videoLink !== "") setCardVideoPlayerUrl(card, messageState.videoLink);
                if (messageState.department !== "") setCardDeptTitle(card, messageState.department);
                if (messageState.summary !== "") setCardSummary(card, messageState.summary);
                if (messageState.author !== "") setCardAuthor(card, messageState.author);
                if (defaultsState.logoLink!=="") setCardLogo(card, defaultsState.logoLink);
                if (defaultsState.bannerLink!=="") setCardBanner(card, defaultsState.bannerLink);
                if (messageState.buttonTitle !== "") setCardBtn(card, messageState.buttonTitle, messageState.buttonLink);
                if (!messageState.title && !messageState.imageLink && !messageState.summary && !messageState.author && !messageState.buttonTitle && !messageState.buttonLink) {
                    getCurrentCardTemplate(selectedTemplate);
/*                    setDefaultCard(card);
*/                }
                updateAdaptiveCard();
            }
        }
        ,
        [messageState, isCardReady]);

    React.useEffect(() => {
        if (id) {
            GetGroupsAction(dispatch, { id });
            getDraftNotificationItem(id);
        }
    }, [id]);


    React.useEffect(() => {
        setTeamsSelectedOptions([]);
        setRostersSelectedOptions([]);
        setSearchSelectedOptions([]);
        setAllUsersState(false);
        if (teams && teams.length > 0) {
            const teamsSelected = teams.filter((c) => messageState.teams.some((s) => s === c.id));
            setTeamsSelectedOptions(teamsSelected || []);
            const roastersSelected = teams.filter((c) => messageState.rosters.some((s) => s === c.id));
            setRostersSelectedOptions(roastersSelected || []);
        }
        if (groups && groups.length > 0) {
            const groupsSelected = groups.filter((c) => messageState.groups.some((s) => s === c.id));
            setSearchSelectedOptions(groupsSelected || []);
        }
        if (messageState.allUsers) {
            setAllUsersState(true);
        }
    }, [teams, groups, messageState.teams, messageState.rosters, messageState.allUsers, messageState.groups]);

    const getCurrentCardTemplate = (cardtemplate: TemplateSelection) => {
        console.log(Templates);
        var currentTemplate = Templates?.find((template: ITemplates) => template.name === cardtemplate)?.card;

        console.log(card);
        var cardTemplate = new ACData.Template(JSON.parse(currentTemplate));
        card = cardTemplate.expand({
            $root: {


            }
        });
        setCardLogo(card, defaultsState.logoLink);
        setCardBanner(card, defaultsState.bannerLink);


    }



    const getDefaultsItem = async () => {

        try {
            await getDefaultData().then((response) => {

                const defaultImages = response.data;
                setDefaultState({
                    logoFileName: defaultImages.logoFileName,
                    logoLink: defaultImages.logoLink,
                    bannerFileName: defaultImages.bannerFileName,
                    bannerLink: defaultImages.bannerLink
                });

            });

            await getAppId().then((reponse) => {
                setInternalAppId(reponse.data);
            });

        }
        catch (error) {
            card = getCurrentCardTemplate(TemplateSelection.Default);
            updateAdaptiveCard();
        }
    }




    const getDraftNotificationItem = async (id: number) => {
        try {
            await getDraftNotification(id).then((response) => {
                const draftMessageDetail = response.data;

                if (draftMessageDetail.teams.length > 0) {
                    setSelectedRadioButton(AudienceSelection.Teams);
                } else if (draftMessageDetail.rosters.length > 0) {
                    setSelectedRadioButton(AudienceSelection.Rosters);
                } else if (draftMessageDetail.groups.length > 0) {
                    setSelectedRadioButton(AudienceSelection.Groups);
                } else if (draftMessageDetail.allUsers) {
                    setSelectedRadioButton(AudienceSelection.AllUsers);
                }
                setMessageState({
                    ...messageState,
                    id: draftMessageDetail.id,
                    title: draftMessageDetail.title,
                    department: draftMessageDetail.department,
                    imageLink: draftMessageDetail.imageLink,
                    posterLink: draftMessageDetail.posterLink,
                    videoLink: draftMessageDetail.videoLink,
                    summary: draftMessageDetail.summary,
                    author: draftMessageDetail.author,
                    buttonTitle: draftMessageDetail.buttonTitle,
                    buttonLink: draftMessageDetail.buttonLink,
                    teams: draftMessageDetail.teams,
                    rosters: draftMessageDetail.rosters,
                    groups: draftMessageDetail.groups,
                    allUsers: draftMessageDetail.allUsers,
                    template: draftMessageDetail.template
                });
                setSelectedTemplate(draftMessageDetail.template);
                setCardTitle(card, draftMessageDetail.title);
                setCardDeptTitle(card, draftMessageDetail.department);
                setCardImageLink(card, draftMessageDetail.imageLink);
                setCardVideoPlayerPoster(card, draftMessageDetail.posterLink);
                setCardVideoPlayerUrl(card, draftMessageDetail.videoLink);
                setCardDeptTitle(card, draftMessageDetail.department);
                setCardSummary(card, draftMessageDetail.summary);
                setCardAuthor(card, draftMessageDetail.author);
                setCardBtn(card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink);
                setCardLogo(card, defaultsState.logoLink);
                setCardBanner(card, defaultsState.bannerLink);
                setTeamsSelectedOptions(draftMessageDetail.template)


                updateAdaptiveCard();
            });
        } catch (error) {
            return error;
        }
    };


    const templateSelectionChange = (ev: any, data: RadioGroupOnChangeData) => {
        let input = data.value as keyof typeof TemplateSelection;
        setSelectedTemplate(TemplateSelection[input]);
        getCurrentCardTemplate(TemplateSelection[input]);
/*        setDefaultCard(card);
*/        updateAdaptiveCard();
    };



    const updateAdaptiveCard = () => {
        var adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        const renderCard = adaptiveCard.render();

        if (renderCard && pageSelection === CurrentPageSelection.CardCreation) {
            document.getElementsByClassName('card-area-1')[0].innerHTML = '';
            document.getElementsByClassName('card-area-1')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        } else if (renderCard && pageSelection === CurrentPageSelection.AudienceSelection) {
            document.getElementsByClassName('card-area-2')[0].innerHTML = '';
            document.getElementsByClassName('card-area-2')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        } else if (renderCard && pageSelection === CurrentPageSelection.TemplateCreation) {
            document.getElementsByClassName('card-area-3')[0].innerHTML = '';
            document.getElementsByClassName('card-area-3')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        }
        adaptiveCard.onExecuteAction = function (action: any) {
            window.open(action.url, '_blank');
        };
    };



    const checkValidSizeOfImage = (resizedImageAsBase64: string) => {
        var stringLength = resizedImageAsBase64.length - 'data:image/png;base64,'.length;
        var sizeInBytes = 4 * Math.ceil(stringLength / 3) * 0.5624896334383812;
        var sizeInKb = sizeInBytes / 1000;

        if (sizeInKb <= 1024) return true;
        else return false;
    };


    const handleUploadClick = (event: any) => {
        if (fileInput.current) {
            fileInput.current.click();
        }
    };
    const handlePosterUploadClick = (event: any) => {
        if (posterFileInput.current) {
            posterFileInput.current.click();
        }
    }
    const handlePosterSelection = () => {
        const file = posterFileInput.current?.files[0];
        imageselection(file, "poster");


    }
    const handleImageSelection = () => {
        const file = fileInput.current?.files[0];

        imageselection(file, "image");



    };

    const imageselection = (file: any, field: string): any => {

        if (file) {
            const fileType = file['type'];
            const { type: mimeType } = file;

            if (!validImageTypes.includes(fileType)) {
                setImageUploadErrorMessage(t('ErrorImageTypesMessage'));
                return;
            }



            const fileReader = new FileReader();
            fileReader.readAsDataURL(file);
            fileReader.onload = () => {

                var image = new Image();
                image.src = fileReader.result as string;
                var resizedImageAsBase64 = fileReader.result as string;


                image.onload = function (e: any) {
                    const MAX_WIDTH = 1024;

                    if (image.width > MAX_WIDTH) {
                        const canvas = document.createElement('canvas');
                        canvas.width = MAX_WIDTH;
                        canvas.height = ~~(image.height * (MAX_WIDTH / image.width));
                        const context = canvas.getContext('2d', { alpha: false });
                        if (!context) {
                            return;
                        }
                        context.drawImage(image, 0, 0, canvas.width, canvas.height);
                        resizedImageAsBase64 = canvas.toDataURL(mimeType);
                    }
                };

                if (!checkValidSizeOfImage(resizedImageAsBase64)) {
                    setImageUploadErrorMessage(t('ErrorImageSizeMessage'));
                    return;
                }

                if (resizedImageAsBase64 && field === 'image') {
                    setImageFileName(file['name']);
                    setImageUploadErrorMessage('');
                    setCardImageLink(card, resizedImageAsBase64);
                    setMessageState({ ...messageState, imageLink: resizedImageAsBase64 });
                } else if (resizedImageAsBase64 && field === 'poster') {
                    setPosterFileName(file['name']);
                    setImageUploadErrorMessage('');
                    setCardVideoPlayerPoster(card, resizedImageAsBase64);
                    setMessageState({
                        ...messageState, posterLink: resizedImageAsBase64
                    });
                }


                updateAdaptiveCard();
            };
        }
    }

    const isSaveBtnDisabled = () => {
        const msg_page_conditions = messageState.title !== '' && imageUploadErrorMessage === '' && btnLinkErrorMessage === '';
        const aud_page_conditions =
            (teamsSelectedOptions.length > 0 && selectedRadioButton === AudienceSelection.Teams) ||
            (rostersSelectedOptions.length > 0 && selectedRadioButton === AudienceSelection.Rosters) ||
            (searchSelectedOptions.length > 0 && selectedRadioButton === AudienceSelection.Groups) ||
            selectedRadioButton === AudienceSelection.AllUsers;

        if (msg_page_conditions && aud_page_conditions) {
            return false;
        } else {
            return true;
        }
    };

    const isNextBtnDisabled = () => {
        if (messageState.title !== '' && imageUploadErrorMessage === '' && btnLinkErrorMessage === '') {
            return false;
        } else {
            return true;
        }
    };

    const onSave = () => {
        let finalSelectedTeams: string[] = [];
        let finalSelectedRosters: string[] = [];
        let finalSelectedGroups: string[] = [];
        let finalAllUsers: boolean = false;

        if (selectedRadioButton === AudienceSelection.Teams) {
            finalSelectedTeams = [...teams.filter((t1) => teamsSelectedOptions.some((sp) => sp.id === t1.id)).map((t2) => t2.id)];
        }
        if (selectedRadioButton === AudienceSelection.Rosters) {
            finalSelectedRosters = [...teams.filter((t1) => rostersSelectedOptions.some((sp) => sp.id === t1.id)).map((t2) => t2.id)];
        }
        if (selectedRadioButton === AudienceSelection.Groups) {
            finalSelectedGroups = [...searchSelectedOptions.map((g) => g.id)];
        }
        if (selectedRadioButton === AudienceSelection.AllUsers) {
            finalAllUsers = allUsersState;
        }

        const finalMessage = {
            ...messageState,
            template: selectedTemplate,
            teams: finalSelectedTeams,
            rosters: finalSelectedRosters,
            groups: finalSelectedGroups,
            allUsers: finalAllUsers,
        };

        setShowMsgDraftingSpinner(true);

        if (id) {
            editDraftMessage(finalMessage);
        } else {
            postDraftMessage(finalMessage);
        }
    };

    const editDraftMessage = (msg: IMessageState) => {
        try {
            updateDraftNotification(msg)
                .then(() => {
                    GetDraftMessagesSilentAction(dispatch);
                })
                .finally(() => {
                    setShowMsgDraftingSpinner(false);
                    microsoftTeams.tasks.submitTask();
                });
        } catch (error) {
            return error;
        }
    };

    const postDraftMessage = (msg: IMessageState) => {
        try {
            createDraftNotification(msg)
                .then(() => {
                    GetDraftMessagesSilentAction(dispatch);
                })
                .finally(() => {
                    setShowMsgDraftingSpinner(false);
                    microsoftTeams.tasks.submitTask();
                });
        } catch (error) {
            return error;
        }
    };

    const onNext = (event: any) => {
        switch (pageSelection) {
            case (CurrentPageSelection.TemplateCreation):
                setPageSelection(CurrentPageSelection.CardCreation);
                break;
            case (CurrentPageSelection.CardCreation):
                setPageSelection(CurrentPageSelection.AudienceSelection)
                break;
            default:


        }

    };

    const onBack = (event: any) => {

        switch (pageSelection) {
            case (CurrentPageSelection.CardCreation):
                setPageSelection(CurrentPageSelection.TemplateCreation);
                break;
            case (CurrentPageSelection.AudienceSelection):
                setPageSelection(CurrentPageSelection.CardCreation);
                setAllUserAria('none');
                setGroupsAria('none');
                break;
            default:


        }
    };

    const onTitleChanged = (event: any) => {
        if (event.target.value === '') {
            setTitleErrorMessage('Title is required.');
        } else {
            setTitleErrorMessage('');
        }
        setCardTitle(card, event.target.value);
        setMessageState({ ...messageState, title: event.target.value });
        updateAdaptiveCard();
    };

    const onDeptChanged = (event: any) => {
        setCardDeptTitle(card, event.target.value);
        setMessageState({ ...messageState, department: event.target.value });
        updateAdaptiveCard();
    };

    const onImageLinkChanged = (event: any) => {
        const urlOrDataUrl = event.target.value;
        let isGoodLink = true;
        setImageFileName(urlOrDataUrl);

        if (
            !(
                urlOrDataUrl === '' ||
                urlOrDataUrl.startsWith('https://') ||
                urlOrDataUrl.startsWith('data:image/png;base64,') ||
                urlOrDataUrl.startsWith('data:image/jpeg;base64,') ||
                urlOrDataUrl.startsWith('data:image/gif;base64,')
            )
        ) {
            isGoodLink = false;
            setImageUploadErrorMessage(t('ErrorURLMessage'));
        } else {
            isGoodLink = true;
            setImageUploadErrorMessage(t(''));
        }

        if (isGoodLink) {
            setMessageState({ ...messageState, imageLink: urlOrDataUrl });
            setCardImageLink(card, event.target.value);
            updateAdaptiveCard();
        }
    };


    const onPosterLinkChanged = (event: any) => {
        const urlOrDataUrl = event.target.value;
        let isGoodLink = true;
        setPosterFileName(urlOrDataUrl);

        if (
            !(
                urlOrDataUrl === '' ||
                urlOrDataUrl.startsWith('https://') ||
                urlOrDataUrl.startsWith('data:image/png;base64,') ||
                urlOrDataUrl.startsWith('data:image/jpeg;base64,') ||
                urlOrDataUrl.startsWith('data:image/gif;base64,')
            )
        ) {
            isGoodLink = false;
            setImageUploadErrorMessage(t('ErrorURLMessage'));
        } else {
            isGoodLink = true;
            setImageUploadErrorMessage(t(''));
        }

        if (isGoodLink) {
            setMessageState({ ...messageState, posterLink: urlOrDataUrl });
            setCardVideoPlayerPoster(card, event.target.value);
            updateAdaptiveCard();
        }
    };

    const onVideoLinkChanged = (event: any) => {
        const urlOrDataUrl = event.target.value;
        let isGoodLink = true;
        setVideoFileName(urlOrDataUrl);

        if (
            !(
                urlOrDataUrl === '' ||
                urlOrDataUrl.startsWith('https://')
            )
        ) {
            isGoodLink = false;
            setImageUploadErrorMessage(t('ErrorURLMessage'));
        } else {
            isGoodLink = true;
            setImageUploadErrorMessage(t(''));
        }


        var url = "https://teams.microsoft.com/l/task/" + InternalAppId
            + "?url=" + "https://companycommunicator.blueridgeit.com/videoplayer/" + urlOrDataUrl
            + "&height=large&width=large&title=ADVideo";

        if (isGoodLink) {
            setMessageState({ ...messageState, videoLink: url });
            setCardVideoPlayerUrl(card, url);
            updateAdaptiveCard();
        }

    };

    const onSummaryChanged = (event: any) => {
        setCardSummary(card, event.target.value);
        setMessageState({ ...messageState, summary: event.target.value });
        updateAdaptiveCard();
    };

    const onAuthorChanged = (event: any) => {
        setCardAuthor(card, event.target.value);
        setMessageState({ ...messageState, author: event.target.value });
        updateAdaptiveCard();
    };

    const onBtnTitleChanged = (event: any) => {
        setCardBtn(card, event.target.value, messageState.buttonLink);
        setMessageState({ ...messageState, buttonTitle: event.target.value });
        updateAdaptiveCard();
    };

    const onBtnLinkChanged = (event: any) => {
        if (validator.isURL(event.target.value, { require_protocol: true, protocols: ['https'] }) || event.target.value === '') {
            setBtnLinkErrorMessage('');
        } else {
            setBtnLinkErrorMessage(`${event.target.value} is invalid. Please enter a valid https URL`);
        }
        setCardBtn(card, messageState.buttonTitle, event.target.value);
        setMessageState({ ...messageState, buttonLink: event.target.value });
        updateAdaptiveCard();
    };

    // generate ids for handling labelling
    const teamsComboId = useId('teams-combo-multi');
    const teamsSelectedListId = `${teamsComboId}-selection`;

    const rostersComboId = useId('rosters-combo-multi');
    const rostersSelectedListId = `${rostersComboId}-selection`;

    const searchComboId = useId('search-combo-multi');
    const searchSelectedListId = `${searchComboId}-selection`;

    // refs for managing focus when removing tags
    const teamsSelectedListRef = React.useRef<HTMLUListElement>(null);
    const teamsComboboxInputRef = React.useRef<HTMLInputElement>(null);
    const teamsComboboxOptionRef = React.useRef<HTMLUListElement>(null);


    const rostersSelectedListRef = React.useRef<HTMLUListElement>(null);
    const rostersComboboxInputRef = React.useRef<HTMLInputElement>(null);
    const rosterComboboxOptionRef = React.useRef<HTMLUListElement>(null);


    const searchSelectedListRef = React.useRef<HTMLUListElement>(null);
    const searchComboboxInputRef = React.useRef<HTMLInputElement>(null);
    const searchComboboxOptionRef = React.useRef<HTMLUListElement>(null);


    //Custom Combobox functions
    const handleInputClick = (optionList: React.MutableRefObject<HTMLUListElement | null>) => {
        if (optionList.current && optionList.current.style) {
            optionList.current.style.display = 'block';
        }
    };
    const handleClickOutside = (event: MouseEvent) => {

        if (teamsComboboxOptionRef.current && teamsComboboxOptionRef.current.style) {
            teamsComboboxOptionRef.current.style.display = 'none';
        }
        if (rosterComboboxOptionRef.current && rosterComboboxOptionRef.current.style) {
            rosterComboboxOptionRef.current.style.display = 'none';
        }
        if (searchComboboxOptionRef.current && searchComboboxOptionRef.current.style) {
            searchComboboxOptionRef.current.style.display = 'none';
        }
    };
    React.useEffect(() => {
        document.addEventListener('click', handleClickOutside);
        return () => {
            document.removeEventListener('click', handleClickOutside);
        };
    }, []);

    //functions for handling teams. groups and users
    const onTeamsSelect: ComboboxProps['onOptionSelect'] = (event, data) => {
        if (data.selectedOptions.length <= MAX_SELECTED_TEAMS_NUM) {
            setTeamsSelectedOptions(teams.filter((t1) => data.selectedOptions.some((t2) => t2 === t1.id)));
        }
    };

    const onTeamsSelectOpt = (opt: ITeamTemplate) => {
        if (teamsSelectedOptions.length < MAX_SELECTED_TEAMS_NUM) {
            const teamExists = teamsSelectedOptions.some(team => team.id === opt.id);
            if (!teamExists) {
                setTeamsSelectedOptions([...teamsSelectedOptions, opt]);
            }
        }
    };

    const onRostersSelect: ComboboxProps['onOptionSelect'] = (event, data) => {
        if (data.selectedOptions.length <= MAX_SELECTED_TEAMS_NUM) {
            setRostersSelectedOptions(teams.filter((t1) => data.selectedOptions.some((t2) => t2 === t1.id)));
        }
    };

    const onRosterSelectOpt = (opt: ITeamTemplate) => {
        if (rostersSelectedOptions.length < MAX_SELECTED_TEAMS_NUM) {
            const teamExists = rostersSelectedOptions.some(team => team.id === opt.id);
            if (!teamExists) {
                setRostersSelectedOptions([...rostersSelectedOptions, opt]);
            }
        }
    };

    const onSearchSelect: ComboboxProps['onOptionSelect'] = (event, data: any) => {
        if (data.optionText && !searchSelectedOptions.find((x) => x.id === data.optionValue)) {
            setSearchSelectedOptions([...searchSelectedOptions, { id: data.optionValue, name: data.optionText }]);
        }
    };

    const onSearchSelectOpt = (opt: ITeamTemplate) => {
        if (searchSelectedOptions.length < MAX_SELECTED_TEAMS_NUM) {
            const teamExists = searchSelectedOptions.some(team => team.id === opt.id);
            if (!teamExists) {
                setSearchSelectedOptions([...searchSelectedOptions, opt]);
            }
        }
    };

    const onSearchChange = (event: any) => {
        if (event && event.target && event.target.value) {
            const q = encodeURIComponent(event.target.value);
            SearchGroupsAction(dispatch, { query: q });
        }
    };

    const onTeamsTagClick = (option: ITeamTemplate, index: number) => {
        // remove selected option
        setTeamsSelectedOptions(teamsSelectedOptions.filter((o) => o.id !== option.id));

        // focus previous or next option, defaulting to focusing back to the combo input
        const indexToFocus = index === 0 ? 1 : index - 1;
        const optionToFocus = teamsSelectedListRef.current?.querySelector(`#${teamsComboId}-remove-${indexToFocus}`);
        if (optionToFocus) {
            (optionToFocus as HTMLButtonElement).focus();
        } else {
            teamsComboboxInputRef.current?.focus();
        }
    };

    const onRostersTagClick = (option: ITeamTemplate, index: number) => {
        // remove selected option
        setRostersSelectedOptions(rostersSelectedOptions.filter((o) => o.id !== option.id));

        // focus previous or next option, defaulting to focusing back to the combo input
        const indexToFocus = index === 0 ? 1 : index - 1;
        const optionToFocus = rostersSelectedListRef.current?.querySelector(`#${rostersComboId}-remove-${indexToFocus}`);
        if (optionToFocus) {
            (optionToFocus as HTMLButtonElement).focus();
        } else {
            rostersComboboxInputRef.current?.focus();
        }
    };

    const onSearchTagClick = (option: ITeamTemplate, index: number) => {
        // remove selected option
        setSearchSelectedOptions(searchSelectedOptions.filter((o) => o.id !== option.id));

        // focus previous or next option, defaulting to focusing back to the combo input
        const indexToFocus = index === 0 ? 1 : index - 1;
        const optionToFocus = searchSelectedListRef.current?.querySelector(`#${searchComboId}-remove-${indexToFocus}`);
        if (optionToFocus) {
            (optionToFocus as HTMLButtonElement).focus();
        } else {
            searchComboboxInputRef.current?.focus();
        }
    };

    const teamsLabelledBy = teamsSelectedOptions.length > 0 ? `${teamsComboId} ${teamsSelectedListId}` : teamsComboId;
    const rostersLabelledBy = rostersSelectedOptions.length > 0 ? `${rostersComboId} ${rostersSelectedListId}` : rostersComboId;

    const searchLabelledBy = searchSelectedOptions.length > 0 ? `${searchComboId} ${searchSelectedListId}` : searchComboId;

    const cmb_styles = useComboboxStyles();
    const field_styles = useFieldStyles();

    const audienceSelectionChange = (ev: any, data: RadioGroupOnChangeData) => {
        let input = data.value as keyof typeof AudienceSelection;
        setSelectedRadioButton(AudienceSelection[input]);

        if (AudienceSelection[input] === AudienceSelection.AllUsers) {
            setAllUsersState(true);
        } else if (allUsersState) {
            setAllUsersState(false);
        }

        AudienceSelection[input] === AudienceSelection.AllUsers ? setAllUserAria('alert') : setAllUserAria('none');
        AudienceSelection[input] === AudienceSelection.Groups ? setGroupsAria('alert') : setGroupsAria('none');
    };

    return (
        <>
            {pageSelection === CurrentPageSelection.TemplateCreation && Templates && Templates.length > 0 && (
                <>
                    <span role='alert' aria-label={t('NewMessageStep2')} />
                    <div className='adaptive-task-grid'>
                        <div className='form-area'>
                            <Label size='large' id='TemplateSelectionGroupLabelId'>
                                {t('SendHeadingText')}
                            </Label>
                            <RadioGroup defaultValue={selectedTemplate} aria-labelledby='TemplateSelectionGroupLabelId' onChange={templateSelectionChange}>
                                <Radio id='radio1' value="Default" label={TemplateSelection.Default} />

                                <Radio id='radio2' value="infromational" label={TemplateSelection.infromational} />

                                <Radio id='radio4' value="department" label={TemplateSelection.department} />

                                <Radio id='radio5' value="departmentVideo" label={TemplateSelection.departmentVideo} />

                                <Radio id='radio7' value="Default_ar" label={TemplateSelection.Default_ar} />

                                <Radio id='radio10' value="department_ar" label={TemplateSelection.department_ar} />

                                <Radio id='radio11' value="departmentVideo_ar" label={TemplateSelection.departmentVideo_ar} />



                            </RadioGroup>
                        </div>
                        <div className='card-area'>
                            <div className={cardAreaBorderClass}>
                                <div className='card-area-3'></div>
                            </div>
                        </div>
                    </div>
                    <div>
                        <div className='fixed-footer'>
                            <div className='footer-action-right'>
                                <div className='footer-actions-flex'>
                                    {showMsgDraftingSpinner && (
                                        <Spinner
                                            role='alert'
                                            id='draftingLoader'
                                            size='small'
                                            label={t('DraftingMessageLabel')}
                                            labelPosition='after'
                                        />
                                    )}
                                    <Button
                                        style={{ marginLeft: '16px' }}
                                        //disabled={isSaveBtnDisabled() || showMsgDraftingSpinner}
                                        id='saveBtn'
                                        onClick={onNext}
                                        appearance='primary'
                                    >
                                        {t('SetAsDraft')}
                                    </Button>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            )}
            {pageSelection === CurrentPageSelection.CardCreation && Templates && Templates.length > 0 && (
                <div className="page-container">
                    <span role='alert' aria-label={t('NewMessageStep1')} />
                    <div className='adaptive-task-grid'>
                        <div className='form-area'>
                            <Field size='large' className={field_styles.styles} label={t('TitleText')} required={true} validationMessage={titleErrorMessage}>
                                <Input
                                    placeholder={t('PlaceHolderTitle')}
                                    onChange={onTitleChanged}
                                    autoComplete='off'
                                    size='large'
                                    required={true}
                                    appearance='filled-darker'
                                    value={messageState.title || ''}
                                />
                            </Field>

                            {
                                (selectedTemplate === TemplateSelection.department
                                    || selectedTemplate === TemplateSelection.departmentVideo
                                    || selectedTemplate === TemplateSelection.department_ar
                                    || selectedTemplate === TemplateSelection.departmentVideo_ar)
                                && (<>



                                    <Field size='large' className={field_styles.styles} label={t('departmentText')} required={false} >
                                        <Input
                                            placeholder={t('PlaceHolderDepartment')}
                                            onChange={onDeptChanged}
                                            autoComplete='off'
                                            size='large'
                                            required={false}
                                            appearance='filled-darker'
                                            value={messageState.department || ''}
                                        />
                                    </Field>

                                </>)

                            }
                            {
                                (selectedTemplate === TemplateSelection.Default
                                    || selectedTemplate === TemplateSelection.Default_ar
                                    || selectedTemplate === TemplateSelection.infromational
                                    || selectedTemplate === TemplateSelection.infromational_ar
                                    || selectedTemplate === TemplateSelection.uae50)
                                && (<> <Field
                                    size='large'
                                    className={field_styles.styles}
                                    label={{
                                        children: (_: unknown, imageInfoProps: LabelProps) => (
                                            <InfoLabel {...imageInfoProps} info={t('ImageSizeInfoContent') || ''}>
                                                {t('ImageURL')}
                                            </InfoLabel>
                                        ),
                                    }}

                                >
                                    <div
                                        style={{
                                            display: 'grid',
                                            gridTemplateColumns: '1fr auto',
                                            gridTemplateAreas: 'input-area btn-area',
                                        }}
                                    >
                                        <Input
                                            size='large'
                                            style={{ gridColumn: '1' }}
                                            appearance='filled-darker'
                                            value={imageFileName || ''}
                                            placeholder={t('ImageURL')}
                                            onChange={onImageLinkChanged}
                                        />
                                        <Button
                                            style={{ gridColumn: '2', marginLeft: '5px' }}
                                            onClick={handleUploadClick}
                                            size='large'
                                            appearance='secondary'
                                            aria-label={imageFileName ? t('UploadImageSuccessful') : t('UploadImageInfo')}
                                            icon={<ArrowUpload24Regular />}
                                        >
                                            {t('Upload')}
                                        </Button>
                                        <input
                                            type='file'
                                            accept='.jpg, .jpeg, .png, .gif'
                                            style={{ display: 'none' }}
                                            multiple={false}
                                            onChange={handleImageSelection}
                                            ref={fileInput}
                                        />
                                    </div>
                                </Field></>)
                            }

                            {
                                (selectedTemplate === TemplateSelection.infoVideo
                                    || selectedTemplate === TemplateSelection.departmentVideo
                                    || selectedTemplate === TemplateSelection.infoVideo_ar
                                    || selectedTemplate === TemplateSelection.departmentVideo_ar
                                    || selectedTemplate === TemplateSelection.video)
                                && (<> <Field
                                    size='large'
                                    className={field_styles.styles}
                                    label={{
                                        children: (_: unknown, imageInfoProps: LabelProps) => (
                                            <InfoLabel {...imageInfoProps} info={t('PosterSizeInfoContent') || ''}>
                                                {t('posterURL')}
                                            </InfoLabel>
                                        ),
                                    }}
                                    validationMessage={imageUploadErrorMessage}

                                >
                                    <div
                                        style={{
                                            display: 'grid',
                                            gridTemplateColumns: '1fr auto',
                                            gridTemplateAreas: 'input-area btn-area',
                                        }}
                                    >
                                        <Input
                                            size='large'
                                            style={{ gridColumn: '1' }}
                                            appearance='filled-darker'
                                            value={posterFileName || ''}
                                            placeholder={t('PosterURL')}
                                            onChange={onPosterLinkChanged}
                                        />
                                        <Button
                                            style={{ gridColumn: '2', marginLeft: '5px' }}
                                            onClick={handlePosterUploadClick}
                                            size='large'
                                            appearance='secondary'
                                            aria-label={posterFileName ? t('UploadImageSuccessful') : t('UploadImageInfo')}
                                            icon={<ArrowUpload24Regular />}
                                        >
                                            {t('Upload')}
                                        </Button>
                                        <input
                                            type='file'
                                            accept='.jpg, .jpeg, .png, .gif'
                                            style={{ display: 'none' }}
                                            multiple={false}
                                            onChange={handlePosterSelection}
                                            ref={posterFileInput}
                                        />
                                    </div>
                                </Field>
                                    <Field
                                        size='large'
                                        className={field_styles.styles}
                                        label={{
                                            children: (_: unknown, imageInfoProps: LabelProps) => (
                                                <InfoLabel {...imageInfoProps} info={t('VideoSizeInfoContent') || ''}>
                                                    {t('videoURL')}
                                                </InfoLabel>
                                            ),
                                        }}

                                    >
                                        <div
                                            style={{
                                                display: 'grid',
                                                gridTemplateColumns: '1fr auto',
                                                gridTemplateAreas: 'input-area btn-area',
                                            }}
                                        >
                                            <Input
                                                size='large'
                                                style={{ gridColumn: '1' }}
                                                appearance='filled-darker'
                                                value={videoFileName || ''}
                                                placeholder={t('VideoURL')}
                                                onChange={onVideoLinkChanged}
                                            />
                                        </div>

                                    </Field>
                                </>)
                            }

                            <Field size='large' className={field_styles.styles} label={t('Summary')}>
                                <Textarea
                                    size='large'
                                    appearance='filled-darker'
                                    placeholder={t('Summary')}
                                    value={messageState.summary || ''}
                                    onChange={onSummaryChanged}
                                />
                            </Field>
                            <Field size='large' className={field_styles.styles} label={t('Author')}>
                                <Input
                                    placeholder={t('Author')}
                                    size='large'
                                    onChange={onAuthorChanged}
                                    autoComplete='off'
                                    appearance='filled-darker'
                                    value={messageState.author || ''}
                                />
                            </Field>
                            <Field size='large' className={field_styles.styles} label={t('ButtonTitle')}>
                                <Input
                                    size='large'
                                    placeholder={t('ButtonTitle')}
                                    onChange={onBtnTitleChanged}
                                    autoComplete='off'
                                    appearance='filled-darker'
                                    value={messageState.buttonTitle || ''}
                                />
                            </Field>
                            <Field size='large' className={field_styles.styles} label={t('ButtonURL')} validationMessage={btnLinkErrorMessage}>
                                <Input
                                    size='large'
                                    placeholder={t('ButtonURL')}
                                    onChange={onBtnLinkChanged}
                                    type='url'
                                    autoComplete='off'
                                    appearance='filled-darker'
                                    value={messageState.buttonLink || ''}
                                />
                            </Field>
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
                                {showMsgDraftingSpinner && (
                                    <Spinner
                                        role='alert'
                                        id='draftingLoader'
                                        size='small'
                                        label={t('DraftingMessageLabel')}
                                        labelPosition='after'
                                    />
                                )}
                                <Button id='backBtn' style={{ marginLeft: '16px' }} onClick={onBack} disabled={showMsgDraftingSpinner} appearance='secondary'>
                                    {t('Back')}
                                </Button>
                                <Button
                                    style={{ marginLeft: '16px' }}
                                    disabled={isNextBtnDisabled() || showMsgDraftingSpinner}
                                    id='saveBtn'
                                    onClick={onNext}
                                    appearance='primary'
                                >
                                    {t('Next')}
                                </Button>
                            </div>
                        </div>
                    </div>
                </div>
            )}
            {pageSelection === CurrentPageSelection.AudienceSelection && (
                <>
                    <span role='alert' aria-label={t('NewMessageStep2')} />
                    <div className='adaptive-task-grid'>
                        <div className='form-area'>
                            <Label size='large' id='audienceSelectionGroupLabelId'>
                                {t('SendHeadingText')}
                            </Label>
                            <RadioGroup defaultValue={selectedRadioButton} aria-labelledby='audienceSelectionGroupLabelId' onChange={audienceSelectionChange}>
                                <Radio id='radio1' value={AudienceSelection.Teams} label={t('SendToGeneralChannel')} />
                                {selectedRadioButton === AudienceSelection.Teams && (
                                    <div className={cmb_styles.root}>
                                        <Label id={teamsComboId}>Pick team(s)</Label>
                                        {teamsSelectedOptions.length ? (
                                            <ul id={teamsSelectedListId} className={cmb_styles.tagsList} ref={teamsSelectedListRef}>
                                                {/* The "Remove" span is used for naming the buttons without affecting the Combobox name */}
                                                <span id={`${teamsComboId}-remove`} hidden>
                                                    Remove
                                                </span>
                                                {teamsSelectedOptions.map((option, i) => (
                                                    <li key={option.id}>
                                                        <Button
                                                            size='small'
                                                            shape='rounded'
                                                            appearance='subtle'
                                                            icon={<Dismiss12Regular />}
                                                            iconPosition='after'
                                                            onClick={() => onTeamsTagClick(option, i)}
                                                            id={`${teamsComboId}-remove-${i}`}
                                                            aria-labelledby={`${teamsComboId}-remove ${teamsComboId}-remove-${i}`}
                                                        >
                                                            <Persona name={option.name} secondaryText={'Team'} avatar={{ shape: 'square', color: 'colorful' }} />
                                                        </Button>
                                                    </li>
                                                ))}
                                            </ul>
                                        ) : (
                                            <></>
                                        )}

                                        <div className={cmb_styles.root}>
                                            <div className={cmb_styles.combobox}>

                                                <div className={cmb_styles.comboboxInputContainer}>

                                                    <input
                                                        type="text"
                                                        className={cmb_styles.comboboxInput}
                                                        ref={teamsComboboxInputRef}
                                                        placeholder={teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage')}
                                                        onClick={() => handleInputClick(teamsComboboxOptionRef)}
                                                        aria-labelledby={teamsLabelledBy}
                                                    />
                                                    <div className={cmb_styles.comboboxIcon} onClick={() => handleInputClick(teamsComboboxOptionRef)}>
                                                        <ChevronDownRegular />
                                                    </div>

                                                </div>
                                                <ul className={`${cmb_styles.comboboxOptions} combobox-options`} ref={teamsComboboxOptionRef}>
                                                    {teams.map((opt) => (
                                                        <li
                                                            className={cmb_styles.comboboxOption}
                                                            key={opt.id}
                                                            onClick={() => onTeamsSelectOpt(opt)}
                                                        >
                                                            <Persona name={opt.name} secondaryText={'Team'} avatar={{ shape: 'square', color: 'colorful' }} />
                                                        </li>
                                                    ))}
                                                </ul>
                                            </div>
                                        </div>

                                        {/*    <Combobox
                                            multiselect={true}
                                            selectedOptions={teamsSelectedOptions.map((op) => op.id)}
                                            appearance='filled-darker'
                                            size='large'
                                            onOptionSelect={onTeamsSelect}
                                            ref={teamsComboboxInputRef}
                                            aria-labelledby={teamsLabelledBy}
                                            placeholder={teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage')}
                                        >
                                            {teams.map((opt) => (
                                                <Option text={opt.name} value={opt.id} key={opt.id}>
                                                    <Persona name={opt.name} secondaryText={'Team'} avatar={{ shape: 'square', color: 'colorful' }} />
                                                </Option>
                                            ))}
                                        </Combobox>*/}
                                    </div>
                                )}
                                <Radio id='radio2' value={AudienceSelection.Rosters} label={t('SendToRosters')} />
                                {selectedRadioButton === AudienceSelection.Rosters && (
                                    <div className={cmb_styles.root}>
                                        <Label id={rostersComboId}>Pick team(s)</Label>
                                        {rostersSelectedOptions.length ? (
                                            <ul id={rostersSelectedListId} className={cmb_styles.tagsList} ref={rostersSelectedListRef}>
                                                {/* The "Remove" span is used for naming the buttons without affecting the Combobox name */}
                                                <span id={`${rostersComboId}-remove`} hidden>
                                                    Remove
                                                </span>
                                                {rostersSelectedOptions.map((option, i) => (
                                                    <li key={option.id}>
                                                        <Button
                                                            size='small'
                                                            shape='rounded'
                                                            appearance='subtle'
                                                            icon={<Dismiss12Regular />}
                                                            iconPosition='after'
                                                            onClick={() => onRostersTagClick(option, i)}
                                                            id={`${rostersComboId}-remove-${i}`}
                                                            aria-labelledby={`${rostersComboId}-remove ${rostersComboId}-remove-${i}`}
                                                        >
                                                            <Persona name={option.name} secondaryText={'Team'} avatar={{ shape: 'square', color: 'colorful' }} />
                                                        </Button>
                                                    </li>
                                                ))}
                                            </ul>
                                        ) : (
                                            <></>
                                        )}

                                        <div className={cmb_styles.root}>
                                            <div className={cmb_styles.combobox}>

                                                <div className={cmb_styles.comboboxInputContainer}>

                                                    <input
                                                        type="text"
                                                        className={cmb_styles.comboboxInput}
                                                        ref={teamsComboboxInputRef}
                                                        placeholder={teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage')}
                                                        onClick={() => handleInputClick(rosterComboboxOptionRef)}
                                                        aria-labelledby={rostersLabelledBy}
                                                    />
                                                    <div className={cmb_styles.comboboxIcon} onClick={() => handleInputClick(rosterComboboxOptionRef)}>
                                                        <ChevronDownRegular />

                                                    </div>

                                                </div>
                                                <ul className={`${cmb_styles.comboboxOptions} combobox-options`} ref={rosterComboboxOptionRef}>
                                                    {teams.map((opt) => (
                                                        <li
                                                            className={cmb_styles.comboboxOption}
                                                            key={opt.id}
                                                            onClick={() => onRosterSelectOpt(opt)}
                                                        >
                                                            <Persona name={opt.name} secondaryText={'Team'} avatar={{ shape: 'square', color: 'colorful' }} />
                                                        </li>
                                                    ))}
                                                </ul>
                                            </div>
                                        </div>

                                        {/* <Combobox
                                            multiselect={true}
                                            selectedOptions={rostersSelectedOptions.map((op) => op.id)}
                                            appearance='filled-darker'
                                            size='large'
                                            onOptionSelect={onRostersSelect}
                                            ref={rostersComboboxInputRef}
                                            aria-labelledby={rostersLabelledBy}
                                            placeholder={teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage')}
                                        >
                                            {teams.map((opt) => (
                                                <Option text={opt.name} value={opt.id} key={opt.id}>
                                                    <Persona name={opt.name} secondaryText={'Team'} avatar={{ shape: 'square', color: 'colorful' }} />
                                                </Option>
                                            ))}
                                        </Combobox> */}
                                    </div>
                                )}
                                <Radio id='radio3' value={AudienceSelection.AllUsers} label={t('SendToAllUsers')} />
                                <div className={cmb_styles.root}>
                                    {selectedRadioButton === AudienceSelection.AllUsers && (
                                        <Text id='radio3Note' role={allUsersAria} className='info-text'>
                                            {t('SendToAllUsersNote')}
                                        </Text>
                                    )}
                                </div>
                                <Radio id='radio4' value={AudienceSelection.Groups} label={t('SendToGroups')} />
                                {selectedRadioButton === AudienceSelection.Groups && (
                                    <div className={cmb_styles.root}>
                                        {!canAccessGroups && (
                                            <Text role={groupsAria} className='info-text'>
                                                {t('SendToGroupsPermissionNote')}
                                            </Text>
                                        )}
                                        {canAccessGroups && (
                                            <>
                                                <Label id={searchComboId}>Pick group(s)</Label>
                                                {searchSelectedOptions.length ? (
                                                    <ul id={searchSelectedListId} className={cmb_styles.tagsList} ref={searchSelectedListRef}>
                                                        {/* The "Remove" span is used for naming the buttons without affecting the Combobox name */}
                                                        <span id={`${searchComboId}-remove`} hidden>
                                                            Remove
                                                        </span>
                                                        {searchSelectedOptions.map((option, i) => (
                                                            <li key={option.id}>
                                                                <Button
                                                                    size='small'
                                                                    shape='rounded'
                                                                    appearance='subtle'
                                                                    icon={<Dismiss12Regular />}
                                                                    iconPosition='after'
                                                                    onClick={() => onSearchTagClick(option, i)}
                                                                    id={`${searchComboId}-remove-${i}`}
                                                                    aria-labelledby={`${searchComboId}-remove ${searchComboId}-remove-${i}`}
                                                                >
                                                                    <Persona name={option.name} secondaryText={'Group'} avatar={{ color: 'colorful' }} />
                                                                </Button>
                                                            </li>
                                                        ))}
                                                    </ul>
                                                ) : (
                                                    <></>
                                                )}
                                                                                        <div className={cmb_styles.root}>
                                            <div className={cmb_styles.combobox}>

                                                <div className={cmb_styles.comboboxInputContainer}>

                                                    <input
                                                        type="text"
                                                        className={cmb_styles.comboboxInput}
                                                        ref={teamsComboboxInputRef}
                                                        placeholder={teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage')}
                                                                onClick={() => handleInputClick(searchComboboxOptionRef)}
                                                                aria-labelledby={searchLabelledBy}
                                                    />
                                                            <div className={cmb_styles.comboboxIcon} onClick={() => handleInputClick(searchComboboxOptionRef)}>
                                                                <ChevronDownRegular />
                                                    </div>

                                                </div>
                                                <ul className={`${cmb_styles.comboboxOptions} combobox-options`} ref={searchComboboxOptionRef}>
                                                    {queryGroups.map((opt) => (
                                                        <li
                                                            className={cmb_styles.comboboxOption}
                                                            key={opt.id}
                                                            onClick={() => onSearchSelectOpt(opt)}
                                                        >
                                                            <Persona name={opt.name} secondaryText={'Team'} avatar={{ shape: 'square', color: 'colorful' }} />
                                                        </li>
                                                    ))}
                                                </ul>
                                            </div>
                                        </div>

                                                {/* <Combobox
                                                    appearance='filled-darker'
                                                    size='large'
                                                    onOptionSelect={onSearchSelect}
                                                    onChange={onSearchChange}
                                                    aria-labelledby={searchLabelledBy}
                                                    placeholder={'Search for groups'}
                                                >
                                                    {queryGroups.map((opt) => (
                                                        <Option text={opt.name} value={opt.id} key={opt.id}>
                                                            <Persona name={opt.name} secondaryText={'Group'} avatar={{ color: 'colorful' }} />
                                                        </Option>
                                                    ))}
                                                </Combobox> */}
                                                <Text role={groupsAria} className='info-text'>
                                                    {t('SendToGroupsNote')}
                                                </Text>
                                            </>
                                        )}
                                    </div>
                                )}
                            </RadioGroup>
                        </div>
                        <div className='card-area'>
                            <div className={cardAreaBorderClass}>
                                <div className='card-area-2'></div>
                            </div>
                        </div>
                    </div>
                    <div>
                        <div className='fixed-footer'>
                            <div className='footer-action-right'>
                                <div className='footer-actions-flex'>
                                    {showMsgDraftingSpinner && (
                                        <Spinner
                                            role='alert'
                                            id='draftingLoader'
                                            size='small'
                                            label={t('DraftingMessageLabel')}
                                            labelPosition='after'
                                        />
                                    )}
                                    <Button id='backBtn' style={{ marginLeft: '16px' }} onClick={onBack} disabled={showMsgDraftingSpinner} appearance='secondary'>
                                        {t('Back')}
                                    </Button>
                                    <Button
                                        style={{ marginLeft: '16px' }}
                                        disabled={isSaveBtnDisabled() || showMsgDraftingSpinner}
                                        id='saveBtn'
                                        onClick={onSave}
                                        appearance='primary'
                                    >
                                        {t('SaveAsDraft')}
                                    </Button>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            )}
        </>
    );
};
