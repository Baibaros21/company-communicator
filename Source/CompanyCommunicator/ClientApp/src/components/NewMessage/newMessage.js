"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.NewMessage = void 0;
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
require("./newMessages.scss");
var react_icons_1 = require("@fluentui/react-icons");
var AdaptiveCards = require("adaptivecards");
var React = require("react");
var react_i18next_1 = require("react-i18next");
var react_router_dom_1 = require("react-router-dom");
var validator_1 = require("validator");
var react_components_1 = require("@fluentui/react-components");
var unstable_1 = require("@fluentui/react-components/unstable");
var react_icons_2 = require("@fluentui/react-icons");
var microsoftTeams = require("@microsoft/teams-js");
var ACData = require("adaptivecards-templating");
var actions_1 = require("../../actions");
var messageListApi_1 = require("../../apis/messageListApi");
var store_1 = require("../../store");
var adaptiveCard_1 = require("../AdaptiveCard/adaptiveCard");
var validImageTypes = ['image/gif', 'image/jpeg', 'image/png', 'image/jpg'];
var useComboboxStyles = (0, react_components_1.makeStyles)({
    root: __assign(__assign({ display: 'grid', gridTemplateRows: 'repeat(1fr)', justifyItems: 'start' }, react_components_1.shorthands.gap('2px')), { paddingLeft: '36px' }),
    combobox: {
        position: 'relative',
        width: '100%',
        backgroundColor: react_components_1.tokens.colorNeutralBackground3,
    },
    comboboxInputContainer: __assign(__assign(__assign(__assign({ display: 'flex', alignItems: 'center', width: '100%', borderBlockColor: react_components_1.tokens.colorTransparentStroke, borderInlineColor: react_components_1.tokens.colorTransparentStroke }, react_components_1.shorthands.borderRadius("4px")), react_components_1.shorthands.borderWidth('0px')), react_components_1.shorthands.borderStyle('solid')), react_components_1.shorthands.padding('1px', '4px')),
    comboboxInput: __assign(__assign(__assign(__assign({ width: '100%', backgroundColor: react_components_1.tokens.colorNeutralBackground3, fontSize: react_components_1.tokens.fontSizeBase400, fontWeight: react_components_1.tokens.fontWeightRegular, color: react_components_1.tokens.colorNeutralForeground1, height: '38px', paddingLeft: react_components_1.tokens.spacingHorizontalM, borderBlockColor: react_components_1.tokens.colorTransparentStroke, borderInlineColor: react_components_1.tokens.colorTransparentStroke }, react_components_1.shorthands.borderRadius("4px")), react_components_1.shorthands.borderWidth('0px')), react_components_1.shorthands.borderStyle('solid')), react_components_1.shorthands.padding('1px', '2px')),
    comboboxIcon: __assign(__assign({}, react_components_1.shorthands.padding('10px')), { cursor: 'pointer' }),
    comboboxOptions: __assign(__assign({ listStyleType: 'none', backgroundColor: react_components_1.tokens.colorTransparentBackground, position: 'absolute', width: '100%' }, react_components_1.shorthands.borderRadius(react_components_1.tokens.borderRadiusMedium)), { maxHeight: '200px', overflowY: 'auto', zIndex: 1000, display: 'none' }),
    comboboxOption: {
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        backgroundColor: react_components_1.tokens.colorNeutralBackground1,
    },
    comboboxOptionHover: {
        backgroundColor: react_components_1.tokens.colorNeutralBackground1Hover,
    },
    tagsList: {
        listStyleType: 'none',
        marginBottom: react_components_1.tokens.spacingVerticalXXS,
        marginTop: 0,
        paddingLeft: 0,
        // display: "flex",
        gridGap: react_components_1.tokens.spacingHorizontalXXS,
    },
});
var useFieldStyles = (0, react_components_1.makeStyles)({
    styles: {
        marginBottom: react_components_1.tokens.spacingVerticalM,
        gridGap: react_components_1.tokens.spacingHorizontalXXS,
    },
});
var AudienceSelection;
(function (AudienceSelection) {
    AudienceSelection["Teams"] = "Teams";
    AudienceSelection["Rosters"] = "Rosters";
    AudienceSelection["Groups"] = "Groups";
    AudienceSelection["AllUsers"] = "AllUsers";
    AudienceSelection["None"] = "None";
})(AudienceSelection || (AudienceSelection = {}));
var CurrentPageSelection;
(function (CurrentPageSelection) {
    CurrentPageSelection["TemplateCreation"] = "TemplateCreation";
    CurrentPageSelection["CardCreation"] = "CardCreation";
    CurrentPageSelection["AudienceSelection"] = "AudienceSelection";
})(CurrentPageSelection || (CurrentPageSelection = {}));
var card;
var MAX_SELECTED_TEAMS_NUM = 20;
var NewMessage = function () {
    var fileInput = React.createRef();
    var posterFileInput = React.createRef();
    var t = (0, react_i18next_1.useTranslation)().t;
    var id = (0, react_router_dom_1.useParams)().id;
    var dispatch = (0, store_1.useAppDispatch)();
    var Templates = (0, store_1.useAppSelector)(function (state) { return state.messages; }).cardTemplates.payload;
    var teams = (0, store_1.useAppSelector)(function (state) { return state.messages; }).teamsData.payload;
    var groups = (0, store_1.useAppSelector)(function (state) { return state.messages; }).groups.payload;
    var queryGroups = (0, store_1.useAppSelector)(function (state) { return state.messages; }).queryGroups.payload;
    var canAccessGroups = (0, store_1.useAppSelector)(function (state) { return state.messages; }).verifyGroup.payload;
    var _a = React.useState(AudienceSelection.None), selectedRadioButton = _a[0], setSelectedRadioButton = _a[1];
    var _b = React.useState(store_1.TemplateSelection.Default), selectedTemplate = _b[0], setSelectedTemplate = _b[1];
    var _c = React.useState(CurrentPageSelection.TemplateCreation), pageSelection = _c[0], setPageSelection = _c[1];
    var _d = React.useState(false), allUsersState = _d[0], setAllUsersState = _d[1];
    var _e = React.useState(''), imageFileName = _e[0], setImageFileName = _e[1];
    var _f = React.useState(''), videoFileName = _f[0], setVideoFileName = _f[1];
    var _g = React.useState(''), InternalAppId = _g[0], setInternalAppId = _g[1];
    var _h = React.useState(''), posterFileName = _h[0], setPosterFileName = _h[1];
    var _j = React.useState(''), imageUploadErrorMessage = _j[0], setImageUploadErrorMessage = _j[1];
    var _k = React.useState(''), titleErrorMessage = _k[0], setTitleErrorMessage = _k[1];
    var _l = React.useState(''), btnLinkErrorMessage = _l[0], setBtnLinkErrorMessage = _l[1];
    var _m = React.useState(false), showMsgDraftingSpinner = _m[0], setShowMsgDraftingSpinner = _m[1];
    var _o = React.useState(false), isCardReady = _o[0], setIsCardReady = _o[1];
    var _p = React.useState('none'), allUsersAria = _p[0], setAllUserAria = _p[1];
    var _q = React.useState('none'), groupsAria = _q[0], setGroupsAria = _q[1];
    var _r = React.useState(''), cardAreaBorderClass = _r[0], setCardAreaBorderClass = _r[1];
    var _s = React.useState({
        logoFileName: "",
        logoLink: "",
        bannerLink: "",
        bannerFileName: ""
    }), defaultsState = _s[0], setDefaultState = _s[1];
    var _t = React.useState({
        title: '',
        template: store_1.TemplateSelection.Default,
        teams: [],
        rosters: [],
        groups: [],
        allUsers: false,
    }), messageState = _t[0], setMessageState = _t[1];
    // Handle selectedOptions both when an option is selected or deselected in the Combobox,
    // and when an option is removed by clicking on a tag
    var _u = React.useState([]), teamsSelectedOptions = _u[0], setTeamsSelectedOptions = _u[1];
    var _v = React.useState([]), rostersSelectedOptions = _v[0], setRostersSelectedOptions = _v[1];
    var _w = React.useState([]), searchSelectedOptions = _w[0], setSearchSelectedOptions = _w[1];
    //const [cardTemplates, setCardTemplates] = React.useState<[ITemplates]>();
    React.useEffect(function () {
        (0, actions_1.GetTeamsDataAction)(dispatch);
        (0, actions_1.VerifyGroupAccessAction)(dispatch);
        (0, actions_1.GetAllCardTemplatesAction)(dispatch);
        getDefaultsItem();
    }, []);
    React.useEffect(function () {
        if (Templates && Templates.length > 0) {
            getCurrentCardTemplate(selectedTemplate);
            /*            setDefaultCard(card);
            */ updateAdaptiveCard();
            setIsCardReady(true);
        }
    }, [Templates]);
    React.useEffect(function () {
        if (isCardReady) {
            updateAdaptiveCard();
        }
    }, [pageSelection]);
    React.useEffect(function () {
        if (isCardReady) {
            if (messageState.title !== "")
                (0, adaptiveCard_1.setCardTitle)(card, messageState.title);
            if (messageState.imageLink !== "")
                (0, adaptiveCard_1.setCardImageLink)(card, messageState.imageLink);
            if (messageState.department !== "")
                (0, adaptiveCard_1.setCardDeptTitle)(card, messageState.department);
            if (messageState.posterLink !== "")
                (0, adaptiveCard_1.setCardVideoPlayerPoster)(card, messageState.posterLink);
            if (messageState.videoLink !== "")
                (0, adaptiveCard_1.setCardVideoPlayerUrl)(card, messageState.videoLink);
            if (messageState.department !== "")
                (0, adaptiveCard_1.setCardDeptTitle)(card, messageState.department);
            if (messageState.summary !== "")
                (0, adaptiveCard_1.setCardSummary)(card, messageState.summary);
            if (messageState.author !== "")
                (0, adaptiveCard_1.setCardAuthor)(card, messageState.author);
            (0, adaptiveCard_1.setCardLogo)(card, defaultsState.logoLink);
            (0, adaptiveCard_1.setCardBanner)(card, defaultsState.bannerLink);
            if (messageState.buttonTitle !== "")
                (0, adaptiveCard_1.setCardBtn)(card, messageState.buttonTitle, messageState.buttonLink);
            if (!messageState.title && !messageState.imageLink && !messageState.summary && !messageState.author && !messageState.buttonTitle && !messageState.buttonLink) {
                getCurrentCardTemplate(selectedTemplate);
                /*                    setDefaultCard(card);
                */ }
            updateAdaptiveCard();
        }
    }, [messageState, isCardReady]);
    React.useEffect(function () {
        if (id) {
            (0, actions_1.GetGroupsAction)(dispatch, { id: id });
            getDraftNotificationItem(id);
        }
    }, [id]);
    React.useEffect(function () {
        setTeamsSelectedOptions([]);
        setRostersSelectedOptions([]);
        setSearchSelectedOptions([]);
        setAllUsersState(false);
        if (teams && teams.length > 0) {
            var teamsSelected = teams.filter(function (c) { return messageState.teams.some(function (s) { return s === c.id; }); });
            setTeamsSelectedOptions(teamsSelected || []);
            var roastersSelected = teams.filter(function (c) { return messageState.rosters.some(function (s) { return s === c.id; }); });
            setRostersSelectedOptions(roastersSelected || []);
        }
        if (groups && groups.length > 0) {
            var groupsSelected = groups.filter(function (c) { return messageState.groups.some(function (s) { return s === c.id; }); });
            setSearchSelectedOptions(groupsSelected || []);
        }
        if (messageState.allUsers) {
            setAllUsersState(true);
        }
    }, [teams, groups, messageState.teams, messageState.rosters, messageState.allUsers, messageState.groups]);
    var getCurrentCardTemplate = function (cardtemplate) {
        var _a;
        console.log(Templates);
        var currentTemplate = (_a = Templates === null || Templates === void 0 ? void 0 : Templates.find(function (template) { return template.name === cardtemplate; })) === null || _a === void 0 ? void 0 : _a.card;
        console.log(card);
        var cardTemplate = new ACData.Template(JSON.parse(currentTemplate));
        card = cardTemplate.expand({
            $root: {}
        });
        (0, adaptiveCard_1.setCardLogo)(card, defaultsState.logoLink);
        (0, adaptiveCard_1.setCardBanner)(card, defaultsState.bannerLink);
    };
    var getDefaultsItem = function () { return __awaiter(void 0, void 0, void 0, function () {
        var error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 3, , 4]);
                    return [4 /*yield*/, (0, messageListApi_1.getDefaultData)().then(function (response) {
                            var defaultImages = response.data;
                            setDefaultState({
                                logoFileName: defaultImages.logoFileName,
                                logoLink: defaultImages.logoLink,
                                bannerFileName: defaultImages.bannerFileName,
                                bannerLink: defaultImages.bannerLink
                            });
                        })];
                case 1:
                    _a.sent();
                    return [4 /*yield*/, (0, messageListApi_1.getAppId)().then(function (reponse) {
                            setInternalAppId(reponse.data);
                        })];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 4];
                case 3:
                    error_1 = _a.sent();
                    card = getCurrentCardTemplate(store_1.TemplateSelection.Default);
                    updateAdaptiveCard();
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var getDraftNotificationItem = function (id) { return __awaiter(void 0, void 0, void 0, function () {
        var error_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, (0, messageListApi_1.getDraftNotification)(id).then(function (response) {
                            var draftMessageDetail = response.data;
                            if (draftMessageDetail.teams.length > 0) {
                                setSelectedRadioButton(AudienceSelection.Teams);
                            }
                            else if (draftMessageDetail.rosters.length > 0) {
                                setSelectedRadioButton(AudienceSelection.Rosters);
                            }
                            else if (draftMessageDetail.groups.length > 0) {
                                setSelectedRadioButton(AudienceSelection.Groups);
                            }
                            else if (draftMessageDetail.allUsers) {
                                setSelectedRadioButton(AudienceSelection.AllUsers);
                            }
                            setMessageState(__assign(__assign({}, messageState), { id: draftMessageDetail.id, title: draftMessageDetail.title, department: draftMessageDetail.department, imageLink: draftMessageDetail.imageLink, posterLink: draftMessageDetail.posterLink, videoLink: draftMessageDetail.videoLink, summary: draftMessageDetail.summary, author: draftMessageDetail.author, buttonTitle: draftMessageDetail.buttonTitle, buttonLink: draftMessageDetail.buttonLink, teams: draftMessageDetail.teams, rosters: draftMessageDetail.rosters, groups: draftMessageDetail.groups, allUsers: draftMessageDetail.allUsers, template: draftMessageDetail.template }));
                            setSelectedTemplate(draftMessageDetail.template);
                            (0, adaptiveCard_1.setCardTitle)(card, draftMessageDetail.title);
                            (0, adaptiveCard_1.setCardDeptTitle)(card, draftMessageDetail.department);
                            (0, adaptiveCard_1.setCardImageLink)(card, draftMessageDetail.imageLink);
                            (0, adaptiveCard_1.setCardVideoPlayerPoster)(card, draftMessageDetail.posterLink);
                            (0, adaptiveCard_1.setCardVideoPlayerUrl)(card, draftMessageDetail.videoLink);
                            (0, adaptiveCard_1.setCardDeptTitle)(card, draftMessageDetail.department);
                            (0, adaptiveCard_1.setCardSummary)(card, draftMessageDetail.summary);
                            (0, adaptiveCard_1.setCardAuthor)(card, draftMessageDetail.author);
                            (0, adaptiveCard_1.setCardBtn)(card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink);
                            (0, adaptiveCard_1.setCardLogo)(card, defaultsState.logoLink);
                            (0, adaptiveCard_1.setCardBanner)(card, defaultsState.bannerLink);
                            setTeamsSelectedOptions(draftMessageDetail.template);
                            updateAdaptiveCard();
                        })];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    error_2 = _a.sent();
                    return [2 /*return*/, error_2];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var templateSelectionChange = function (ev, data) {
        var input = data.value;
        setSelectedTemplate(store_1.TemplateSelection[input]);
        getCurrentCardTemplate(store_1.TemplateSelection[input]);
        /*        setDefaultCard(card);
        */ updateAdaptiveCard();
    };
    var updateAdaptiveCard = function () {
        var adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        var renderCard = adaptiveCard.render();
        if (renderCard && pageSelection === CurrentPageSelection.CardCreation) {
            document.getElementsByClassName('card-area-1')[0].innerHTML = '';
            document.getElementsByClassName('card-area-1')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        }
        else if (renderCard && pageSelection === CurrentPageSelection.AudienceSelection) {
            document.getElementsByClassName('card-area-2')[0].innerHTML = '';
            document.getElementsByClassName('card-area-2')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        }
        else if (renderCard && pageSelection === CurrentPageSelection.TemplateCreation) {
            document.getElementsByClassName('card-area-3')[0].innerHTML = '';
            document.getElementsByClassName('card-area-3')[0].appendChild(renderCard);
            setCardAreaBorderClass('card-area-border');
        }
        adaptiveCard.onExecuteAction = function (action) {
            window.open(action.url, '_blank');
        };
    };
    var checkValidSizeOfImage = function (resizedImageAsBase64) {
        var stringLength = resizedImageAsBase64.length - 'data:image/png;base64,'.length;
        var sizeInBytes = 4 * Math.ceil(stringLength / 3) * 0.5624896334383812;
        var sizeInKb = sizeInBytes / 1000;
        if (sizeInKb <= 1024)
            return true;
        else
            return false;
    };
    var handleUploadClick = function (event) {
        if (fileInput.current) {
            fileInput.current.click();
        }
    };
    var handlePosterUploadClick = function (event) {
        if (posterFileInput.current) {
            posterFileInput.current.click();
        }
    };
    var handlePosterSelection = function () {
        var _a;
        var file = (_a = posterFileInput.current) === null || _a === void 0 ? void 0 : _a.files[0];
        imageselection(file, "poster");
    };
    var handleImageSelection = function () {
        var _a;
        var file = (_a = fileInput.current) === null || _a === void 0 ? void 0 : _a.files[0];
        imageselection(file, "image");
    };
    var imageselection = function (file, field) {
        if (file) {
            var fileType = file['type'];
            var mimeType_1 = file.type;
            if (!validImageTypes.includes(fileType)) {
                setImageUploadErrorMessage(t('ErrorImageTypesMessage'));
                return;
            }
            var fileReader_1 = new FileReader();
            fileReader_1.readAsDataURL(file);
            fileReader_1.onload = function () {
                var image = new Image();
                image.src = fileReader_1.result;
                var resizedImageAsBase64 = fileReader_1.result;
                image.onload = function (e) {
                    var MAX_WIDTH = 1024;
                    if (image.width > MAX_WIDTH) {
                        var canvas = document.createElement('canvas');
                        canvas.width = MAX_WIDTH;
                        canvas.height = ~~(image.height * (MAX_WIDTH / image.width));
                        var context = canvas.getContext('2d', { alpha: false });
                        if (!context) {
                            return;
                        }
                        context.drawImage(image, 0, 0, canvas.width, canvas.height);
                        resizedImageAsBase64 = canvas.toDataURL(mimeType_1);
                    }
                };
                if (!checkValidSizeOfImage(resizedImageAsBase64)) {
                    setImageUploadErrorMessage(t('ErrorImageSizeMessage'));
                    return;
                }
                if (resizedImageAsBase64 && field === 'image') {
                    setImageFileName(file['name']);
                    setImageUploadErrorMessage('');
                    (0, adaptiveCard_1.setCardImageLink)(card, resizedImageAsBase64);
                    setMessageState(__assign(__assign({}, messageState), { imageLink: resizedImageAsBase64 }));
                }
                else if (resizedImageAsBase64 && field === 'poster') {
                    setPosterFileName(file['name']);
                    setImageUploadErrorMessage('');
                    (0, adaptiveCard_1.setCardVideoPlayerPoster)(card, resizedImageAsBase64);
                    setMessageState(__assign(__assign({}, messageState), { posterLink: resizedImageAsBase64 }));
                }
                updateAdaptiveCard();
            };
        }
    };
    var isSaveBtnDisabled = function () {
        var msg_page_conditions = messageState.title !== '' && imageUploadErrorMessage === '' && btnLinkErrorMessage === '';
        var aud_page_conditions = (teamsSelectedOptions.length > 0 && selectedRadioButton === AudienceSelection.Teams) ||
            (rostersSelectedOptions.length > 0 && selectedRadioButton === AudienceSelection.Rosters) ||
            (searchSelectedOptions.length > 0 && selectedRadioButton === AudienceSelection.Groups) ||
            selectedRadioButton === AudienceSelection.AllUsers;
        if (msg_page_conditions && aud_page_conditions) {
            return false;
        }
        else {
            return true;
        }
    };
    var isNextBtnDisabled = function () {
        if (messageState.title !== '' && imageUploadErrorMessage === '' && btnLinkErrorMessage === '') {
            return false;
        }
        else {
            return true;
        }
    };
    var onSave = function () {
        var finalSelectedTeams = [];
        var finalSelectedRosters = [];
        var finalSelectedGroups = [];
        var finalAllUsers = false;
        if (selectedRadioButton === AudienceSelection.Teams) {
            finalSelectedTeams = __spreadArray([], teams.filter(function (t1) { return teamsSelectedOptions.some(function (sp) { return sp.id === t1.id; }); }).map(function (t2) { return t2.id; }), true);
        }
        if (selectedRadioButton === AudienceSelection.Rosters) {
            finalSelectedRosters = __spreadArray([], teams.filter(function (t1) { return rostersSelectedOptions.some(function (sp) { return sp.id === t1.id; }); }).map(function (t2) { return t2.id; }), true);
        }
        if (selectedRadioButton === AudienceSelection.Groups) {
            finalSelectedGroups = __spreadArray([], searchSelectedOptions.map(function (g) { return g.id; }), true);
        }
        if (selectedRadioButton === AudienceSelection.AllUsers) {
            finalAllUsers = allUsersState;
        }
        var finalMessage = __assign(__assign({}, messageState), { template: selectedTemplate, teams: finalSelectedTeams, rosters: finalSelectedRosters, groups: finalSelectedGroups, allUsers: finalAllUsers });
        setShowMsgDraftingSpinner(true);
        if (id) {
            editDraftMessage(finalMessage);
        }
        else {
            postDraftMessage(finalMessage);
        }
    };
    var editDraftMessage = function (msg) {
        try {
            (0, messageListApi_1.updateDraftNotification)(msg)
                .then(function () {
                (0, actions_1.GetDraftMessagesSilentAction)(dispatch);
            })
                .finally(function () {
                setShowMsgDraftingSpinner(false);
                microsoftTeams.tasks.submitTask();
            });
        }
        catch (error) {
            return error;
        }
    };
    var postDraftMessage = function (msg) {
        try {
            (0, messageListApi_1.createDraftNotification)(msg)
                .then(function () {
                (0, actions_1.GetDraftMessagesSilentAction)(dispatch);
            })
                .finally(function () {
                setShowMsgDraftingSpinner(false);
                microsoftTeams.tasks.submitTask();
            });
        }
        catch (error) {
            return error;
        }
    };
    var onNext = function (event) {
        switch (pageSelection) {
            case (CurrentPageSelection.TemplateCreation):
                setPageSelection(CurrentPageSelection.CardCreation);
                break;
            case (CurrentPageSelection.CardCreation):
                setPageSelection(CurrentPageSelection.AudienceSelection);
                break;
            default:
        }
    };
    var onBack = function (event) {
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
    var onTitleChanged = function (event) {
        if (event.target.value === '') {
            setTitleErrorMessage('Title is required.');
        }
        else {
            setTitleErrorMessage('');
        }
        (0, adaptiveCard_1.setCardTitle)(card, event.target.value);
        setMessageState(__assign(__assign({}, messageState), { title: event.target.value }));
        updateAdaptiveCard();
    };
    var onDeptChanged = function (event) {
        (0, adaptiveCard_1.setCardDeptTitle)(card, event.target.value);
        setMessageState(__assign(__assign({}, messageState), { department: event.target.value }));
        updateAdaptiveCard();
    };
    var onImageLinkChanged = function (event) {
        var urlOrDataUrl = event.target.value;
        var isGoodLink = true;
        setImageFileName(urlOrDataUrl);
        if (!(urlOrDataUrl === '' ||
            urlOrDataUrl.startsWith('https://') ||
            urlOrDataUrl.startsWith('data:image/png;base64,') ||
            urlOrDataUrl.startsWith('data:image/jpeg;base64,') ||
            urlOrDataUrl.startsWith('data:image/gif;base64,'))) {
            isGoodLink = false;
            setImageUploadErrorMessage(t('ErrorURLMessage'));
        }
        else {
            isGoodLink = true;
            setImageUploadErrorMessage(t(''));
        }
        if (isGoodLink) {
            setMessageState(__assign(__assign({}, messageState), { imageLink: urlOrDataUrl }));
            (0, adaptiveCard_1.setCardImageLink)(card, event.target.value);
            updateAdaptiveCard();
        }
    };
    var onPosterLinkChanged = function (event) {
        var urlOrDataUrl = event.target.value;
        var isGoodLink = true;
        setPosterFileName(urlOrDataUrl);
        if (!(urlOrDataUrl === '' ||
            urlOrDataUrl.startsWith('https://') ||
            urlOrDataUrl.startsWith('data:image/png;base64,') ||
            urlOrDataUrl.startsWith('data:image/jpeg;base64,') ||
            urlOrDataUrl.startsWith('data:image/gif;base64,'))) {
            isGoodLink = false;
            setImageUploadErrorMessage(t('ErrorURLMessage'));
        }
        else {
            isGoodLink = true;
            setImageUploadErrorMessage(t(''));
        }
        if (isGoodLink) {
            setMessageState(__assign(__assign({}, messageState), { posterLink: urlOrDataUrl }));
            (0, adaptiveCard_1.setCardVideoPlayerPoster)(card, event.target.value);
            updateAdaptiveCard();
        }
    };
    var onVideoLinkChanged = function (event) {
        var urlOrDataUrl = event.target.value;
        var isGoodLink = true;
        setVideoFileName(urlOrDataUrl);
        if (!(urlOrDataUrl === '' ||
            urlOrDataUrl.startsWith('https://'))) {
            isGoodLink = false;
            setImageUploadErrorMessage(t('ErrorURLMessage'));
        }
        else {
            isGoodLink = true;
            setImageUploadErrorMessage(t(''));
        }
        var url = "https://teams.microsoft.com/l/task/" + InternalAppId
            + "?url=" + "https://companycommunicator.blueridgeit.com/videoplayer/" + urlOrDataUrl
            + "&height=large&width=large&title=ADVideo";
        if (isGoodLink) {
            setMessageState(__assign(__assign({}, messageState), { videoLink: url }));
            (0, adaptiveCard_1.setCardVideoPlayerUrl)(card, url);
            updateAdaptiveCard();
        }
    };
    var onSummaryChanged = function (event) {
        (0, adaptiveCard_1.setCardSummary)(card, event.target.value);
        setMessageState(__assign(__assign({}, messageState), { summary: event.target.value }));
        updateAdaptiveCard();
    };
    var onAuthorChanged = function (event) {
        (0, adaptiveCard_1.setCardAuthor)(card, event.target.value);
        setMessageState(__assign(__assign({}, messageState), { author: event.target.value }));
        updateAdaptiveCard();
    };
    var onBtnTitleChanged = function (event) {
        (0, adaptiveCard_1.setCardBtn)(card, event.target.value, messageState.buttonLink);
        setMessageState(__assign(__assign({}, messageState), { buttonTitle: event.target.value }));
        updateAdaptiveCard();
    };
    var onBtnLinkChanged = function (event) {
        if (validator_1.default.isURL(event.target.value, { require_protocol: true, protocols: ['https'] }) || event.target.value === '') {
            setBtnLinkErrorMessage('');
        }
        else {
            setBtnLinkErrorMessage("".concat(event.target.value, " is invalid. Please enter a valid https URL"));
        }
        (0, adaptiveCard_1.setCardBtn)(card, messageState.buttonTitle, event.target.value);
        setMessageState(__assign(__assign({}, messageState), { buttonLink: event.target.value }));
        updateAdaptiveCard();
    };
    // generate ids for handling labelling
    var teamsComboId = (0, react_components_1.useId)('teams-combo-multi');
    var teamsSelectedListId = "".concat(teamsComboId, "-selection");
    var rostersComboId = (0, react_components_1.useId)('rosters-combo-multi');
    var rostersSelectedListId = "".concat(rostersComboId, "-selection");
    var searchComboId = (0, react_components_1.useId)('search-combo-multi');
    var searchSelectedListId = "".concat(searchComboId, "-selection");
    // refs for managing focus when removing tags
    var teamsSelectedListRef = React.useRef(null);
    var teamsComboboxInputRef = React.useRef(null);
    var teamsComboboxOptionRef = React.useRef(null);
    var rostersSelectedListRef = React.useRef(null);
    var rostersComboboxInputRef = React.useRef(null);
    var rosterComboboxOptionRef = React.useRef(null);
    var searchSelectedListRef = React.useRef(null);
    var searchComboboxInputRef = React.useRef(null);
    var searchComboboxOptionRef = React.useRef(null);
    //Custom Combobox functions
    var handleInputClick = function (optionList) {
        if (optionList.current && optionList.current.style) {
            optionList.current.style.display = 'block';
        }
    };
    var handleClickOutside = function (event) {
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
    React.useEffect(function () {
        document.addEventListener('click', handleClickOutside);
        return function () {
            document.removeEventListener('click', handleClickOutside);
        };
    }, []);
    //functions for handling teams. groups and users
    var onTeamsSelect = function (event, data) {
        if (data.selectedOptions.length <= MAX_SELECTED_TEAMS_NUM) {
            setTeamsSelectedOptions(teams.filter(function (t1) { return data.selectedOptions.some(function (t2) { return t2 === t1.id; }); }));
        }
    };
    var onTeamsSelectOpt = function (opt) {
        if (teamsSelectedOptions.length < MAX_SELECTED_TEAMS_NUM) {
            var teamExists = teamsSelectedOptions.some(function (team) { return team.id === opt.id; });
            if (!teamExists) {
                setTeamsSelectedOptions(__spreadArray(__spreadArray([], teamsSelectedOptions, true), [opt], false));
            }
        }
    };
    var onRostersSelect = function (event, data) {
        if (data.selectedOptions.length <= MAX_SELECTED_TEAMS_NUM) {
            setRostersSelectedOptions(teams.filter(function (t1) { return data.selectedOptions.some(function (t2) { return t2 === t1.id; }); }));
        }
    };
    var onRosterSelectOpt = function (opt) {
        if (rostersSelectedOptions.length < MAX_SELECTED_TEAMS_NUM) {
            var teamExists = rostersSelectedOptions.some(function (team) { return team.id === opt.id; });
            if (!teamExists) {
                setRostersSelectedOptions(__spreadArray(__spreadArray([], rostersSelectedOptions, true), [opt], false));
            }
        }
    };
    var onSearchSelect = function (event, data) {
        if (data.optionText && !searchSelectedOptions.find(function (x) { return x.id === data.optionValue; })) {
            setSearchSelectedOptions(__spreadArray(__spreadArray([], searchSelectedOptions, true), [{ id: data.optionValue, name: data.optionText }], false));
        }
    };
    var onSearchSelectOpt = function (opt) {
        if (searchSelectedOptions.length < MAX_SELECTED_TEAMS_NUM) {
            var teamExists = searchSelectedOptions.some(function (team) { return team.id === opt.id; });
            if (!teamExists) {
                setSearchSelectedOptions(__spreadArray(__spreadArray([], searchSelectedOptions, true), [opt], false));
            }
        }
    };
    var onSearchChange = function (event) {
        if (event && event.target && event.target.value) {
            var q = encodeURIComponent(event.target.value);
            (0, actions_1.SearchGroupsAction)(dispatch, { query: q });
        }
    };
    var onTeamsTagClick = function (option, index) {
        var _a, _b;
        // remove selected option
        setTeamsSelectedOptions(teamsSelectedOptions.filter(function (o) { return o.id !== option.id; }));
        // focus previous or next option, defaulting to focusing back to the combo input
        var indexToFocus = index === 0 ? 1 : index - 1;
        var optionToFocus = (_a = teamsSelectedListRef.current) === null || _a === void 0 ? void 0 : _a.querySelector("#".concat(teamsComboId, "-remove-").concat(indexToFocus));
        if (optionToFocus) {
            optionToFocus.focus();
        }
        else {
            (_b = teamsComboboxInputRef.current) === null || _b === void 0 ? void 0 : _b.focus();
        }
    };
    var onRostersTagClick = function (option, index) {
        var _a, _b;
        // remove selected option
        setRostersSelectedOptions(rostersSelectedOptions.filter(function (o) { return o.id !== option.id; }));
        // focus previous or next option, defaulting to focusing back to the combo input
        var indexToFocus = index === 0 ? 1 : index - 1;
        var optionToFocus = (_a = rostersSelectedListRef.current) === null || _a === void 0 ? void 0 : _a.querySelector("#".concat(rostersComboId, "-remove-").concat(indexToFocus));
        if (optionToFocus) {
            optionToFocus.focus();
        }
        else {
            (_b = rostersComboboxInputRef.current) === null || _b === void 0 ? void 0 : _b.focus();
        }
    };
    var onSearchTagClick = function (option, index) {
        var _a, _b;
        // remove selected option
        setSearchSelectedOptions(searchSelectedOptions.filter(function (o) { return o.id !== option.id; }));
        // focus previous or next option, defaulting to focusing back to the combo input
        var indexToFocus = index === 0 ? 1 : index - 1;
        var optionToFocus = (_a = searchSelectedListRef.current) === null || _a === void 0 ? void 0 : _a.querySelector("#".concat(searchComboId, "-remove-").concat(indexToFocus));
        if (optionToFocus) {
            optionToFocus.focus();
        }
        else {
            (_b = searchComboboxInputRef.current) === null || _b === void 0 ? void 0 : _b.focus();
        }
    };
    var teamsLabelledBy = teamsSelectedOptions.length > 0 ? "".concat(teamsComboId, " ").concat(teamsSelectedListId) : teamsComboId;
    var rostersLabelledBy = rostersSelectedOptions.length > 0 ? "".concat(rostersComboId, " ").concat(rostersSelectedListId) : rostersComboId;
    var searchLabelledBy = searchSelectedOptions.length > 0 ? "".concat(searchComboId, " ").concat(searchSelectedListId) : searchComboId;
    var cmb_styles = useComboboxStyles();
    var field_styles = useFieldStyles();
    var audienceSelectionChange = function (ev, data) {
        var input = data.value;
        setSelectedRadioButton(AudienceSelection[input]);
        if (AudienceSelection[input] === AudienceSelection.AllUsers) {
            setAllUsersState(true);
        }
        else if (allUsersState) {
            setAllUsersState(false);
        }
        AudienceSelection[input] === AudienceSelection.AllUsers ? setAllUserAria('alert') : setAllUserAria('none');
        AudienceSelection[input] === AudienceSelection.Groups ? setGroupsAria('alert') : setGroupsAria('none');
    };
    return (React.createElement(React.Fragment, null,
        pageSelection === CurrentPageSelection.TemplateCreation && Templates && Templates.length > 0 && (React.createElement(React.Fragment, null,
            React.createElement("span", { role: 'alert', "aria-label": t('NewMessageStep2') }),
            React.createElement("div", { className: 'adaptive-task-grid' },
                React.createElement("div", { className: 'form-area' },
                    React.createElement(react_components_1.Label, { size: 'large', id: 'TemplateSelectionGroupLabelId' }, t('SendHeadingText')),
                    React.createElement(react_components_1.RadioGroup, { defaultValue: selectedTemplate, "aria-labelledby": 'TemplateSelectionGroupLabelId', onChange: templateSelectionChange },
                        React.createElement(react_components_1.Radio, { id: 'radio1', value: "Default", label: store_1.TemplateSelection.Default }),
                        React.createElement(react_components_1.Radio, { id: 'radio2', value: "infromational", label: store_1.TemplateSelection.infromational }),
                        React.createElement(react_components_1.Radio, { id: 'radio4', value: "department", label: store_1.TemplateSelection.department }),
                        React.createElement(react_components_1.Radio, { id: 'radio5', value: "departmentVideo", label: store_1.TemplateSelection.departmentVideo }),
                        React.createElement(react_components_1.Radio, { id: 'radio7', value: "Default_ar", label: store_1.TemplateSelection.Default_ar }),
                        React.createElement(react_components_1.Radio, { id: 'radio10', value: "department_ar", label: store_1.TemplateSelection.department_ar }),
                        React.createElement(react_components_1.Radio, { id: 'radio11', value: "departmentVideo_ar", label: store_1.TemplateSelection.departmentVideo_ar }))),
                React.createElement("div", { className: 'card-area' },
                    React.createElement("div", { className: cardAreaBorderClass },
                        React.createElement("div", { className: 'card-area-3' })))),
            React.createElement("div", null,
                React.createElement("div", { className: 'fixed-footer' },
                    React.createElement("div", { className: 'footer-action-right' },
                        React.createElement("div", { className: 'footer-actions-flex' },
                            showMsgDraftingSpinner && (React.createElement(react_components_1.Spinner, { role: 'alert', id: 'draftingLoader', size: 'small', label: t('DraftingMessageLabel'), labelPosition: 'after' })),
                            React.createElement(react_components_1.Button, { style: { marginLeft: '16px' }, 
                                //disabled={isSaveBtnDisabled() || showMsgDraftingSpinner}
                                id: 'saveBtn', onClick: onNext, appearance: 'primary' }, t('SetAsDraft')))))))),
        pageSelection === CurrentPageSelection.CardCreation && Templates && Templates.length > 0 && (React.createElement("div", { className: "page-container" },
            React.createElement("span", { role: 'alert', "aria-label": t('NewMessageStep1') }),
            React.createElement("div", { className: 'adaptive-task-grid' },
                React.createElement("div", { className: 'form-area' },
                    React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: t('TitleText'), required: true, validationMessage: titleErrorMessage },
                        React.createElement(react_components_1.Input, { placeholder: t('PlaceHolderTitle'), onChange: onTitleChanged, autoComplete: 'off', size: 'large', required: true, appearance: 'filled-darker', value: messageState.title || '' })),
                    (selectedTemplate === store_1.TemplateSelection.department
                        || selectedTemplate === store_1.TemplateSelection.departmentVideo
                        || selectedTemplate === store_1.TemplateSelection.department_ar
                        || selectedTemplate === store_1.TemplateSelection.departmentVideo_ar)
                        && (React.createElement(React.Fragment, null,
                            React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: t('departmentText'), required: false },
                                React.createElement(react_components_1.Input, { placeholder: t('PlaceHolderDepartment'), onChange: onDeptChanged, autoComplete: 'off', size: 'large', required: false, appearance: 'filled-darker', value: messageState.department || '' })))),
                    (selectedTemplate === store_1.TemplateSelection.Default
                        || selectedTemplate === store_1.TemplateSelection.Default_ar
                        || selectedTemplate === store_1.TemplateSelection.infromational
                        || selectedTemplate === store_1.TemplateSelection.infromational_ar
                        || selectedTemplate === store_1.TemplateSelection.uae50)
                        && (React.createElement(React.Fragment, null,
                            " ",
                            React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: {
                                    children: function (_, imageInfoProps) { return (React.createElement(unstable_1.InfoLabel, __assign({}, imageInfoProps, { info: t('ImageSizeInfoContent') || '' }), t('ImageURL'))); },
                                } },
                                React.createElement("div", { style: {
                                        display: 'grid',
                                        gridTemplateColumns: '1fr auto',
                                        gridTemplateAreas: 'input-area btn-area',
                                    } },
                                    React.createElement(react_components_1.Input, { size: 'large', style: { gridColumn: '1' }, appearance: 'filled-darker', value: imageFileName || '', placeholder: t('ImageURL'), onChange: onImageLinkChanged }),
                                    React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: handleUploadClick, size: 'large', appearance: 'secondary', "aria-label": imageFileName ? t('UploadImageSuccessful') : t('UploadImageInfo'), icon: React.createElement(react_icons_2.ArrowUpload24Regular, null) }, t('Upload')),
                                    React.createElement("input", { type: 'file', accept: '.jpg, .jpeg, .png, .gif', style: { display: 'none' }, multiple: false, onChange: handleImageSelection, ref: fileInput }))))),
                    (selectedTemplate === store_1.TemplateSelection.infoVideo
                        || selectedTemplate === store_1.TemplateSelection.departmentVideo
                        || selectedTemplate === store_1.TemplateSelection.infoVideo_ar
                        || selectedTemplate === store_1.TemplateSelection.departmentVideo_ar
                        || selectedTemplate === store_1.TemplateSelection.video)
                        && (React.createElement(React.Fragment, null,
                            " ",
                            React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: {
                                    children: function (_, imageInfoProps) { return (React.createElement(unstable_1.InfoLabel, __assign({}, imageInfoProps, { info: t('PosterSizeInfoContent') || '' }), t('posterURL'))); },
                                }, validationMessage: imageUploadErrorMessage },
                                React.createElement("div", { style: {
                                        display: 'grid',
                                        gridTemplateColumns: '1fr auto',
                                        gridTemplateAreas: 'input-area btn-area',
                                    } },
                                    React.createElement(react_components_1.Input, { size: 'large', style: { gridColumn: '1' }, appearance: 'filled-darker', value: posterFileName || '', placeholder: t('PosterURL'), onChange: onPosterLinkChanged }),
                                    React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: handlePosterUploadClick, size: 'large', appearance: 'secondary', "aria-label": posterFileName ? t('UploadImageSuccessful') : t('UploadImageInfo'), icon: React.createElement(react_icons_2.ArrowUpload24Regular, null) }, t('Upload')),
                                    React.createElement("input", { type: 'file', accept: '.jpg, .jpeg, .png, .gif', style: { display: 'none' }, multiple: false, onChange: handlePosterSelection, ref: posterFileInput }))),
                            React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: {
                                    children: function (_, imageInfoProps) { return (React.createElement(unstable_1.InfoLabel, __assign({}, imageInfoProps, { info: t('VideoSizeInfoContent') || '' }), t('videoURL'))); },
                                } },
                                React.createElement("div", { style: {
                                        display: 'grid',
                                        gridTemplateColumns: '1fr auto',
                                        gridTemplateAreas: 'input-area btn-area',
                                    } },
                                    React.createElement(react_components_1.Input, { size: 'large', style: { gridColumn: '1' }, appearance: 'filled-darker', value: videoFileName || '', placeholder: t('VideoURL'), onChange: onVideoLinkChanged }))))),
                    React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: t('Summary') },
                        React.createElement(react_components_1.Textarea, { size: 'large', appearance: 'filled-darker', placeholder: t('Summary'), value: messageState.summary || '', onChange: onSummaryChanged })),
                    React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: t('Author') },
                        React.createElement(react_components_1.Input, { placeholder: t('Author'), size: 'large', onChange: onAuthorChanged, autoComplete: 'off', appearance: 'filled-darker', value: messageState.author || '' })),
                    React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: t('ButtonTitle') },
                        React.createElement(react_components_1.Input, { size: 'large', placeholder: t('ButtonTitle'), onChange: onBtnTitleChanged, autoComplete: 'off', appearance: 'filled-darker', value: messageState.buttonTitle || '' })),
                    React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: t('ButtonURL'), validationMessage: btnLinkErrorMessage },
                        React.createElement(react_components_1.Input, { size: 'large', placeholder: t('ButtonURL'), onChange: onBtnLinkChanged, type: 'url', autoComplete: 'off', appearance: 'filled-darker', value: messageState.buttonLink || '' }))),
                React.createElement("div", { className: 'card-area' },
                    React.createElement("div", { className: cardAreaBorderClass },
                        React.createElement("div", { className: 'card-area-1' })))),
            React.createElement("div", { className: 'fixed-footer' },
                React.createElement("div", { className: 'footer-action-right' },
                    React.createElement("div", { className: 'footer-actions-flex' },
                        showMsgDraftingSpinner && (React.createElement(react_components_1.Spinner, { role: 'alert', id: 'draftingLoader', size: 'small', label: t('DraftingMessageLabel'), labelPosition: 'after' })),
                        React.createElement(react_components_1.Button, { id: 'backBtn', style: { marginLeft: '16px' }, onClick: onBack, disabled: showMsgDraftingSpinner, appearance: 'secondary' }, t('Back')),
                        React.createElement(react_components_1.Button, { style: { marginLeft: '16px' }, disabled: isNextBtnDisabled() || showMsgDraftingSpinner, id: 'saveBtn', onClick: onNext, appearance: 'primary' }, t('Next'))))))),
        pageSelection === CurrentPageSelection.AudienceSelection && (React.createElement(React.Fragment, null,
            React.createElement("span", { role: 'alert', "aria-label": t('NewMessageStep2') }),
            React.createElement("div", { className: 'adaptive-task-grid' },
                React.createElement("div", { className: 'form-area' },
                    React.createElement(react_components_1.Label, { size: 'large', id: 'audienceSelectionGroupLabelId' }, t('SendHeadingText')),
                    React.createElement(react_components_1.RadioGroup, { defaultValue: selectedRadioButton, "aria-labelledby": 'audienceSelectionGroupLabelId', onChange: audienceSelectionChange },
                        React.createElement(react_components_1.Radio, { id: 'radio1', value: AudienceSelection.Teams, label: t('SendToGeneralChannel') }),
                        selectedRadioButton === AudienceSelection.Teams && (React.createElement("div", { className: cmb_styles.root },
                            React.createElement(react_components_1.Label, { id: teamsComboId }, "Pick team(s)"),
                            teamsSelectedOptions.length ? (React.createElement("ul", { id: teamsSelectedListId, className: cmb_styles.tagsList, ref: teamsSelectedListRef },
                                React.createElement("span", { id: "".concat(teamsComboId, "-remove"), hidden: true }, "Remove"),
                                teamsSelectedOptions.map(function (option, i) { return (React.createElement("li", { key: option.id },
                                    React.createElement(react_components_1.Button, { size: 'small', shape: 'rounded', appearance: 'subtle', icon: React.createElement(react_icons_2.Dismiss12Regular, null), iconPosition: 'after', onClick: function () { return onTeamsTagClick(option, i); }, id: "".concat(teamsComboId, "-remove-").concat(i), "aria-labelledby": "".concat(teamsComboId, "-remove ").concat(teamsComboId, "-remove-").concat(i) },
                                        React.createElement(react_components_1.Persona, { name: option.name, secondaryText: 'Team', avatar: { shape: 'square', color: 'colorful' } })))); }))) : (React.createElement(React.Fragment, null)),
                            React.createElement("div", { className: cmb_styles.root },
                                React.createElement("div", { className: cmb_styles.combobox },
                                    React.createElement("div", { className: cmb_styles.comboboxInputContainer },
                                        React.createElement("input", { type: "text", className: cmb_styles.comboboxInput, ref: teamsComboboxInputRef, placeholder: teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage'), onClick: function () { return handleInputClick(teamsComboboxOptionRef); }, "aria-labelledby": teamsLabelledBy }),
                                        React.createElement("div", { className: cmb_styles.comboboxIcon, onClick: function () { return handleInputClick(teamsComboboxOptionRef); } },
                                            React.createElement(react_icons_1.ChevronDownRegular, null))),
                                    React.createElement("ul", { className: "".concat(cmb_styles.comboboxOptions, " combobox-options"), ref: teamsComboboxOptionRef }, teams.map(function (opt) { return (React.createElement("li", { className: cmb_styles.comboboxOption, key: opt.id, onClick: function () { return onTeamsSelectOpt(opt); } },
                                        React.createElement(react_components_1.Persona, { name: opt.name, secondaryText: 'Team', avatar: { shape: 'square', color: 'colorful' } }))); })))))),
                        React.createElement(react_components_1.Radio, { id: 'radio2', value: AudienceSelection.Rosters, label: t('SendToRosters') }),
                        selectedRadioButton === AudienceSelection.Rosters && (React.createElement("div", { className: cmb_styles.root },
                            React.createElement(react_components_1.Label, { id: rostersComboId }, "Pick team(s)"),
                            rostersSelectedOptions.length ? (React.createElement("ul", { id: rostersSelectedListId, className: cmb_styles.tagsList, ref: rostersSelectedListRef },
                                React.createElement("span", { id: "".concat(rostersComboId, "-remove"), hidden: true }, "Remove"),
                                rostersSelectedOptions.map(function (option, i) { return (React.createElement("li", { key: option.id },
                                    React.createElement(react_components_1.Button, { size: 'small', shape: 'rounded', appearance: 'subtle', icon: React.createElement(react_icons_2.Dismiss12Regular, null), iconPosition: 'after', onClick: function () { return onRostersTagClick(option, i); }, id: "".concat(rostersComboId, "-remove-").concat(i), "aria-labelledby": "".concat(rostersComboId, "-remove ").concat(rostersComboId, "-remove-").concat(i) },
                                        React.createElement(react_components_1.Persona, { name: option.name, secondaryText: 'Team', avatar: { shape: 'square', color: 'colorful' } })))); }))) : (React.createElement(React.Fragment, null)),
                            React.createElement("div", { className: cmb_styles.root },
                                React.createElement("div", { className: cmb_styles.combobox },
                                    React.createElement("div", { className: cmb_styles.comboboxInputContainer },
                                        React.createElement("input", { type: "text", className: cmb_styles.comboboxInput, ref: teamsComboboxInputRef, placeholder: teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage'), onClick: function () { return handleInputClick(rosterComboboxOptionRef); }, "aria-labelledby": rostersLabelledBy }),
                                        React.createElement("div", { className: cmb_styles.comboboxIcon, onClick: function () { return handleInputClick(rosterComboboxOptionRef); } },
                                            React.createElement(react_icons_1.ChevronDownRegular, null))),
                                    React.createElement("ul", { className: "".concat(cmb_styles.comboboxOptions, " combobox-options"), ref: rosterComboboxOptionRef }, teams.map(function (opt) { return (React.createElement("li", { className: cmb_styles.comboboxOption, key: opt.id, onClick: function () { return onRosterSelectOpt(opt); } },
                                        React.createElement(react_components_1.Persona, { name: opt.name, secondaryText: 'Team', avatar: { shape: 'square', color: 'colorful' } }))); })))))),
                        React.createElement(react_components_1.Radio, { id: 'radio3', value: AudienceSelection.AllUsers, label: t('SendToAllUsers') }),
                        React.createElement("div", { className: cmb_styles.root }, selectedRadioButton === AudienceSelection.AllUsers && (React.createElement(react_components_1.Text, { id: 'radio3Note', role: allUsersAria, className: 'info-text' }, t('SendToAllUsersNote')))),
                        React.createElement(react_components_1.Radio, { id: 'radio4', value: AudienceSelection.Groups, label: t('SendToGroups') }),
                        selectedRadioButton === AudienceSelection.Groups && (React.createElement("div", { className: cmb_styles.root },
                            !canAccessGroups && (React.createElement(react_components_1.Text, { role: groupsAria, className: 'info-text' }, t('SendToGroupsPermissionNote'))),
                            canAccessGroups && (React.createElement(React.Fragment, null,
                                React.createElement(react_components_1.Label, { id: searchComboId }, "Pick group(s)"),
                                searchSelectedOptions.length ? (React.createElement("ul", { id: searchSelectedListId, className: cmb_styles.tagsList, ref: searchSelectedListRef },
                                    React.createElement("span", { id: "".concat(searchComboId, "-remove"), hidden: true }, "Remove"),
                                    searchSelectedOptions.map(function (option, i) { return (React.createElement("li", { key: option.id },
                                        React.createElement(react_components_1.Button, { size: 'small', shape: 'rounded', appearance: 'subtle', icon: React.createElement(react_icons_2.Dismiss12Regular, null), iconPosition: 'after', onClick: function () { return onSearchTagClick(option, i); }, id: "".concat(searchComboId, "-remove-").concat(i), "aria-labelledby": "".concat(searchComboId, "-remove ").concat(searchComboId, "-remove-").concat(i) },
                                            React.createElement(react_components_1.Persona, { name: option.name, secondaryText: 'Group', avatar: { color: 'colorful' } })))); }))) : (React.createElement(React.Fragment, null)),
                                React.createElement("div", { className: cmb_styles.root },
                                    React.createElement("div", { className: cmb_styles.combobox },
                                        React.createElement("div", { className: cmb_styles.comboboxInputContainer },
                                            React.createElement("input", { type: "text", className: cmb_styles.comboboxInput, ref: teamsComboboxInputRef, placeholder: teams.length !== 0 ? 'Pick one or more teams' : t('NoMatchMessage'), onClick: function () { return handleInputClick(searchComboboxOptionRef); }, "aria-labelledby": searchLabelledBy }),
                                            React.createElement("div", { className: cmb_styles.comboboxIcon, onClick: function () { return handleInputClick(searchComboboxOptionRef); } },
                                                React.createElement(react_icons_1.ChevronDownRegular, null))),
                                        React.createElement("ul", { className: "".concat(cmb_styles.comboboxOptions, " combobox-options"), ref: searchComboboxOptionRef }, queryGroups.map(function (opt) { return (React.createElement("li", { className: cmb_styles.comboboxOption, key: opt.id, onClick: function () { return onSearchSelectOpt(opt); } },
                                            React.createElement(react_components_1.Persona, { name: opt.name, secondaryText: 'Team', avatar: { shape: 'square', color: 'colorful' } }))); })))),
                                React.createElement(react_components_1.Text, { role: groupsAria, className: 'info-text' }, t('SendToGroupsNote')))))))),
                React.createElement("div", { className: 'card-area' },
                    React.createElement("div", { className: cardAreaBorderClass },
                        React.createElement("div", { className: 'card-area-2' })))),
            React.createElement("div", null,
                React.createElement("div", { className: 'fixed-footer' },
                    React.createElement("div", { className: 'footer-action-right' },
                        React.createElement("div", { className: 'footer-actions-flex' },
                            showMsgDraftingSpinner && (React.createElement(react_components_1.Spinner, { role: 'alert', id: 'draftingLoader', size: 'small', label: t('DraftingMessageLabel'), labelPosition: 'after' })),
                            React.createElement(react_components_1.Button, { id: 'backBtn', style: { marginLeft: '16px' }, onClick: onBack, disabled: showMsgDraftingSpinner, appearance: 'secondary' }, t('Back')),
                            React.createElement(react_components_1.Button, { style: { marginLeft: '16px' }, disabled: isSaveBtnDisabled() || showMsgDraftingSpinner, id: 'saveBtn', onClick: onSave, appearance: 'primary' }, t('SaveAsDraft'))))))))));
};
exports.NewMessage = NewMessage;
//# sourceMappingURL=newMessage.js.map