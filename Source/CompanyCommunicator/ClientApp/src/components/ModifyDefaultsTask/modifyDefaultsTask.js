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
Object.defineProperty(exports, "__esModule", { value: true });
exports.ModifyDefaultsTask = void 0;
var AdaptiveCards = require("adaptivecards");
var React = require("react");
var react_i18next_1 = require("react-i18next");
var react_components_1 = require("@fluentui/react-components");
var unstable_1 = require("@fluentui/react-components/unstable");
var react_icons_1 = require("@fluentui/react-icons");
var microsoftTeams = require("@microsoft/teams-js");
var messageListApi_1 = require("../../apis/messageListApi");
var adaptiveCard_1 = require("../AdaptiveCard/adaptiveCard");
var validImageTypes = ['image/gif', 'image/jpeg', 'image/png', 'image/jpg'];
var useFieldStyles = (0, react_components_1.makeStyles)({
    styles: {
        marginBottom: react_components_1.tokens.spacingVerticalM,
        gridGap: react_components_1.tokens.spacingHorizontalXXS,
    },
    defaultsType: {
        gridGap: react_components_1.tokens.spacingHorizontalXXS,
        marginTop: react_components_1.tokens.spacingVerticalM,
        marginLeft: react_components_1.tokens.spacingHorizontalS,
    }
});
var card;
var ModifyDefaultsTask = function () {
    var t = (0, react_i18next_1.useTranslation)().t;
    var logoFileInput = React.createRef();
    var headerLogoFileInput = React.createRef();
    var bannerFileInput = React.createRef();
    var field_styles = useFieldStyles();
    var _a = React.useState(''), logoFileName = _a[0], setLogoFileName = _a[1];
    var _b = React.useState(''), headerLogoFileName = _b[0], setHeaderLogoFileName = _b[1];
    var _c = React.useState(''), bannerFileName = _c[0], setBannerFileName = _c[1];
    var _d = React.useState(null), headerLogoImagePath = _d[0], setHeaderLogoImagePath = _d[1];
    var _e = React.useState(), headerText = _e[0], setHeaderText = _e[1];
    var _f = React.useState(''), imageUploadErrorMessage = _f[0], setImageUploadErrorMessage = _f[1];
    var _g = React.useState('card-area-border'), cardAreaBorderClass = _g[0], setCardAreaBorderClass = _g[1];
    var _h = React.useState({
        logoFileName: "",
        logoLink: "",
        bannerLink: "",
        bannerFileName: "",
        headerLogoLink: "",
        headerText: "",
        headerLogoFileName: ""
    }), defaultsState = _h[0], setDefaultState = _h[1];
    React.useEffect(function () {
        getDefaultsItem();
    }, []);
    var getDefaultsItem = function () { return __awaiter(void 0, void 0, void 0, function () {
        var error_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, (0, messageListApi_1.getDefaultData)().then(function (response) {
                            var defaultImages = response.data;
                            console.log(defaultImages);
                            setDefaultState({
                                logoFileName: defaultImages.logoFileName,
                                logoLink: defaultImages.logoLink,
                                bannerFileName: defaultImages.bannerFileName,
                                bannerLink: defaultImages.bannerLink,
                                headerLogoLink: defaultImages.headerLogoLink,
                                headerLogoFileName: defaultImages.headerLogoFileName,
                                headerText: defaultImages.headerText
                            });
                            setHeaderLogoImagePath(defaultImages.headerLogoLink);
                            card = (0, adaptiveCard_1.getInitAdaptiveCard)("title", "viewDefaults");
                            setHeaderText(defaultImages.headerText);
                            (0, adaptiveCard_1.setCardLogo)(card, defaultImages.logoLink);
                            (0, adaptiveCard_1.setCardBanner)(card, defaultImages.bannerLink);
                            updateAdaptiveCard();
                        })];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    error_1 = _a.sent();
                    card = (0, adaptiveCard_1.getInitAdaptiveCard)("title", "viewDefaults");
                    updateAdaptiveCard();
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); };
    var updateAdaptiveCard = function () {
        var adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        var renderCard = adaptiveCard.render();
        document.getElementsByClassName('card-area-1')[0].innerHTML = '';
        document.getElementsByClassName('card-area-1')[0].appendChild(renderCard);
        setCardAreaBorderClass('card-area-border');
        adaptiveCard.onExecuteAction = function (action) {
            window.open(action.url, '_blank');
        };
    };
    var onLogoLinkChanged = function (event) {
        var urlOrDataUrl = event.target.value;
        var isGoodLink = true;
        setLogoFileName(urlOrDataUrl);
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
            (0, adaptiveCard_1.setCardLogo)(card, event.target.value);
            setDefaultState(__assign(__assign({}, defaultsState), { logoLink: urlOrDataUrl }));
            updateAdaptiveCard();
        }
    };
    var onLogoLinkDelete = function (event) {
        (0, adaptiveCard_1.setCardLogo)(card, "");
        setLogoFileName('');
        updateAdaptiveCard();
        setDefaultState(__assign(__assign({}, defaultsState), { logoLink: 'DELETE' }));
    };
    var handleLogoUploadClick = function (event) {
        if (logoFileInput.current) {
            logoFileInput.current.click();
        }
    };
    var handleLogoSelection = function () {
        var _a;
        var file = (_a = logoFileInput.current) === null || _a === void 0 ? void 0 : _a.files[0];
        imageselection(file, "logo");
    };
    var onBannerLinkChanged = function (event) {
        var urlOrDataUrl = event.target.value;
        var isGoodLink = true;
        setBannerFileName(urlOrDataUrl);
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
            (0, adaptiveCard_1.setCardBanner)(card, event.target.value);
            setDefaultState(__assign(__assign({}, defaultsState), { bannerLink: urlOrDataUrl }));
            updateAdaptiveCard();
        }
    };
    var onBannerLinkDelete = function (event) {
        (0, adaptiveCard_1.setCardBanner)(card, "");
        setBannerFileName('');
        updateAdaptiveCard();
        setDefaultState(__assign(__assign({}, defaultsState), { bannerLink: 'DELETE' }));
    };
    var handleBannerUploadClick = function (event) {
        if (bannerFileInput.current) {
            bannerFileInput.current.click();
        }
    };
    var handleBannerSelection = function () {
        var _a;
        var file = (_a = bannerFileInput.current) === null || _a === void 0 ? void 0 : _a.files[0];
        imageselection(file, "banner");
    };
    var onHeaderLogoLinkChanged = function (event) {
        var urlOrDataUrl = event.target.value;
        var isGoodLink = true;
        setHeaderLogoFileName(urlOrDataUrl);
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
        }
    };
    var onHeaderLogoLinkDelete = function (event) {
        setHeaderLogoFileName('');
        setHeaderLogoImagePath(null);
        setDefaultState(__assign(__assign({}, defaultsState), { headerLogoLink: 'DELETE' }));
    };
    var handleHeaderLogoUploadClick = function (event) {
        if (headerLogoFileInput.current) {
            headerLogoFileInput.current.click();
        }
    };
    var handleHeaderLogoSelection = function () {
        var _a;
        var file = (_a = headerLogoFileInput.current) === null || _a === void 0 ? void 0 : _a.files[0];
        imageselection(file, "headerlogo");
    };
    var onHeaderTextChanged = function (event) {
        setHeaderText(event.target.value);
        setDefaultState(__assign(__assign({}, defaultsState), { headerText: event.target.value }));
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
                if (resizedImageAsBase64 && field === 'logo') {
                    setLogoFileName(file['name']);
                    setImageUploadErrorMessage('');
                    (0, adaptiveCard_1.setCardLogo)(card, resizedImageAsBase64);
                    setDefaultState(__assign(__assign({}, defaultsState), { logoLink: resizedImageAsBase64 }));
                    updateAdaptiveCard();
                }
                else if (resizedImageAsBase64 && field === 'banner') {
                    setBannerFileName(file['name']);
                    setImageUploadErrorMessage('');
                    (0, adaptiveCard_1.setCardBanner)(card, resizedImageAsBase64);
                    setDefaultState(__assign(__assign({}, defaultsState), { bannerLink: resizedImageAsBase64 }));
                    updateAdaptiveCard();
                }
                else if (resizedImageAsBase64 && field === "headerlogo") {
                    setHeaderLogoFileName(file['name']);
                    setHeaderLogoImagePath(resizedImageAsBase64);
                    setDefaultState(__assign(__assign({}, defaultsState), { headerLogoLink: resizedImageAsBase64 }));
                }
            };
        }
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
    var onSave = function () {
        try {
            (0, messageListApi_1.updateDefaultData)(defaultsState)
                .then(function () {
            })
                .finally(function () {
                microsoftTeams.tasks.submitTask();
            });
        }
        catch (error) {
            return error;
        }
    };
    return (React.createElement(React.Fragment, null,
        React.createElement(react_components_1.Field, { size: 'large', className: field_styles.defaultsType, label: t("ModifyAppDefaults") }),
        React.createElement("div", { className: 'adaptive-task-grid-app-defaults ms-motion-slideLeftIn' },
            React.createElement("div", { className: 'form-area' },
                React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: {
                        children: function (_, imageInfoProps) { return (React.createElement(unstable_1.InfoLabel, __assign({}, imageInfoProps, { info: t('ImageSizeInfoContent') || '' }), t('HeaderLogoURL'))); },
                    } },
                    React.createElement("div", { style: {
                            display: 'grid',
                            gridTemplateColumns: '1fr auto',
                            gridTemplateAreas: 'input-area btn-area',
                        } },
                        React.createElement(react_components_1.Input, { size: 'large', style: { gridColumn: '1' }, appearance: 'filled-darker', value: headerLogoFileName || '', placeholder: t('HeaderLogoUrl'), onChange: onHeaderLogoLinkChanged }),
                        React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: handleHeaderLogoUploadClick, size: 'large', appearance: 'secondary', "aria-label": headerLogoFileInput ? t('UploadImageSuccessful') : t('UploadImageInfo'), icon: React.createElement(react_icons_1.ArrowUpload24Regular, null) }, t('Upload')),
                        React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: onHeaderLogoLinkDelete, size: 'large', appearance: 'secondary', "aria-label": headerLogoFileInput ? t('DeleteImageSuccessful') : t('DeleteImageInfo'), icon: React.createElement(react_icons_1.Delete24Regular, null), disabled: defaultsState.headerLogoLink === null || defaultsState.headerLogoLink === '' || defaultsState.headerLogoLink === 'DELETE' }, t('Delete')),
                        React.createElement("input", { type: 'file', accept: '.jpg, .jpeg, .png, .gif', style: { display: 'none' }, multiple: false, onChange: handleHeaderLogoSelection, ref: headerLogoFileInput }))),
                React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: t('TitleText') },
                    React.createElement(react_components_1.Input, { placeholder: t('HeaderTextPlaceholder'), onChange: onHeaderTextChanged, autoComplete: 'off', size: 'large', required: true, appearance: 'filled-darker', value: headerText || '' }))),
            React.createElement("div", { className: 'card-area' },
                React.createElement("div", { className: cardAreaBorderClass },
                    React.createElement("div", { className: 'card-area-2' },
                        React.createElement("div", { className: "cc-main-left" },
                            headerLogoImagePath != null && React.createElement("img", { src: headerLogoImagePath, alt: "logo", className: "cc-logo", title: headerText }),
                            React.createElement("span", { className: "cc-header-text", title: headerText }, headerText)))))),
        React.createElement(react_components_1.Divider, null),
        React.createElement(react_components_1.Field, { size: 'large', className: field_styles.defaultsType, label: t("ModifyCardDefaults") }),
        React.createElement("div", { className: 'adaptive-task-grid ms-motion-slideLeftIn' },
            React.createElement("div", { className: 'form-area' },
                React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: {
                        children: function (_, imageInfoProps) { return (React.createElement(unstable_1.InfoLabel, __assign({}, imageInfoProps, { info: t('ImageSizeInfoContent') || '' }), t('LogoURL'))); },
                    } },
                    React.createElement("div", { style: {
                            display: 'grid',
                            gridTemplateColumns: '1fr auto',
                            gridTemplateAreas: 'input-area btn-area',
                        } },
                        React.createElement(react_components_1.Input, { size: 'large', style: { gridColumn: '1' }, appearance: 'filled-darker', value: logoFileName || '', placeholder: t('LogoUrl'), onChange: onLogoLinkChanged }),
                        React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: handleLogoUploadClick, size: 'large', appearance: 'secondary', "aria-label": logoFileName ? t('UploadImageSuccessful') : t('UploadImageInfo'), icon: React.createElement(react_icons_1.ArrowUpload24Regular, null) }, t('Upload')),
                        React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: onLogoLinkDelete, size: 'large', appearance: 'secondary', disabled: defaultsState.logoLink === null || defaultsState.logoLink === '' || defaultsState.logoLink === 'DELETE', "aria-label": logoFileName ? t('DeleteImageSuccessful') : t('DeleteImageInfo'), icon: React.createElement(react_icons_1.Delete24Regular, null) }, t('Delete')),
                        React.createElement("input", { type: 'file', accept: '.jpg, .jpeg, .png, .gif', style: { display: 'none' }, multiple: false, onChange: handleLogoSelection, ref: logoFileInput }))),
                React.createElement(react_components_1.Field, { size: 'large', className: field_styles.styles, label: {
                        children: function (_, imageInfoProps) { return (React.createElement(unstable_1.InfoLabel, __assign({}, imageInfoProps, { info: t('ImageSizeInfoContent') || '' }), t('BannerURL'))); },
                    } },
                    React.createElement("div", { style: {
                            display: 'grid',
                            gridTemplateColumns: '1fr auto',
                            gridTemplateAreas: 'input-area btn-area',
                        } },
                        React.createElement(react_components_1.Input, { size: 'large', style: { gridColumn: '1' }, appearance: 'filled-darker', value: bannerFileName || '', placeholder: t('BannerUrl'), onChange: onBannerLinkChanged }),
                        React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: handleBannerUploadClick, size: 'large', appearance: 'secondary', "aria-label": bannerFileName ? t('UploadImageSuccessful') : t('UploadImageInfo'), icon: React.createElement(react_icons_1.ArrowUpload24Regular, null) }, t('Upload')),
                        React.createElement(react_components_1.Button, { style: { gridColumn: '2', marginLeft: '5px' }, onClick: onBannerLinkDelete, size: 'large', appearance: 'secondary', "aria-label": bannerFileName ? t('DeleteImageSuccessful') : t('DeleteImageInfo'), disabled: defaultsState.bannerLink === null || defaultsState.bannerLink === '' || defaultsState.bannerLink === 'DELETE', icon: React.createElement(react_icons_1.Delete24Regular, null) }, t('Delete')),
                        React.createElement("input", { type: 'file', accept: '.jpg, .jpeg, .png, .gif', style: { display: 'none' }, multiple: false, onChange: handleBannerSelection, ref: bannerFileInput })))),
            React.createElement("div", { className: 'card-area' },
                React.createElement("div", { className: cardAreaBorderClass },
                    React.createElement("div", { className: 'card-area-1' })))),
        React.createElement("div", { className: 'fixed-footer' },
            React.createElement("div", { className: 'footer-action-right' },
                React.createElement("div", { className: 'footer-actions-flex' },
                    React.createElement(react_components_1.Button, { style: { marginLeft: '16px' }, disabled: false, id: 'saveBtn', onClick: onSave, appearance: 'primary' }, t('save')))))));
};
exports.ModifyDefaultsTask = ModifyDefaultsTask;
//# sourceMappingURL=modifyDefaultsTask.js.map