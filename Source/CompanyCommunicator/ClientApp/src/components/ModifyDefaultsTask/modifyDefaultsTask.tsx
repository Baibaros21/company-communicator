import * as AdaptiveCards from 'adaptivecards';
import * as React from 'react';
import { useTranslation } from 'react-i18next';
import { Button, Field, Input, LabelProps, makeStyles, tokens, Divider } from '@fluentui/react-components';
import { InfoLabel } from '@fluentui/react-components/unstable';
import { ArrowUpload24Regular } from '@fluentui/react-icons';
import * as microsoftTeams from '@microsoft/teams-js';
import { getDefaultData, updateDefaultData } from '../../apis/messageListApi';
import { getInitAdaptiveCard, setCardBanner, setCardLogo } from '../AdaptiveCard/adaptiveCard';
import mslogo from "../../assets/Images/mslogo.png";

const validImageTypes = ['image/gif', 'image/jpeg', 'image/png', 'image/jpg'];

const useFieldStyles = makeStyles({
    styles: {
        marginBottom: tokens.spacingVerticalM,
        gridGap: tokens.spacingHorizontalXXS,
    },
    defaultsType: {
        gridGap: tokens.spacingHorizontalXXS,
        marginTop: tokens.spacingVerticalM,
        marginLeft: tokens.spacingHorizontalS,
    }
});

interface IDefaults {

    logoFileName: string;
    logoLink: string;
    bannerFileName: string;
    bannerLink: string;
    headerLogoLink: string;
}


let card: any;

export const ModifyDefaultsTask = () => {

    const { t } = useTranslation();
    let logoFileInput = React.createRef<any>();
    let headerLogoFileInput = React.createRef<any>();
    let bannerFileInput = React.createRef<any>();

    const field_styles = useFieldStyles();

    const [logoFileName, setLogoFileName] = React.useState('');
    const [headerLogoFileName, setHeaderLogoFileName] = React.useState('');
    const [bannerFileName, setBannerFileName] = React.useState('');
    const [headerLogoImagePath, setHeaderLogoImagePath] = React.useState<any>(mslogo);
    const [headerText, setHeaderText] = React.useState<any>();
    const [imageUploadErrorMessage, setImageUploadErrorMessage] = React.useState('');
    const [cardAreaBorderClass, setCardAreaBorderClass] = React.useState('card-area-border');
    const [defaultsState, setDefaultState] = React.useState<IDefaults>({
        logoFileName: "",
        logoLink: "",
        bannerLink: "",
        bannerFileName: "",
        headerLogoLink: ""
    });

    React.useEffect(() => {
        getDefaultsItem();

    }, []);

    const getDefaultsItem = async () => {

        try {
            await getDefaultData().then((response) => {

                const defaultImages = response.data;
                console.log(defaultImages);
                setDefaultState({
                    logoFileName: defaultImages.logoFileName,
                    logoLink: defaultImages.logoLink,
                    bannerFileName: defaultImages.bannerFileName,
                    bannerLink: defaultImages.bannerLink,
                    headerLogoLink: defaultImages.headerLogoLink
                });

                setHeaderLogoImagePath(defaultImages.headerLogoLink);
                card = getInitAdaptiveCard("title", "viewDefaults");
                setCardLogo(card, defaultImages.logoLink);
                setCardBanner(card, defaultImages.bannerLink);
                updateAdaptiveCard();

            });
        }
        catch (error) {
            card = getInitAdaptiveCard("title", "viewDefaults");
            updateAdaptiveCard();
        }
    }

    const onLogoLinkChanged = (event: any) => {
        const urlOrDataUrl = event.target.value;
        let isGoodLink = true;
        setLogoFileName(urlOrDataUrl);

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
            setCardLogo(card, event.target.value);
            setDefaultState({ ...defaultsState, logoLink: urlOrDataUrl });
            updateAdaptiveCard();
        }
    };


    const onBannerLinkChanged = (event: any) => {
        const urlOrDataUrl = event.target.value;
        let isGoodLink = true;
        setBannerFileName(urlOrDataUrl);

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

            setCardBanner(card, event.target.value);
            setDefaultState({ ...defaultsState, bannerLink: urlOrDataUrl });
            updateAdaptiveCard();
        }
    };

    const onHeaderLogoLinkChanged = (event: any) => {
        const urlOrDataUrl = event.target.value;
        let isGoodLink = true;
        setHeaderLogoFileName(urlOrDataUrl);

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

        }
    };

    const handleHeaderLogoUploadClick = (event: any) => {
        if (headerLogoFileInput.current) {
            headerLogoFileInput.current.click();
        }
    };

    const handleHeaderLogoSelection = () => {
        const file = headerLogoFileInput.current?.files[0];

        imageselection(file, "headerlogo");


    };


    const handleLogoUploadClick = (event: any) => {
        if (logoFileInput.current) {
            logoFileInput.current.click();
        }
    };

    const handleLogoSelection = () => {
        const file = logoFileInput.current?.files[0];

        imageselection(file, "logo");


    };



    const handleBannerUploadClick = (event: any) => {
        if (bannerFileInput.current) {
            bannerFileInput.current.click();
        }
    };

    const handleBannerSelection = () => {
        const file = bannerFileInput.current?.files[0];

        imageselection(file, "banner");


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

                if (resizedImageAsBase64 && field === 'logo') {
                    setLogoFileName(file['name']);
                    setImageUploadErrorMessage('');
                    setCardLogo(card, resizedImageAsBase64);
                    setDefaultState({ ...defaultsState, logoLink: resizedImageAsBase64 });
                    updateAdaptiveCard();
                } else if (resizedImageAsBase64 && field === 'banner') {
                    setBannerFileName(file['name']);
                    setImageUploadErrorMessage('');
                    setCardBanner(card, resizedImageAsBase64);
                    setDefaultState({ ...defaultsState, bannerLink: resizedImageAsBase64 });
                    updateAdaptiveCard();
                } else if (resizedImageAsBase64 && field === "headerlogo") {
                    setHeaderLogoFileName(file['name']);
                    setHeaderLogoImagePath(resizedImageAsBase64);
                    setDefaultState({ ...defaultsState, headerLogoLink: resizedImageAsBase64 });
                }



            };
        }
    }

    const updateAdaptiveCard = () => {
        var adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(card);
        const renderCard = adaptiveCard.render();

        document.getElementsByClassName('card-area-1')[0].innerHTML = '';
        document.getElementsByClassName('card-area-1')[0].appendChild(renderCard);
        setCardAreaBorderClass('card-area-border');
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

    const onHeaderTextChanged = (event: any) => {
        setHeaderText(event.target.value);
    }

    const onSave = () => {

        process.env.REACT_APP_HEADERIMAGE = headerLogoImagePath;
        process.env.REACT_APP_HEADERTEXT = headerText;
        console.log(headerLogoImagePath);
        console.log(process.env.REACT_APP_HEADERIMAGE);
        try {
            updateDefaultData(defaultsState)
                .then(() => {


                })
                .finally(() => {

                    microsoftTeams.tasks.submitTask();
                });
        } catch (error) {
            return error;
        }

    };

    React.useEffect(() => {
        setHeaderLogoImagePath(process.env.REACT_APP_HEADERIMAGE);
        setHeaderText(process.env.REACT_APP_HEADERTEXT
            ? t(process.env.REACT_APP_HEADERTEXT)
            : t("Mersal"));
    }, []);

    return (
        <>
            <Field
                size='large'
                className={field_styles.defaultsType}
                label={
                    t("ModifyAppDefaults")
                }></Field>

            <div className='adaptive-task-grid-app-defaults ms-motion-slideLeftIn'>
                <div className='form-area'>
                    <Field
                        size='large'
                        className={field_styles.styles}
                        label={{
                            children: (_: unknown, imageInfoProps: LabelProps) => (
                                <InfoLabel {...imageInfoProps} info={t('ImageSizeInfoContent') || ''}>
                                    {t('HeaderLogoURL')}
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
                                value={headerLogoFileName || ''}
                                placeholder={t('HeaderLogoUrl')}
                                onChange={onHeaderLogoLinkChanged}
                            />
                            <Button
                                style={{ gridColumn: '2', marginLeft: '5px' }}
                                onClick={handleHeaderLogoUploadClick}
                                size='large'
                                appearance='secondary'
                                aria-label={headerLogoFileInput ? t('UploadImageSuccessful') : t('UploadImageInfo')}
                                icon={<ArrowUpload24Regular />}
                            >
                                {t('Upload')}
                            </Button>
                            <input
                                type='file'
                                accept='.jpg, .jpeg, .png, .gif'
                                style={{ display: 'none' }}
                                multiple={false}
                                onChange={handleHeaderLogoSelection}
                                ref={headerLogoFileInput}
                            />
                        </div>
                    </Field>
                    <Field size='large' className={field_styles.styles} label={t('TitleText')} >
                        <Input
                            placeholder={t('HeaderTextPlaceholder')}
                            onChange={onHeaderTextChanged}
                            autoComplete='off'
                            size='large'
                            required={true}
                            appearance='filled-darker'
                            value={headerText || ''}
                        />
                    </Field>
                </div>

                <div className='card-area'>
                    <div className={cardAreaBorderClass}>
                        <div className='card-area-2'>
                            <div className="cc-main-left">
                                <img
                                    src={headerLogoImagePath}
                                    alt="logo"
                                    className="cc-logo"
                                    title={headerText}
                                />
                                <span className="cc-header-text" title={headerText}>
                                    {headerText}
                                </span>
                            </div>
                        </div>
                    </div>
                </div>

            </div>
            <Divider />
            <Field
                size='large'
                className={field_styles.defaultsType}
                label={
                    t("ModifyCardDefaults")
                }></Field>
            <div className='adaptive-task-grid ms-motion-slideLeftIn'>
                <div className='form-area'>
                    <Field
                        size='large'
                        className={field_styles.styles}
                        label={{
                            children: (_: unknown, imageInfoProps: LabelProps) => (
                                <InfoLabel {...imageInfoProps} info={t('ImageSizeInfoContent') || ''}>
                                    {t('LogoURL')}
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
                                value={logoFileName || ''}
                                placeholder={t('LogoUrl')}
                                onChange={onLogoLinkChanged}
                            />
                            <Button
                                style={{ gridColumn: '2', marginLeft: '5px' }}
                                onClick={handleLogoUploadClick}
                                size='large'
                                appearance='secondary'
                                aria-label={logoFileName ? t('UploadImageSuccessful') : t('UploadImageInfo')}
                                icon={<ArrowUpload24Regular />}
                            >
                                {t('Upload')}
                            </Button>
                            <input
                                type='file'
                                accept='.jpg, .jpeg, .png, .gif'
                                style={{ display: 'none' }}
                                multiple={false}
                                onChange={handleLogoSelection}
                                ref={logoFileInput}
                            />
                        </div>
                    </Field>
                    <Field
                        size='large'
                        className={field_styles.styles}
                        label={{
                            children: (_: unknown, imageInfoProps: LabelProps) => (
                                <InfoLabel {...imageInfoProps} info={t('ImageSizeInfoContent') || ''}>
                                    {t('BannerURL')}
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
                                value={bannerFileName || ''}
                                placeholder={t('BannerUrl')}
                                onChange={onBannerLinkChanged}
                            />
                            <Button
                                style={{ gridColumn: '2', marginLeft: '5px' }}
                                onClick={handleBannerUploadClick}
                                size='large'
                                appearance='secondary'
                                aria-label={bannerFileName ? t('UploadImageSuccessful') : t('UploadImageInfo')}
                                icon={<ArrowUpload24Regular />}
                            >
                                {t('Upload')}
                            </Button>
                            <input
                                type='file'
                                accept='.jpg, .jpeg, .png, .gif'
                                style={{ display: 'none' }}
                                multiple={false}
                                onChange={handleBannerSelection}
                                ref={bannerFileInput}
                            />
                        </div>
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
                        <Button
                            style={{ marginLeft: '16px' }}
                            disabled={false}
                            id='saveBtn'
                            onClick={onSave}
                            appearance='primary'
                        >
                            {t('save')}
                        </Button>
                    </div>
                </div>
            </div>

        </>
    );
};