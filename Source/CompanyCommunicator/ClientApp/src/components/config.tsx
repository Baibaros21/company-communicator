import * as React from "react";
import * as microsoftTeams from '@microsoft/teams-js';
import { getBaseUrl } from "../configVariables";

const Configuration: React.FC = () => {
    const [url, setUrl] = React.useState(getBaseUrl() + "/messages?locale={locale}");

    React.useEffect(() => {
        microsoftTeams.initialize();
        microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
            microsoftTeams.settings.setSettings({
                entityId: "Company_Communicator_App",
                contentUrl: url,
                suggestedDisplayName: "ComCast",
            });
            saveEvent.notifySuccess();
        });

        microsoftTeams.settings.setValidityState(true);
    }, [url]);

    return (
        <div className="configContainer">
            <h3>Please click Save to get started.</h3>
        </div>
    );
};

export default Configuration;
