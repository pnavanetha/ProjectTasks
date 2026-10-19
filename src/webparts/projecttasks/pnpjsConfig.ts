import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";

import { LogLevel, PnPLogging } from "@pnp/logging";

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

let _sp: SPFI | null = null;

export const getSP = (context?: WebPartContext): SPFI => {
    if (_sp === null && context) {
        _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _sp as SPFI;
};
