import { IListItemService } from "./IListItemService";
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { INewsItem } from "../interfaces";
import pnp, { List } from "sp-pnp-js";

export class ListItemService implements IListItemService {
    public static readonly serviceKey: ServiceKey<IListItemService> = ServiceKey.create<IListItemService>('ast:IListItemService', ListItemService);
    private _spHttpClient: SPHttpClient;
    private _pageContext: PageContext;
    private _currentWebUrl: string;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() =>{
            this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
            this._currentWebUrl = this._pageContext.web.absoluteUrl;

            pnp.setup({
                baseUrl: this._currentWebUrl
            });
        });
    }

    public getNewsItems(): Promise<INewsItem[]> {
        return pnp.sp.web.lists.getByTitle("News")
                .items
                .select("Id","Title")
                .usingCaching()
                .get()
                .then((newsItems: INewsItem[]) => {
                    return newsItems;
                });
    }
}