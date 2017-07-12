import { INewsItem } from "../interfaces";

export interface IListItemService {
    getNewsItems(): Promise<INewsItem[]>;
}