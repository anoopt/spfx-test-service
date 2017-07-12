import { INewsItem } from "../interfaces";
export interface ISpfxserviceState {
    items: INewsItem[];
    errors: string[];
    status: string;
}