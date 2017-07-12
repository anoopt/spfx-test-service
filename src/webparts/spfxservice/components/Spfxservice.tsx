import * as React from 'react';
import styles from './Spfxservice.module.scss';
import { INewsItem } from "../interfaces";
import { Logger, FunctionListener, LogEntry, LogLevel, Web } from "sp-pnp-js";
import { Log } from "@microsoft/sp-core-library";
import { ServiceScope, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { ISpfxserviceProps } from './ISpfxserviceProps';
import { ISpfxserviceState } from './ISpfxserviceState';
import { IListItemService, ListItemService } from '../services';

import { escape } from '@microsoft/sp-lodash-subset';

export default class Spfxservice extends React.Component<ISpfxserviceProps, ISpfxserviceState> {

  private _listItemServiceInstance: IListItemService;

  constructor(props: ISpfxserviceProps){
    super(props);
    this.state = {
      items: [],
      errors: [],
      status: "Ready"
    };

    let serviceScope: ServiceScope;
    serviceScope = this.props.serviceScope;

    this._listItemServiceInstance = serviceScope.consume(ListItemService.serviceKey);
    this._readItems.bind(this);
    this._enableLogging();
  }

  public componentDidMount(): void {
    this.setState({
        items: [],
        errors: [],
        status: "Loading"
      });
    this._readItems("News");
  }

  public render(): React.ReactElement<ISpfxserviceProps> {
    return (
      <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">SPFx Services using Vrdmn and Jose's code</span>
              <div>
                {this._gerErrors()}
                {this.state.status}
              </div>
              <p className="ms-font-l ms-fontColor-white">News Items</p>
              <div>
                <div className={styles.row}>
                  <div className={styles.left}>Id</div>
                  <div className={styles.right}>Title</div>
                </div>
                {
                  this.state.items.map((item, idx) => {
                    return(
                      <div className={styles.row}>
                        <div className={styles.left}>{item.Id}</div>
                        <div className={styles.right}>{escape(item.Title)}</div>
                      </div>
                    );
                  })
                }
              </div>
            </div>
          </div>
        </div>
    );
  }

  private _enableLogging(): void {
    ////////////////////////////////////////////////////////////////////////
    // enable Logging system
    ////////////////////////////////////////////////////////////////////////
    // we will integrate PnP JS Logging System with SPFx Logging system
    // 1. Logger object => PnP JS Logger
    //    https://github.com/SharePoint/PnP-JS-Core/wiki/Working-With:-Logging
    // 2. Log object => SPFx Logger
    //    https://github.com/SharePoint/sp-dev-docs/wiki/Working-with-the-Logging-API
    ////////////////////////////////////////////////////////////////////////
    // [PnP JS Logging] activate Info level
    Logger.activeLogLevel = LogLevel.Info;

    // [PnP JS Logging] create a custom FunctionListener to integrate PnP JS and SPFx Logging systems
    const listener = new FunctionListener((entry: LogEntry) => {

      // get React component name
      const componentName: string = (this as any)._reactInternalInstance._currentElement.props.description;

      // mapping betwween PnP JS Log types and SPFx logging methods
      // instead of using switch we use object easy syntax
      const logLevelConversion = { Verbose: "verbose", Info: "info", Warning: "warn", Error: "error" };

      // create Message. Two importante notes here:
      // 1. Use JSON.stringify to output everything. It´s helpful when some internal exception comes thru.
      // 2. Use JavaScript´s Error constructor allows us to output more than 100 characters using SPFx logging
      let formatedMessage;
      if (entry.level === LogLevel.Error) {
        formatedMessage = new Error(`Message: ${entry.message} Data: ${JSON.stringify(entry.data)}`);
        // formatedMessage = `Message: ${entry.message} Data: ${JSON.stringify(entry.data)}`;
      } else {
        formatedMessage = `Message: ${entry.message} Data: ${JSON.stringify(entry.data)}`;
      }

      // [SPFx Logging] Calculate method to invoke verbose, info, warn or error
      const method = logLevelConversion[LogLevel[entry.level]];

      // [SPFx Logging] Call SPFx Logging system with the message received from PnP JS Logging
      Log[method](componentName, formatedMessage);
    });

    // [PnP JS Logging] Once create the custom listerner we should subscribe to it
    Logger.subscribe(listener);
  }

  private async _readItems(listName: string): Promise<void> {
    try {
      const items: INewsItem[] = await this._listItemServiceInstance.getNewsItems();
      const status: string = "Loaded news items using service";
      this.setState({ ...this.state, items, status});
    } catch (error) {
      this.setState({ ...this.state, errors: [...this.state.errors, error] });
    }
  }

  private _gerErrors() {
    return this.state.errors.length > 0
      ?
      <div style={{ color: "orangered" }} >
        <div>Errors:</div>
        {
          this.state.errors.map((item, idx) => {
            return (<div key={idx} >{JSON.stringify(item)}</div>);
          })
        }
      </div>
      : null;
  }
}
