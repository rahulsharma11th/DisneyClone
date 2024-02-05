import React from 'react';
import './App.css';
import Header from './components/Header';
import Home from './components/Home';

function App() {
  return (
    <div className="App">
  
    <Header />
    <Home />
    </div>
  );
}

export default App;



import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-http';

import UserReporteeInfo, { IUserReporteeInfoProps } from './components/UserReporteeInfo';

export interface IYourWebPartProps {
  userId: string;
}

export default class YourWebPart extends BaseClientSideWebPart<IYourWebPartProps> {
  public render(): void {
    this.context.msGraphClientFactory
      .getClient()
      .then((graphClient: MSGraphClient) => {
        const element: React.ReactElement<IUserReporteeInfoProps> = React.createElement(
          UserReporteeInfo,
          {
            graphClient: graphClient,
            userId: this.properties.userId
          }
        );

        ReactDom.render(element, this.domElement);
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Web Part Configuration" },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('userId', {
                  label: "User ID"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

