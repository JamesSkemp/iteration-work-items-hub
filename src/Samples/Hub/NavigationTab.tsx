import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IHostNavigationService } from "azure-devops-extension-api";

import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";

export interface INavigationTabState {
    currentQueryParams?: string;
}

export class NavigationTab extends React.Component<{}, INavigationTabState> {

    constructor(props: {}) {
        super(props);
        this.state = {};
    }

    public componentDidMount() {
        this.initialize();
    }

    public render(): JSX.Element {
        const { currentQueryParams } = this.state;
        return (
            <div className="page-content page-content-top flex-column rhythm-vertical-16">
                {
                    currentQueryParams &&
                    <div>Current query params: {currentQueryParams}</div>
                }
                <ButtonGroup>
                    <Button text="Get QueryParams" primary={true} onClick={this.onGetQueryParamsClick} />
                </ButtonGroup>
            </div>
        );
    }

    private async initialize() {
    }

    private onGetQueryParamsClick = async (): Promise<void> => {
        const navService = await SDK.getService<IHostNavigationService>(CommonServiceIds.HostNavigationService);
        const hash = await navService.getQueryParams();
        this.setState({ currentQueryParams: JSON.stringify(hash) });
    }
}
