import "./Hub.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IHostPageLayoutService, IProjectInfo, IProjectPageService, getClient } from "azure-devops-extension-api";

import { Header, TitleSize } from "azure-devops-ui/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/HeaderCommandBar";
import { Page } from "azure-devops-ui/Page";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";

import { OverviewTab } from "./OverviewTab";
import { NavigationTab } from "./NavigationTab";
import { ExtensionDataTab } from "./ExtensionDataTab";
import { MessagesTab } from "./MessagesTab";
import { showRootComponent } from "../../Common";
import { ObservableArray, ObservableValue } from "azure-devops-ui/Core/Observable";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { WorkItem, WorkItemTrackingRestClient, WorkItemType } from "azure-devops-extension-api/WorkItemTracking";
import { TaskboardColumn, TaskboardColumns, TaskboardWorkItemColumn, WorkRestClient } from "azure-devops-extension-api/Work";
import { CoreRestClient, WebApiTeam } from "azure-devops-extension-api/Core";
import { BoardsRestClient } from 'azure-devops-extension-api/Boards';
import { Dropdown } from "azure-devops-ui/Dropdown";
import { ListSelection } from "azure-devops-ui/List";

interface IHubContentState {
    selectedTabId: string;
    fullScreenMode: boolean;
    headerDescription?: string;
    useLargeTitle?: boolean;
    useCompactPivots?: boolean;
}

class HubContent extends React.Component<{}, IHubContentState> {
    private project: IProjectInfo | undefined;
    private teams: WebApiTeam[] = [];
    private taskboardColumns: TaskboardColumns | undefined;
    private workItems: WorkItem[] = [];

    private data = new ObservableArray<IListBoxItem<string>>();
    private workItemTypeValue = new ObservableValue("");
    private selection = new ListSelection();
    private workItemTypes: WorkItemType[] = [];
    private workItemTypesOld = new ObservableArray<IListBoxItem<string>>();

    constructor(props: {}) {
        super(props);

        this.state = {
            selectedTabId: "overview",
            fullScreenMode: false
        };
    }

    public componentDidMount() {
        SDK.init();
        this.getCustomData();
    }

    public render(): JSX.Element {

        const { selectedTabId, headerDescription, useCompactPivots, useLargeTitle } = this.state;


        return (
            <Page className="sample-hub flex-grow">

                <Header title="Sample Hub"
                    commandBarItems={this.getCommandBarItems()}
                    description={headerDescription}
                    titleSize={useLargeTitle ? TitleSize.Large : TitleSize.Medium} />

                <Dropdown<string>
                        className="sample-work-item-type-picker"
                        items={this.data}
                        onSelect={(event, item) => { this.workItemTypeValue.value = item.data! }}
                        selection={this.selection}
                    />

                <TabBar
                    onSelectedTabChanged={this.onSelectedTabChanged}
                    selectedTabId={selectedTabId}
                    tabSize={useCompactPivots ? TabSize.Compact : TabSize.Tall}>

                    <Tab name="Overview" id="overview" />
                    <Tab name="Navigation" id="navigation" />
                    <Tab name="Extension Data" id="extensionData" />
                    <Tab name="Messages" id="messages" />
                </TabBar>

                { this.getPageContent() }
            </Page>
        );
    }

    private async getCustomData() {
        await SDK.ready();

        // Get the project.
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        this.project = await projectService.getProject();

        if (!this.project) {
            console.log('No project found.');
            return;
        }

        console.log(this.project);
        //this.data.push({ id: this.project.id, data: this.project.name, text: this.project.id + ' ' + this.project.name });

        // Get teams.
        const coreClient = getClient(CoreRestClient);
        this.teams = await coreClient.getTeams(this.project.id);
        console.log('need one of these teams');
        console.log(this.teams);

        let teamId = "1e538049-e108-44be-9480-74fbfc79500f";

        const teamContext = { projectId: this.project.id, teamId: teamId, project: "", team: "" };

        // Get taskboard columns.
        const workClient = getClient(WorkRestClient);
        const iterations = await workClient.getTeamIterations(teamContext);
        console.log('need one of these iterations');
        console.log(iterations); // 8

        let iterationId = "7e52b420-0877-4c87-bfed-54637e976bdc";

        const iterationWorkItems = await workClient.getIterationWorkItems(teamContext, iterationId);
        console.log('need this list of items');
        console.log(iterationWorkItems); // 10 (stories + tasks + bugs)

        this.taskboardColumns = await workClient.getColumns(teamContext);
        console.log('need this list of columns');
        console.log(this.taskboardColumns); // 5 - need this

        const workItemColumns = await workClient.getWorkItemColumns(teamContext, iterationId);
        console.log(workItemColumns); // 6 (does not include user stories)

        const teamIteration = await workClient.getTeamIteration(teamContext, iterationId);
        console.log(teamIteration);

        const witClient = getClient(WorkItemTrackingRestClient);
        // TODO handle more than 200 work items
        this.workItems = await witClient.getWorkItems(iterationWorkItems.workItemRelations.map(wi => wi.target.id));
        console.log(this.workItems);

        this.workItemTypes = await witClient.getWorkItemTypes(this.project.id);
        // will probably just hard-code these
        console.log(this.workItemTypes);
    }

    private onSelectedTabChanged = (newTabId: string) => {
        this.setState({
            selectedTabId: newTabId
        })
    }

    private getPageContent() {
        const { selectedTabId } = this.state;
        if (selectedTabId === "overview") {
            return <OverviewTab />;
        }
        else if (selectedTabId === "navigation") {
            return <NavigationTab />;
        }
        else if (selectedTabId === "extensionData") {
            return <ExtensionDataTab />;
        }
        else if (selectedTabId === "messages") {
            return <MessagesTab />;
        }
    }

    private getCommandBarItems(): IHeaderCommandBarItem[] {
        return [
            {
              id: "panel",
              text: "Panel",
              onActivate: () => { this.onPanelClick() },
              iconProps: {
                iconName: 'Add'
              },
              isPrimary: true,
              tooltipProps: {
                text: "Open a panel with custom extension content"
              }
            },
            {
              id: "messageDialog",
              text: "Message",
              onActivate: () => { this.onMessagePromptClick() },
              tooltipProps: {
                text: "Open a simple message dialog"
              }
            },
            {
                id: "fullScreen",
                ariaLabel: this.state.fullScreenMode ? "Exit full screen mode" : "Enter full screen mode",
                iconProps: {
                    iconName: this.state.fullScreenMode ? "BackToWindow" : "FullScreen"
                },
                onActivate: () => { this.onToggleFullScreenMode() }
            },
            {
              id: "customDialog",
              text: "Custom Dialog",
              onActivate: () => { this.onCustomPromptClick() },
              tooltipProps: {
                text: "Open a dialog with custom extension content"
              }
            }
        ];
    }

    private async onMessagePromptClick(): Promise<void> {
        const dialogService = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
        dialogService.openMessageDialog("Use large title?", {
            showCancel: true,
            title: "Message dialog",
            onClose: (result) => {
                this.setState({ useLargeTitle: result });
            }
        });
    }

    private async onCustomPromptClick(): Promise<void> {
        const dialogService = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
        dialogService.openCustomDialog<boolean | undefined>(SDK.getExtensionContext().id + ".panel-content", {
            title: "Custom dialog",
            configuration: {
                message: "Use compact pivots?",
                initialValue: this.state.useCompactPivots
            },
            onClose: (result) => {
                if (result !== undefined) {
                    this.setState({ useCompactPivots: result });
                }
            }
        });
    }

    private async onPanelClick(): Promise<void> {
        const panelService = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
        panelService.openPanel<boolean | undefined>(SDK.getExtensionContext().id + ".panel-content", {
            title: "My Panel",
            description: "Description of my panel",
            configuration: {
                message: "Show header description?",
                initialValue: !!this.state.headerDescription
            },
            onClose: (result) => {
                if (result !== undefined) {
                    this.setState({ headerDescription: result ? "This is a header description" : undefined });
                }
            }
        });
    }

    private async onToggleFullScreenMode(): Promise<void> {
        const fullScreenMode = !this.state.fullScreenMode;
        this.setState({ fullScreenMode });

        const layoutService = await SDK.getService<IHostPageLayoutService>(CommonServiceIds.HostPageLayoutService);
        layoutService.setFullScreenMode(fullScreenMode);
    }
}

showRootComponent(<HubContent />);

function localeIgnoreCaseComparer(a: string, b: string): number {
    throw new Error("Function not implemented.");
}
