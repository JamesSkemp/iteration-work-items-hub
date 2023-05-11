import "azure-devops-ui/Core/override.css";
import "./Hub.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IGlobalMessagesService, IHostPageLayoutService, IProjectInfo, IProjectPageService, getClient } from "azure-devops-extension-api";

import { Header, TitleSize } from "azure-devops-ui/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/HeaderCommandBar";
import { Page } from "azure-devops-ui/Page";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";
import { IListItemDetails, ListItem } from 'azure-devops-ui/List';
import { DropdownSelection } from "azure-devops-ui/Utilities/DropdownSelection";

import { OverviewTab } from "./OverviewTab";
import { NavigationTab } from "./NavigationTab";
import { ExtensionDataTab } from "./ExtensionDataTab";
import { MessagesTab } from "./MessagesTab";
import { showRootComponent } from "../../Common";
import { ObservableArray, ObservableValue } from "azure-devops-ui/Core/Observable";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { WorkItem, WorkItemTrackingRestClient, WorkItemType } from "azure-devops-extension-api/WorkItemTracking";
import { IterationWorkItems, TaskboardColumns, TeamSettingsIteration, WorkRestClient } from "azure-devops-extension-api/Work";
import { CoreRestClient, ProjectInfo, WebApiTeam } from "azure-devops-extension-api/Core";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { ListSelection } from "azure-devops-ui/List";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";

interface IHubContentState {
    selectedTabId: string;
    headerDescription?: string;
    useLargeTitle?: boolean;
    useCompactPivots?: boolean;

    project: string;
    teams: WebApiTeam[];
    teamIterations: TeamSettingsIteration[];
    selectedTeam: string;
    selectedTeamIteration: string;
    iterationWorkItems?: IterationWorkItems;
    taskboardColumns?: TaskboardColumns;
    workItems: WorkItem[];
    workItemTypes: WorkItemType[];
}

class HubContent extends React.Component<{}, IHubContentState> {
    private project: IProjectInfo | undefined;
    private teams: WebApiTeam[] = [];
    private teamIterations: TeamSettingsIteration[] = [];
    private iterationWorkItems: IterationWorkItems | undefined;
    private taskboardColumns: TaskboardColumns | undefined;
    private workItems: WorkItem[] = [];
    private workItemTypes: WorkItemType[] = [];

    private teamSelection = new ListSelection();
    private teamIterationSelection = new ListSelection();
    private teamItems = new ArrayItemProvider(this.teams);

    private data = new ObservableArray<IListBoxItem<string>>();
    private workItemTypeValue = new ObservableValue("");
    private selection = new ListSelection();
    private workItemTypesOld = new ObservableArray<IListBoxItem<string>>();

    constructor(props: {}) {
        super(props);

        this.state = {
            project: '',
            selectedTabId: "overview",
            teams: [],
            teamIterations: [],
            selectedTeam: '',
            selectedTeamIteration: '',
            workItems: [],
            workItemTypes: []
        };
    }

    public componentDidMount() {
        SDK.init();
        this.getCustomData();
    }

    public render(): JSX.Element {
        const {
            selectedTabId, headerDescription, useCompactPivots, useLargeTitle,
            teams, teamIterations, iterationWorkItems, taskboardColumns, workItems, workItemTypes
        } = this.state;

        const theTeams = teams.map((team, index) => {
            return (
                <li key={team.id}>
                    {team.name}
                </li>
            );
        });

        function teamDropdownItems(): Array<IListBoxItem<{}>> {
            if (teams) {
                return teams.map<IListBoxItem<{}>>(team => ({
                    id: team.id, text: team.name
                }));
            } else {
                return [];
            }
        }

        function teamIterationDropdownItems(): Array<IListBoxItem<{}>> {
            if (teamIterations) {
                return teamIterations.map<IListBoxItem<{}>>(teamIteration => ({
                    id: teamIteration.id, text: teamIteration.name
                }));
            } else {
                return [];
            }
        }

        const theIterationWorkItems = iterationWorkItems?.workItemRelations.map((workItemRelation, index) => {
            return (
                <li key={workItemRelation.target.id}>
                    {workItemRelation.target.id} : {workItemRelation.source?.id}
                </li>
            );
        });

        const theTaskboardColumns = taskboardColumns?.columns.map((taskboardColumn, index) => {
            return (
                <li key={taskboardColumn.id}>
                    {taskboardColumn.name}
                </li>
            );
        });

        const theWorkItems = workItems.map((workItem, index) => {
            const workItemUrl = workItem.url.replace('/_apis/wit/workItems/', '/_workitems/edit/');
            const assignedTo = workItem.fields['System.AssignedTo'] ? workItem.fields['System.AssignedTo'].displayName : 'unassigned';

            //console.log(workItemUrl);

            return (
                <li key={workItem.id}>
                    <a href={workItemUrl}>{workItem.id}</a> : {workItem.fields['System.Title']} ({assignedTo})
                </li>
            );
        });

        const theWorkItemTypes = workItemTypes.map((workItemType, index) => {
            return (
                <li key={index}>
                    {workItemType.name}
                </li>
            );
        });

        return (
            <Page className="sample-hub flex-grow">

                <Header title="Iteration Work Items Hub"
                    commandBarItems={this.getCommandBarItems()}
                    description={headerDescription}
                    titleSize={useLargeTitle ? TitleSize.Large : TitleSize.Medium} />

                <h2>Select a Team</h2>
                <Dropdown
                    ariaLabel="Select a team"
                    className="example-dropdown"
                    placeholder="Select a Team"
                    items={teamDropdownItems()}
                    selection={this.teamSelection}
                    onSelect={this.handleSelectTeam}
                />

                <h2>Select an Iteration</h2>
                <Dropdown
                    ariaLabel="Select a team iteration"
                    className="example-dropdown"
                    placeholder="Select a Team Iteration"
                    items={teamIterationDropdownItems()}
                    selection={this.teamIterationSelection}
                    onSelect={this.handleSelectTeamIteration}
                />

                <ul>{theIterationWorkItems}</ul>

                <ul>{theTaskboardColumns}</ul>

                <ol>{theWorkItems}</ol>

                <ul>{theWorkItemTypes}</ul>

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
            this.showToast('No projects found.');
            return;
        }
        this.setState({ project: this.project.id });

        // Get teams.
        const coreClient = getClient(CoreRestClient);
        this.teams = await coreClient.getTeams(this.state.project);
        if (!this.teams) {
            this.showToast('No teams found.');
            return;
        }
        this.setState({ teams: this.teams });

        let teamId = "";
        if (this.teams.length === 1) {
            teamId = this.teams[0].id;
            this.setState({
                selectedTeam: this.teams[0].id
            });
            this.getTeamData();
        }

    }

    private async getTeamData() {
        await SDK.ready();
        const teamContext = { projectId: this.state.project, teamId: this.state.selectedTeam, project: "", team: "" };

        // Get taskboard columns.
        const workClient = getClient(WorkRestClient);
        this.teamIterations = await workClient.getTeamIterations(teamContext);
        if (!this.teamIterations) {
            this.showToast('No team iterations found.');
            return;
        }
        this.setState({ teamIterations: this.teamIterations });

        let iterationId = "";
        if (this.teamIterations.length === 1) {
            iterationId = this.teamIterations[0].id;
        } else {
            let currentIteration = this.teamIterations.find(i => i.attributes.timeFrame === 1);
            if (currentIteration) {
                iterationId = currentIteration.id;
            }
        }

        if (iterationId !== '') {
            this.setState({
                selectedTeamIteration: iterationId
            });
            this.getTeamIterationData();
        }
    }

    private async getTeamIterationData() {
        await SDK.ready();
        const teamContext = { projectId: this.state.project, teamId: this.state.selectedTeam, project: "", team: "" };

        const workClient = getClient(WorkRestClient);
        this.iterationWorkItems = await workClient.getIterationWorkItems(teamContext, this.state.selectedTeamIteration);
        console.log('need this list of item relations');
        console.log(this.iterationWorkItems); // 10 (stories + tasks + bugs)
        this.setState({ iterationWorkItems: this.iterationWorkItems });

        this.taskboardColumns = await workClient.getColumns(teamContext);
        console.log('need this list of columns');
        console.log(this.taskboardColumns); // 5 - need this
        this.setState({ taskboardColumns: this.taskboardColumns });

        const workItemColumns = await workClient.getWorkItemColumns(teamContext, this.state.selectedTeamIteration);
        console.log(workItemColumns); // 6 (does not include user stories)

        const teamIteration = await workClient.getTeamIteration(teamContext, this.state.selectedTeamIteration);
        console.log(teamIteration);

        const witClient = getClient(WorkItemTrackingRestClient);
        // TODO handle more than 200 work items
        this.workItems = await witClient.getWorkItems(this.iterationWorkItems.workItemRelations.map(wi => wi.target.id));
        console.log('need this list of work items');
        console.log(this.workItems);
        this.setState({ workItems: this.workItems });

        this.workItemTypes = await witClient.getWorkItemTypes(this.state.project);
        // will probably just hard-code these
        console.log(this.workItemTypes);
        this.setState({ workItemTypes: this.workItemTypes });
    }

    private onSelectedTabChanged = (newTabId: string) => {
        this.setState({
            selectedTabId: newTabId
        })
    }

    private handleSelectTeam = (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>): void => {
        this.setState({
            selectedTeam: item.id
        });
        this.setState({
            selectedTeamIteration: ''
        });
        this.getTeamData();
    }

    private handleSelectTeamIteration = (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>): void => {
        this.setState({
            selectedTeamIteration: item.id
        });
        this.getTeamIterationData();
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

    private showToast = async (message: string): Promise<void> => {
        const globalMessagesSvc = await SDK.getService<IGlobalMessagesService>(CommonServiceIds.GlobalMessagesService);
        globalMessagesSvc.addToast({
            duration: 3000,
            message: message
        });
    }
}

showRootComponent(<HubContent />);
