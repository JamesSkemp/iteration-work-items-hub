import "azure-devops-ui/Core/override.css";
import "./Hub.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IGlobalMessagesService, IHostPageLayoutService, IProjectInfo, IProjectPageService, getClient } from "azure-devops-extension-api";

import { Header, TitleSize } from "azure-devops-ui/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/HeaderCommandBar";
import { Page } from "azure-devops-ui/Page";
import { Tab, TabBar, TabSize } from "azure-devops-ui/Tabs";

import { OverviewTab } from "./OverviewTab";
import { NavigationTab } from "./NavigationTab";
import { ExtensionDataTab } from "./ExtensionDataTab";
import { MessagesTab } from "./MessagesTab";
import { showRootComponent } from "../../Common";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { WorkItem, WorkItemTrackingRestClient, WorkItemType } from "azure-devops-extension-api/WorkItemTracking";
import { IterationWorkItems, TaskboardColumns, TaskboardWorkItemColumn, TeamSettingsIteration, WorkRestClient } from "azure-devops-extension-api/Work";
import { CoreRestClient, ProjectInfo, WebApiTeam } from "azure-devops-extension-api/Core";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { ListSelection } from "azure-devops-ui/List";

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
    taskboardWorkItemColumns: TaskboardWorkItemColumn[];
    /**
     * All columns used in project team taskboards.
     */
    taskboardColumns?: TaskboardColumns;
    workItems: WorkItem[];
    /**
     * All work item types, such as Feature, Epic, Bug, Task, User Story.
     */
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

    constructor(props: {}) {
        super(props);

        this.state = {
            project: '',
            selectedTabId: "overview",
            teams: [],
            teamIterations: [],
            selectedTeam: '',
            selectedTeamIteration: '',
            taskboardWorkItemColumns: [],
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
            teams, teamIterations, workItems
        } = this.state;

        const interestedWorkItemTypes = ['Epic', 'Feature', 'User Story', 'Task', 'Bug'];

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

        /**
         * Returns all work items (user stories, tasks, bugs) as a custom object for later display.
         */
        const organizedWorkItems = workItems.map(workItem => {
            const newWorkItem = {
                id: workItem.id,
                title: workItem.fields['System.Title'],
                assignedTo: workItem.fields['System.AssignedTo'] ? workItem.fields['System.AssignedTo'].displayName : 'unassigned',
                url: workItem.url.replace('/_apis/wit/workItems/', '/_workitems/edit/'),
                boardColumn: workItem.fields['System.BoardColumn'],
                state: workItem.fields['System.State'],
                type: workItem.fields['System.WorkItemType']
            };

            const taskboardColumn = this.state.taskboardWorkItemColumns.find(wic => wic.workItemId === workItem.id);
            if (taskboardColumn) {
                newWorkItem.boardColumn = taskboardColumn.column;
            }

            return newWorkItem;
        });

        const sortedWorkItems = interestedWorkItemTypes.map(workItemType => {
            const typeMatchingWorkItems = organizedWorkItems.filter(wi => wi.type === workItemType);

            if (typeMatchingWorkItems.length === 0) {
                return;
            }

            const workItems = typeMatchingWorkItems.map(workItem => {
                return (
                    <li key={workItem.id}>
                        <a href={workItem.url}>{workItem.id}</a> : {workItem.title} ({workItem.assignedTo})
                        <br />{workItem.boardColumn}
                        <br />{workItem.state}
                        <br />{workItem.type}
                    </li>
                );
            });

            console.log(workItemType);
            console.log(typeMatchingWorkItems);

            return (
                <React.Fragment>
                    <h2>{workItemType}</h2>
                    <ul>{workItems}</ul>
                </React.Fragment>
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

                {sortedWorkItems}

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

        if (!this.iterationWorkItems) {
            this.showToast('No work items found for this iteration.');
            return;
        }

        // This will give us not only tasks and bugs, but also the user stories.
        // We'll use this full list to get all the work items later.
        this.setState({ iterationWorkItems: this.iterationWorkItems });

        this.taskboardColumns = await workClient.getColumns(teamContext);
        console.log('need this list of columns');
        console.log(this.taskboardColumns); // 5 - need this
        this.setState({ taskboardColumns: this.taskboardColumns });

        const workItemColumns = await workClient.getWorkItemColumns(teamContext, this.state.selectedTeamIteration);
        console.log(workItemColumns); // 6 (does not include user stories)
        this.setState({ taskboardWorkItemColumns: workItemColumns });

        const teamIteration = await workClient.getTeamIteration(teamContext, this.state.selectedTeamIteration);
        //console.log(teamIteration);

        const witClient = getClient(WorkItemTrackingRestClient);
        // TODO handle more than 200 work items
        this.workItems = await witClient.getWorkItems(this.iterationWorkItems.workItemRelations.map(wi => wi.target.id));
        console.log('need this list of work items');
        console.log(this.workItems);
        this.setState({ workItems: this.workItems });

        this.workItemTypes = await witClient.getWorkItemTypes(this.state.project);
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
