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
import { IterationWorkItems, TaskboardColumn, TaskboardColumns, TaskboardWorkItemColumn, TeamSettingsIteration, WorkRestClient } from "azure-devops-extension-api/Work";
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
    selectedTeamName: string;
    selectedTeamIteration: string;
    selectedTeamIterationName: string;
    iterationWorkItems?: IterationWorkItems;
    taskboardWorkItemColumns: TaskboardWorkItemColumn[];
    /**
     * All columns used in project team taskboards.
     */
    taskboardColumns: TaskboardColumn[];
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
            selectedTeamName: '',
            selectedTeamIteration: '',
            selectedTeamIterationName: '',
            taskboardWorkItemColumns: [],
            taskboardColumns: [],
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

            const taskboardColumns = this.state.taskboardColumns;

            const workItemStates = taskboardColumns.map(column => {
                const columnMatchingWorkItems = typeMatchingWorkItems.filter(wi => wi.boardColumn === column.name);

                if (columnMatchingWorkItems.length === 0) {
                    return;
                }

                const workItems = columnMatchingWorkItems.map(workItem => {
                    return (
                        <li key={workItem.id}>
                            <a href={workItem.url}>{workItem.id}</a> : {workItem.title} - {workItem.assignedTo}
                        </li>
                    );
                });

                return (
                    <React.Fragment>
                        <h4>{column.name}</h4>
                        <ul>{workItems}</ul>
                    </React.Fragment>
                )
            })

            return (
                <React.Fragment>
                    <h3>{workItemType}</h3>
                    <React.Fragment>
                        {workItemStates}
                    </React.Fragment>
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
                    dismissOnSelect={true}
                />

                <h2>Select an Iteration</h2>
                <Dropdown
                    ariaLabel="Select a team iteration"
                    className="example-dropdown"
                    placeholder="Select a Team Iteration"
                    items={teamIterationDropdownItems()}
                    selection={this.teamIterationSelection}
                    onSelect={this.handleSelectTeamIteration}
                    dismissOnSelect={true}
                />

                <h3>Work Items for {this.state.selectedTeamName} : {this.state.selectedTeamIterationName}</h3>

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
            this.setState({
                selectedTeamName: this.teams[0].name
            })
            this.getTeamData();
        }

    }

    private async getTeamData() {
        await SDK.ready();
        const teamContext = { projectId: this.state.project, teamId: this.state.selectedTeam, project: "", team: "" };

        // Get team iterations.
        const workClient = getClient(WorkRestClient);
        this.teamIterations = await workClient.getTeamIterations(teamContext);
        if (!this.teamIterations) {
            this.showToast('No team iterations found.');
            return;
        }
        this.setState({ teamIterations: this.teamIterations });

        let iterationId = "";
        let iterationName = "";
        if (this.teamIterations.length === 1) {
            iterationId = this.teamIterations[0].id;
            iterationName = this.teamIterations[0].name;
        } else {
            let currentIteration = this.teamIterations.find(i => i.attributes.timeFrame === 1);
            if (currentIteration) {
                iterationId = currentIteration.id;
                iterationName = currentIteration.name;
            }
        }

        if (iterationId !== '') {
            this.setState({
                selectedTeamIteration: iterationId
            });
            this.setState({
                selectedTeamIterationName: iterationName
            })
            this.getTeamIterationData();
        }
    }

    private async getTeamIterationData() {
        await SDK.ready();
        const teamContext = { projectId: this.state.project, teamId: this.state.selectedTeam, project: "", team: "" };

        const workClient = getClient(WorkRestClient);
        this.iterationWorkItems = await workClient.getIterationWorkItems(teamContext, this.state.selectedTeamIteration);

        if (!this.iterationWorkItems || this.iterationWorkItems.workItemRelations.length === 0) {
            this.showToast('No work items found for this iteration.');
            this.setState({ workItems: [] });
            return;
        }

        // This will give us not only tasks and bugs, but also the user stories.
        // We'll use this full list to get all the work items later.
        this.setState({ iterationWorkItems: this.iterationWorkItems });

        // Get taskboard columns.
        try {
            this.taskboardColumns = await workClient.getColumns(teamContext);
        } catch (ex) {
            this.taskboardColumns = undefined;
        }
        //console.log(this.taskboardColumns);

        //console.log('need this list of columns');
        //console.log(this.taskboardColumns); // 5 - need this
        if (!this.taskboardColumns || this.taskboardColumns.columns.length === 0) {
            this.showToast('No taskboard columns can be found for this team. Default columns will be used.');
            this.setState({ taskboardColumns: [
                { id: 'New', name: 'New', order: 0, mappings: [] },
                { id: 'Active', name: 'Active', order: 1, mappings: [] },
                { id: 'Resolved', name: 'Resolved', order: 2, mappings: [] },
                { id: 'Closed', name: 'Closed', order: 3, mappings: [] },
            ]});
        } else {
            this.setState({ taskboardColumns: this.taskboardColumns.columns });
        }

        let manuallyGenerateTaskboardWorkItemColumns = false;
        try {
            const workItemColumns = await workClient.getWorkItemColumns(teamContext, this.state.selectedTeamIteration);
            this.setState({ taskboardWorkItemColumns: workItemColumns });
        } catch (ex) {
            this.showToast('Unable to get work item columns for this team. These will be generated ');
            manuallyGenerateTaskboardWorkItemColumns = true;
        }

        //const teamIteration = await workClient.getTeamIteration(teamContext, this.state.selectedTeamIteration);
        //console.log(teamIteration);

        const witClient = getClient(WorkItemTrackingRestClient);
        // TODO handle more than 200 work items
        this.workItems = await witClient.getWorkItems(this.iterationWorkItems.workItemRelations.map(wi => wi.target.id));
        //console.log('need this list of work items');
        //console.log(this.workItems);
        this.setState({ workItems: this.workItems });

        if (manuallyGenerateTaskboardWorkItemColumns) {
            const manualWorkItemColumns = this.workItems.filter(wi => wi.fields['System.WorkItemType'] === 'Task').map<TaskboardWorkItemColumn>(wi => ({
                workItemId: wi.id,
                state: wi.fields['System.State'],
                column: wi.fields['System.State'],
                columnId: wi.fields['System.State']
            }));
            this.setState({ taskboardWorkItemColumns: manualWorkItemColumns });
            // TODO
            // {workItemId: 102, state: 'Closed', column: 'Closed', columnId: 'f60021ec-f65c-4b86-90e8-81db06250b14'}
        }

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
            selectedTeamName: item.text ?? ''
        });
        this.setState({
            selectedTeamIteration: ''
        });
        this.setState({
            selectedTeamIterationName: ''
        });
        this.getTeamData();
    }

    private handleSelectTeamIteration = (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>): void => {
        this.setState({
            selectedTeamIteration: item.id
        });
        this.setState({
            selectedTeamIterationName: item.text ?? ''
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
