import "azure-devops-ui/Core/override.css";
import "./Hub.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";
import { CommonServiceIds, IExtensionDataService, IGlobalMessagesService, IHostNavigationService, IProjectInfo, IProjectPageService, getClient } from "azure-devops-extension-api";

import { Header, TitleSize } from "azure-devops-ui/Header";
import { Page } from "azure-devops-ui/Page";

import { showRootComponent } from "../../Common";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { WorkItem, WorkItemTrackingRestClient, WorkItemType } from "azure-devops-extension-api/WorkItemTracking";
import { IterationWorkItems, TaskboardColumn, TaskboardColumns, TaskboardWorkItemColumn, TeamSettingsIteration, WorkRestClient } from "azure-devops-extension-api/Work";
import { CoreRestClient, WebApiTeam } from "azure-devops-extension-api/Core";
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
    selectedTeamIteration: TeamSettingsIteration | undefined;
    selectedTeamIterationId: string;
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

    private queryParamsTeam: string = '';
    private queryParamsTeamIteration: string = '';

    constructor(props: {}) {
        super(props);

        this.state = {
            project: '',
            selectedTabId: "navigation",
            teams: [],
            teamIterations: [],
            selectedTeam: '',
            selectedTeamName: '',
            selectedTeamIteration: undefined,
            selectedTeamIterationId: '',
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
            headerDescription,
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

        function sprintDatesHeading(selectedTeamIteration: TeamSettingsIteration | undefined): JSX.Element | null {
            if (selectedTeamIteration && (selectedTeamIteration.attributes.startDate || selectedTeamIteration.attributes.finishDate)) {
                return (
                    <p className="iteration-dates">{selectedTeamIteration.attributes.startDate ? selectedTeamIteration.attributes.startDate.toLocaleDateString(undefined, { timeZone: 'UTC' }) : ''} - {selectedTeamIteration.attributes.finishDate ? selectedTeamIteration.attributes.finishDate.toLocaleDateString(undefined, { timeZone: 'UTC' }) : ''}</p>
                );
            } else {
                return null;
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
                type: workItem.fields['System.WorkItemType'],
                storyPoints: workItem.fields['Microsoft.VSTS.Scheduling.StoryPoints'] ?? 0
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
                            <a href={workItem.url}>{workItem.id}</a> : {workItem.title} {workItem.storyPoints !== 0 && <span>({workItem.storyPoints})</span>} - {workItem.assignedTo}
                        </li>
                    );
                });

                return (
                    <div key={column.id} className="work-item-state">
                        <h3>{column.name}</h3>
                        <ul>{workItems}</ul>
                    </div>
                )
            })

            return (
                <div key={workItemType} className="work-item-type">
                    <h2>{workItemType}</h2>
                    <React.Fragment>
                        {workItemStates}
                    </React.Fragment>
                </div>
            );
        });

        return (
            <Page className="iteration-work-items-hub flex-grow">

                <Header title="Iteration Work Items Hub"
                    description={headerDescription}
                    titleSize={TitleSize.Large} />

                <div id="iteration-selections">
                    <p>Select a Team</p>
                    <Dropdown
                        ariaLabel="Select a team"
                        className="example-dropdown"
                        placeholder="Select a Team"
                        items={teamDropdownItems()}
                        selection={this.teamSelection}
                        onSelect={this.handleSelectTeam}
                        dismissOnSelect={true}
                    />

                    <p>Select an Iteration</p>
                    <Dropdown
                        ariaLabel="Select a team iteration"
                        className="example-dropdown"
                        placeholder="Select a Team Iteration"
                        items={teamIterationDropdownItems()}
                        selection={this.teamIterationSelection}
                        onSelect={this.handleSelectTeamIteration}
                        dismissOnSelect={true}
                    />
                </div>

                {this.state.selectedTeamIterationName && <h2 id="selected-iteration">Work Items for {this.state.selectedTeamName} : {this.state.selectedTeamIterationName}</h2>}
                {sprintDatesHeading(this.state.selectedTeamIteration)}

                {sortedWorkItems}
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

        // Check the URL for a stored team and iteration.
        const queryParams = await this.getQueryParams();
        if (queryParams.queryTeam) {
            this.queryParamsTeam = queryParams.queryTeam;
            if (queryParams.queryTeamIteration) {
                this.queryParamsTeamIteration = queryParams.queryTeamIteration;
            }
        }

        const saveDataTeam = await this.getSavedData();

        if (this.teams.length === 1) {
            this.teamSelection.select(0);
            this.setState({
                selectedTeam: this.teams[0].id
            });
            this.setState({
                selectedTeamName: this.teams[0].name
            });
            this.getTeamData();
        } else if (this.queryParamsTeam) {
            // See if the team selection from the URL is a valid team.
            const queryTeamIndex = this.teams.findIndex(t => t.id === this.queryParamsTeam);
            if (queryTeamIndex >= 0) {
                // Select the team.
                this.teamSelection.select(queryTeamIndex);
                this.setState({
                    selectedTeam: this.teams[queryTeamIndex].id
                });
                this.setState({
                    selectedTeamName: this.teams[queryTeamIndex].name
                });
                this.getTeamData();
            }
        } else if (saveDataTeam) {
            const saveDataTeamIndex = this.teams.findIndex(t => t.id === saveDataTeam);
            if (saveDataTeamIndex >= 0) {
                this.teamSelection.select(saveDataTeamIndex);
                this.setState({
                    selectedTeam: this.teams[saveDataTeamIndex].id
                });
                this.setState({
                    selectedTeamName: this.teams[saveDataTeamIndex].name
                });
                this.getTeamData();
            }
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

        let iteration;
        let iterationId = "";
        let iterationName = "";
        if (this.teamIterations.length === 1) {
            this.teamIterationSelection.select(0);

            iteration = this.teamIterations[0];
            iterationId = this.teamIterations[0].id;
            iterationName = this.teamIterations[0].name;
        } else {
            let currentIteration: TeamSettingsIteration | undefined;
            if (this.queryParamsTeamIteration) {
                currentIteration = this.teamIterations.find(i => i.id === this.queryParamsTeamIteration);
            }
            if (!currentIteration) {
                currentIteration = this.teamIterations.find(i => i.attributes.timeFrame === 1);
            }

            if (currentIteration) {
                this.teamIterationSelection.select(this.teamIterations.indexOf(currentIteration));

                iteration = currentIteration;
                iterationId = currentIteration.id;
                iterationName = currentIteration.name;
            }
        }

        if (iterationId !== '') {
            this.setState({
                selectedTeamIteration: iteration
            });
            this.setState({
                selectedTeamIterationId: iterationId
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
        this.iterationWorkItems = await workClient.getIterationWorkItems(teamContext, this.state.selectedTeamIterationId);

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
            const workItemColumns = await workClient.getWorkItemColumns(teamContext, this.state.selectedTeamIterationId);
            this.setState({ taskboardWorkItemColumns: workItemColumns });
        } catch (ex) {
            this.showToast('No work item columns were found for this team. These will be generated automatically from the work items.');
            manuallyGenerateTaskboardWorkItemColumns = true;
        }

        const witClient = getClient(WorkItemTrackingRestClient);
        // TODO handle more than 200 work items; this endpoint only accepts/returns up to 200
        this.workItems = await witClient.getWorkItems(this.iterationWorkItems.workItemRelations.map(wi => wi.target.id));
        this.setState({ workItems: this.workItems });

        if (manuallyGenerateTaskboardWorkItemColumns) {
            const manualWorkItemColumns = this.workItems.filter(wi => wi.fields['System.WorkItemType'] === 'Task').map<TaskboardWorkItemColumn>(wi => ({
                workItemId: wi.id,
                state: wi.fields['System.State'],
                column: wi.fields['System.State'],
                columnId: wi.fields['System.State']
            }));
            this.setState({ taskboardWorkItemColumns: manualWorkItemColumns });
        }

        this.workItemTypes = await witClient.getWorkItemTypes(this.state.project);
        this.setState({ workItemTypes: this.workItemTypes });
    }

    private handleSelectTeam = (_event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>): void => {
        this.setState({
            selectedTeam: item.id
        });
        this.setState({
            selectedTeamName: item.text ?? ''
        });
        this.setState({
            selectedTeamIteration: undefined
        });
        this.setState({
            selectedTeamIterationId: ''
        });
        this.setState({
            selectedTeamIterationName: ''
        });
        this.getTeamData();
        this.updateQueryParams();
        this.saveSelectedTeam();
    }

    private handleSelectTeamIteration = (_event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>): void => {
        this.setState({
            selectedTeamIteration: this.state.teamIterations.find(ti => ti.id === item.id)
        });
        this.setState({
            selectedTeamIterationId: item.id
        });
        this.setState({
            selectedTeamIterationName: item.text ?? ''
        });
        this.getTeamIterationData();
        this.updateQueryParams();
    }

    private async getQueryParams() {
        const navService = await SDK.getService<IHostNavigationService>(CommonServiceIds.HostNavigationService);
        const hash = await navService.getQueryParams();

        return { queryTeam: hash['selectedTeam'], queryTeamIteration: hash['selectedTeamIterationId'] };
    }

    private async getSavedData(): Promise<string> {
        await SDK.ready();
        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);

        let savedData = "";

        await dataManager.getValue<string>("selectedTeam" + this.state.project, {scopeType: "User"}).then((data) => {
            savedData = data;
        }, () => {
            // It's fine if no saved data is found.
        });

        return savedData;
    }

    private async saveSelectedTeam(): Promise<void> {
        await SDK.ready();
        const accessToken = await SDK.getAccessToken();
        const extDataService = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const dataManager = await extDataService.getExtensionDataManager(SDK.getExtensionContext().id, accessToken);
        await dataManager.setValue("selectedTeam" + this.state.project, this.state.selectedTeam, {scopeType: "User"}).then(() => {
            // No need to return anything.
        });
    }

    private showToast = async (message: string): Promise<void> => {
        const globalMessagesSvc = await SDK.getService<IGlobalMessagesService>(CommonServiceIds.GlobalMessagesService);
        globalMessagesSvc.addToast({
            duration: 3000,
            message: message
        });
    }

    private updateQueryParams = async () => {
        const navService = await SDK.getService<IHostNavigationService>(CommonServiceIds.HostNavigationService);
        navService.setQueryParams({ selectedTeam: "" + this.state.selectedTeam, selectedTeamIterationId: this.state.selectedTeamIterationId });
        navService.setDocumentTitle("" + this.state.selectedTeamName + " : " + this.state.selectedTeamIterationName + " - Iteration Work Items");
    }
}

showRootComponent(<HubContent />);
