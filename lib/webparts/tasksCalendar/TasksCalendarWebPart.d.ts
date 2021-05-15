import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import 'fullcalendar';
export interface ITasksCalendarWebPartProps {
    listName: string;
}
export default class TasksCalendarWebPart extends BaseClientSideWebPart<ITasksCalendarWebPartProps> {
    render(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    protected get disableReactivePropertyChanges(): boolean;
    private displayTasks;
    private updateTask;
}
//# sourceMappingURL=TasksCalendarWebPart.d.ts.map