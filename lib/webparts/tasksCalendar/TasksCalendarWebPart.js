var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './TasksCalendarWebPart.module.scss';
import * as strings from 'TasksCalendarWebPartStrings';
import * as $ from 'jquery';
import 'fullcalendar';
import * as moment from 'moment';
import { SPHttpClient } from '@microsoft/sp-http';
var COLORS = ['#466365', '#B49A67', '#93B7BE', '#E07A5F', '#849483', '#084C61', '#DB3A34'];
var TasksCalendarWebPart = /** @class */ (function (_super) {
    __extends(TasksCalendarWebPart, _super);
    function TasksCalendarWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    TasksCalendarWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n      <div class=\"" + styles.tasksCalendar + "\">\n        <link type=\"text/css\" rel=\"stylesheet\" href=\"//cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.4.0/fullcalendar.min.css\" />\n        <div id=\"calendar\"></div>\n      </div>";
        this.displayTasks();
    };
    Object.defineProperty(TasksCalendarWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    TasksCalendarWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('listName', {
                                    label: strings.ListNameFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    Object.defineProperty(TasksCalendarWebPart.prototype, "disableReactivePropertyChanges", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    TasksCalendarWebPart.prototype.displayTasks = function () {
        var _this = this;
        $('#calendar').fullCalendar('destroy');
        $('#calendar').fullCalendar({
            weekends: false,
            header: {
                left: 'prev,next today',
                center: 'title',
                right: 'month,basicWeek,basicDay'
            },
            displayEventTime: false,
            // open up the display form when a user clicks on an event
            eventClick: function (calEvent, jsEvent, view) {
                window.location = _this.context.pageContext.web.absoluteUrl +
                    "/Lists/" + escape(_this.properties.listName) + "/DispForm.aspx?ID=" + calEvent.id;
            },
            editable: true,
            timezone: "UTC",
            droppable: true,
            // update the end date when a user drags and drops an event
            eventDrop: function (event, delta, revertFunc) {
                _this.updateTask(event.id, event.start, event.end);
            },
            // put the events on the calendar
            events: function (start, end, timezone, callback) {
                var startDate = start.format('YYYY-MM-DD');
                var endDate = end.format('YYYY-MM-DD');
                var restQuery = "/_api/Web/Lists/GetByTitle('" + escape(_this.properties.listName) + "')/items?$select=ID,Title,\
    Status,StartDate,DueDate,AssignedTo/Title&$expand=AssignedTo&\
    $filter=((DueDate ge '" + startDate + "' and DueDate le '" + endDate + "')or(StartDate ge '" + startDate + "' and StartDate le '" + endDate + "'))";
                _this.context.spHttpClient.get(_this.context.pageContext.web.absoluteUrl + restQuery, SPHttpClient.configurations.v1, {
                    headers: {
                        'Accept': "application/json;odata.metadata=none"
                    }
                })
                    .then(function (response) {
                    return response.json();
                })
                    .then(function (data) {
                    var personColors = {};
                    var colorNo = 0;
                    var events = data.value.map(function (task) {
                        var assignedTo = task.AssignedTo.map(function (person) {
                            return person.Title;
                        }).join(', ');
                        var color = personColors[assignedTo];
                        if (!color) {
                            color = COLORS[colorNo++];
                            personColors[assignedTo] = color;
                        }
                        if (colorNo >= COLORS.length) {
                            colorNo = 0;
                        }
                        return {
                            title: task.Title + " - " + assignedTo,
                            id: task.ID,
                            color: color,
                            start: moment.utc(task.StartDate).add("1", "days"),
                            end: moment.utc(task.DueDate).add("1", "days") // add one day to end date so that calendar properly shows event ending on that day
                        };
                    });
                    callback(events);
                });
            }
        });
    };
    TasksCalendarWebPart.prototype.updateTask = function (id, startDate, dueDate) {
        var _this = this;
        // subtract the previously added day to the date to store correct date
        var sDate = moment.utc(startDate).add("-1", "days").format('YYYY-MM-DD') + "T" +
            startDate.format("hh:mm") + ":00Z";
        if (!dueDate) {
            dueDate = startDate;
        }
        var dDate = moment.utc(dueDate).add("-1", "days").format('YYYY-MM-DD') + "T" +
            dueDate.format("hh:mm") + ":00Z";
        this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "  /_api/Web/Lists/getByTitle('" + escape(this.properties.listName) + "')/Items(" + id + ")", SPHttpClient.configurations.v1, {
            body: JSON.stringify({
                StartDate: sDate,
                DueDate: dDate,
            }),
            headers: {
                Accept: "application/json;odata=nometadata",
                "Content-Type": "application/json;odata=nometadata",
                "IF-MATCH": "*",
                "X-Http-Method": "PATCH"
            }
        })
            .then(function (response) {
            if (response.ok) {
                alert("Update Successful");
            }
            else {
                alert("Update Failed");
            }
            _this.displayTasks();
        });
    };
    return TasksCalendarWebPart;
}(BaseClientSideWebPart));
export default TasksCalendarWebPart;
//# sourceMappingURL=TasksCalendarWebPart.js.map