import * as React from 'react';
import styles from './TimesheetDashboard.module.scss';
import { ITimesheetDashboardProps } from './ITimesheetDashboardProps';
import { escape, isEqual } from '@microsoft/sp-lodash-subset';
import { DatePicker, DefaultButton, Dialog, DialogFooter, Dropdown, IIconProps, Icon, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { IFieldInfo, Social, sp } from '@pnp/sp/presets/all';
import $ from "jquery";
import * as moment from 'moment';

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, IHttpClientOptions, HttpClient } from '@microsoft/sp-http';


import * as Excel from 'exceljs';
import { saveAs } from 'file-saver';

require("../assets/css/fabric.min.css");
require("../assets/css/jquery.dataTables.css");
require("../assets/css/buttons.dataTables.css");
require("../assets/css/style.css");

require("../assets/Js/jquery.dataTables.js");

let spcontext;

let colums = [
  { header: "Employee Name", key: "EmployeeName" },
  { header: "Project", key: "Project" },
  { header: "Hours", key: "Hours" },
  { header: "Staus", key: "Staus" },
  { header: "Task Description", key: "TaskDescription" },
  { header: "Department", key: "Department" },
  { header: "Date", key: "Date" },

];

const FilterIcon: IIconProps = { iconName: 'Filter' };
const ResetIcon: IIconProps = { iconName: 'Refresh' };
const ExportIcon: IIconProps = { iconName: 'DownloadDocument' };
const SendMail: IIconProps = { iconName: 'MailLowImportance' };

const FlowURL = {
  SendMail: "https://prod-18.centralindia.logic.azure.com:443/workflows/8ae43dd9990c41c5859f73f8797afd08/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=DgICfoaa0bfWVijzB7nbgqpalf15YaSz5xqwgjRHky0",
}
export interface ITimesheetDashboardState {
  TimeTrackingData: any;
  StartDate: any;
  EndDate: any;
  FilteredListData: any;
  ChoiceDepartments: any;
  SelectedDepartment: any;
  SelectedDepartmentID: any;
  ChoiceEmployees: any;
  SelectedEmployee: any;
  SelectedEmployeeID: any;
  HideDialog: boolean;
  MsgDialog: boolean;
  EmailID: any;
  Currentuser:any;
}
const dialogContentProps = {
  title: 'Send Report',
  subText: "Add the recipient's email address to send the Timesheet report."
};

const dialogContentProps1 = {
  title: 'Report Sent',
  subText: "The timesheet report has been sent successfully."
};


export default class TimesheetDashboard extends React.Component<ITimesheetDashboardProps, ITimesheetDashboardState> {

  constructor(props: ITimesheetDashboardProps, state: ITimesheetDashboardState) {
    super(props);
    this.state = {
      TimeTrackingData: [],
      StartDate: "",
      EndDate: "",
      FilteredListData: [],
      ChoiceDepartments: [],
      SelectedDepartment: "",
      SelectedDepartmentID: [],
      ChoiceEmployees: [],
      SelectedEmployee: [],
      SelectedEmployeeID: "",
      HideDialog: true,
      MsgDialog: true,
      EmailID: [],
      Currentuser: ''

    };

  }

  public render(): React.ReactElement<ITimesheetDashboardProps> {

    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <>
        <div id='TimeSheet' className='ms-Grid'>
          <div>
            <h3>Time Tracking</h3>
            <div className='ms-Grid-row'>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                <DatePicker
                  label='Start Date'
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  onSelectDate={(date) => { this.updateDatesForFilter("Start", date); }}
                  value={this.state.StartDate}
                />
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                <DatePicker
                  label='End Date'
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  onSelectDate={(date) => { this.updateDatesForFilter("End", date); }}
                  value={this.state.EndDate}
                />
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                <Dropdown
                  placeholder="Select Department"
                  label="Department"
                  options={this.state.ChoiceDepartments}
                  onChange={(e, option, index) => this.setState({ SelectedDepartment: option.text, SelectedDepartmentID: option.key })}
                  selectedKey={this.state.SelectedDepartmentID}
                />
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2 mb-10">
                <Dropdown
                  placeholder="Select Employee"
                  label="Employee"
                  multiSelect
                  options={this.state.ChoiceEmployees}
                  // onChange={(e, option, index) => this.setState({ SelectedEmployee: option.text, SelectedEmployeeID: option.key })}
                  onChange={(e, option) => this.onItemChange(e, option)}
                  selectedKeys={this.state.SelectedEmployeeID}
                />
              </div>
              <div className="ms-Grid-col ms-sm12 ms-md4 ms-lg4 mb-10 d-flex-custom">
                <PrimaryButton onClick={() => this.filterData()} iconProps={FilterIcon}>Filter</PrimaryButton>
                <PrimaryButton onClick={() => this.resetFilter()} iconProps={ResetIcon}>Reset</PrimaryButton>
                <PrimaryButton onClick={() => this.saveExcel()} iconProps={ExportIcon}>Export</PrimaryButton>
                <PrimaryButton onClick={() => this.setState({ HideDialog: false })} iconProps={SendMail}>Send</PrimaryButton>
              </div>
            </div>
            <br />
            <table id="myTable" className="display">
              <thead>
                <tr>
                  <th>Employee Name</th>
                  <th>Project</th>
                  <th>Hours</th>
                  <th>Staus</th>
                  <th>Task Description</th>
                  <th>Department</th>
                  <th>Date</th>
                </tr>
              </thead>
              <tbody>
                {
                  this.state.TimeTrackingData.length > 0 && (
                    this.state.TimeTrackingData.map((item) => {
                      return (

                        <tr>
                          <td><div> {item.EmployeeName}</div></td>
                          <td><div>{item.Project}</div></td>
                          <td><div> {item.Hours}</div></td>
                          <td><div>{item.Staus}</div></td>
                          <td><div>{item.TaskDescription}</div></td>
                          <td><p className='line-clamp'>{item.Department}</p></td>
                          <td><div> {item.Date}</div></td>
                        </tr>

                      );
                    })
                  )
                }

              </tbody>
            </table>
          </div>
        </div>
        <Dialog
          hidden={this.state.HideDialog}
          // onDismiss={toggleHideDialog}
          dialogContentProps={dialogContentProps}
          minWidth={400}
        >
          <TextField label="Recipient's Email" onChange={(e) => this.setState({ EmailID: e.target["value"] })} />
          <DialogFooter>
            <PrimaryButton onClick={() => this.triggerFlow(FlowURL.SendMail, this.state.TimeTrackingData)} text="Send" />
            <DefaultButton onClick={() => this.setState({ HideDialog: true })} text="Don't send" />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={this.state.MsgDialog}
          // onDismiss={toggleHideDialog}
          dialogContentProps={dialogContentProps1}
          minWidth={400}
        >
          <DialogFooter>
            <PrimaryButton onClick={() => this.setState({ MsgDialog: true })} text="Ok" />
            {/* <DefaultButton onClick={() => this.setState({ MsgDialog: true })} text="Don't send" /> */}
          </DialogFooter>
        </Dialog>
      </>
    );
  }

  public componentDidMount = async () => {
    await this.GetCurrentUser();
    this.GetTimesheetData();
    await this.GetEmpDetails();
    this.GetChoiceFields();
  }

  public async GetCurrentUser(){
    let user = await sp.web.currentUser.get();
    const UserEmail = user.Email.toLowerCase();

    let userdetail = await sp.web.lists.getByTitle("Employee Information").items.select('Title', 'Status','EmailId','Team').filter(`EmailId eq '${UserEmail}'`).get();
    this.setState({Currentuser : userdetail})

  }

  public GetEmpDetails = async () => {
    // await sp.web.lists.getByTitle("Employee Information").items.select('Title, Status', 'TeamLead/Title').expand("TeamLead").filter("Status eq 1 and TeamLead/Title eq 'Piyush K. Solanki'").get().then((data) => {
   
    if(this.state.Currentuser[0].Team == "Management"){
        await sp.web.lists.getByTitle("Employee Information").items.select('Title, Status', 'TeamLead/Title').expand("TeamLead").filter("Status eq 1 and TeamLead/Title ne null").get().then((data) => {
          let AllEmp = [];

          if (data.length > 0) {
            data.forEach(function (dname, i) {
              AllEmp.push({ key: dname.Title, text: dname.Title });
            });
    
            this.setState({ ChoiceEmployees: AllEmp });
    
          }
          console.log(this.state.ChoiceEmployees);
    
        }).catch((err) => {
          console.log(err);
        });
    }
    else{
      await sp.web.lists.getByTitle("Employee Information").items.select('Title, Status', 'TeamLead/Title','TeamLead/EmailId', 'Team').expand("TeamLead").filter(`Status eq 1 and TeamLead/EmailId eq '${this.state.Currentuser[0].EmailId}'`).get().then((data) => {
        let AllEmp = [];
  
        if (data.length > 0) {
          data.forEach(function (dname, i) {
            AllEmp.push({ key: dname.Title, text: dname.Title });
          });
  
          this.setState({ ChoiceEmployees: AllEmp });
  
        }
        console.log(this.state.ChoiceEmployees);
  
      }).catch((err) => {
        console.log(err);
      });
    }

   
  }

  public GetChoiceFields = async () => {
    const field1: IFieldInfo = await sp.web.lists.getByTitle("Employee Information").fields.getByInternalNameOrTitle("Team")();
    let ProjectStatuslist = [];
    field1["Choices"].forEach(function (dname, i) {
      ProjectStatuslist.push({ key: dname, text: dname });
    });
    this.setState({ ChoiceDepartments: ProjectStatuslist });
    console.log(this.state.ChoiceDepartments);

  }

  // public GetTimesheetData = async () => {

  //   let items = [];
  //   let position = 0;
  //   const pageSize = 2100;
  //   let AllData = [];

  //   try {
  //     while (true) {
  //       let response
  //       // if (this.state.Currentuser[0].Team == "Management") {
  //         response = await sp.web.lists.getByTitle("Employee Timesheet").items.select('EmployeeName/Title', 'TeamLead/Title', 'Title', 'TaskDescription', 'TaskStatus', 'Status', 'Hours', 'StartTime', 'EndTime', 'TaskDescription', 'Department', 'Created').expand('EmployeeName', 'TeamLead').top(pageSize).skip(position).get();
  //       // }
  //       // else {
  //         // response = await sp.web.lists.getByTitle("Employee Timesheet").items.select('EmployeeName/Title', 'TeamLead/Title', 'Title', 'TaskDescription', 'TaskStatus', 'Status', 'Hours', 'StartTime', 'EndTime', 'TaskDescription', 'Department', 'Created').expand('EmployeeName', 'TeamLead').filter(`TeamLead/EmailId eq '${this.state.Currentuser[0].EmailId}'`).top(pageSize).skip(position).get();
  //         response = await sp.web.lists.getByTitle("Employee Timesheet").items.select('EmployeeName/Title', 'TeamLead/Title', 'Title', 'TaskDescription', 'TaskStatus', 'Status', 'Hours', 'StartTime', 'EndTime', 'TaskDescription', 'Department', 'Created').expand('EmployeeName', 'TeamLead').filter(`TeamLead/EmailId eq 'piyushs@200oksolutions.com'`).top(pageSize).skip(position).get();

  //       // }
  //       // const response = await sp.web.lists.getByTitle("Employee Timesheet").items.select('EmployeeName/Title', 'TeamLead/Title', 'Title', 'TaskDescription', 'TaskStatus', 'Status', 'Hours', 'StartTime', 'EndTime', 'TaskDescription', 'Department', 'Created').expand('EmployeeName', 'TeamLead').top(pageSize).skip(position).get();
  //       if (response.length === 0) {
  //         break;
  //       }
  //       items = items.concat(response);
  //       position += pageSize;
  //     }
  //     console.log(`Total items retrieved: ${items.length}`);
  //     if (items.length > 0) {
  //       items.forEach((item) => {
  //         AllData.push({
  //           EmployeeName: item.EmployeeName.Title ? item.EmployeeName.Title : "",
  //           Project: item.Title ? item.Title : "",
  //           Hours: item.Hours ? item.Hours : "",
  //           Staus: item.TaskStatus ? item.TaskStatus : "",
  //           TaskDescription: item.TaskDescription ? item.TaskDescription : "",
  //           Department: item.Department ? item.Department : "",
  //           Date: item.Created ? new Date(new Date(item.Created).getTime() + 5.5*60*60*1000).toISOString().replace('T', ' , ').slice(0, 18) : "",
  //         });
  //       });
  //       this.setState({ TimeTrackingData: AllData },
  //         () => {
  //           $('#myTable').DataTable(
  //             {
  //               order: [[6, 'desc']],
  //             }
  //           );
  //         }
  //       );
  //       console.log(this.state.TimeTrackingData);
  //       this.setState({ FilteredListData: AllData });
  //     }
  //     else {
  //       this.setState({ TimeTrackingData: '' },
  //         () => {
  //           $('#myTable').DataTable(
  //             {
  //               order: [[6, 'desc']],
  //             }
  //           );
  //         }
  //       );
  //     }

  //   } catch (error) {
  //     console.error(error);
  //   }


  //   await sp.web.lists.getByTitle("Employee Timesheet").items.select('EmployeeName/Title', 'TeamLead/Title', 'Title', 'TaskDescription', 'TaskStatus', 'Status', 'Hours', 'StartTime', 'EndTime', 'TaskDescription', 'Department', 'Created').expand('EmployeeName', 'TeamLead').top(pageSize).skip(position).get().then((data) => {
  //     let AllData = [];
  //     console.log(data);

  //     if (data.length > 0) {
  //       data.forEach((item) => {
  //         AllData.push({
  //           EmployeeName: item.EmployeeName.Title ? item.EmployeeName.Title : "",
  //           Project: item.Title ? item.Title : "",
  //           Hours: item.Hours ? item.Hours : "",
  //           Staus: item.TaskStatus ? item.TaskStatus : "",
  //           TaskDescription: item.TaskDescription ? item.TaskDescription : "",
  //           Department: item.Department ? item.Department : "",
  //           Date: item.Created ? item.Created.split("T")[0] : "",
  //         });
  //       });
  //       this.setState({ TimeTrackingData: AllData },
  //         () => {
  //           $('#myTable').DataTable(
  //             {
  //               order: [[6, 'desc']],
  //             }
  //           );
  //         }
  //       );
  //       console.log(this.state.TimeTrackingData);
  //       this.setState({ FilteredListData: AllData });
  //     }
  //     else{
  //       this.setState({ TimeTrackingData: '' },
  //         () => {
  //           $('#myTable').DataTable(
  //             {
  //               order: [[6, 'desc']],
  //             }
  //           );
  //         }
  //       );
  //     }

  //   }).catch((err) => {
  //     console.log(err);

  //   });
  // }

  public GetTimesheetData = async () => {
    let items = [];
    let AllData = [];
    const pageSize = 2100;

    try {
      let response
          if (this.state.Currentuser[0].Team === "Management") {
            response = await sp.web.lists.getByTitle("Employee Timesheet")
              .items
              .select('EmployeeName/Title', 'TeamLead/Title', 'Title', 'TaskDescription', 'TaskStatus', 'Status', 'Hours', 'StartTime', 'EndTime', 'Department', 'Created')
              .expand('EmployeeName', 'TeamLead')
              .top(pageSize)
              .getPaged();
          } else {
            response = await sp.web.lists.getByTitle("Employee Timesheet")
              .items
              .select('EmployeeName/Title', 'TeamLead/Title', 'Title', 'TaskDescription', 'TaskStatus', 'Status', 'Hours', 'StartTime', 'EndTime', 'Department', 'Created')
              .expand('EmployeeName', 'TeamLead')
              .filter(`TeamLead/EmailId eq '${this.state.Currentuser[0].EmailId}'`)
              .top(pageSize)
              .getPaged();
          }

            // response = await sp.web.lists
            //     .getByTitle("Employee Timesheet")
            //     .items.select(
            //         'EmployeeName/Title',
            //         'TeamLead/Title',
            //         'Title',
            //         'TaskDescription',
            //         'TaskStatus',
            //         'Status',
            //         'Hours',
            //         'StartTime',
            //         'EndTime',
            //         'TaskDescription',
            //         'Department',
            //         'Created'
            //     )
            //     .expand('EmployeeName', 'TeamLead')
            //     .filter(`TeamLead/EmailId eq 'piyushs@200oksolutions.com'`)
            //     .top(pageSize)
            //     .getPaged();

            while (response && response.results.length > 0) {
                items = items.concat(response.results);

                if (response.hasNext) {
                    response = await response.getNext();
                } else {
                    break;
                }
            }

            console.log(`Total items retrieved: ${items.length}`);

            if (items.length > 0) {
                items.forEach((item) => {
                    AllData.push({
                        EmployeeName: item.EmployeeName.Title ? item.EmployeeName.Title : "",
                        Project: item.Title ? item.Title : "",
                        Hours: item.Hours ? item.Hours : "",
                        Staus: item.TaskStatus ? item.TaskStatus : "",
                        TaskDescription: item.TaskDescription ? item.TaskDescription : "",
                        Department: item.Department ? item.Department : "",
                        Date: item.Created
                            ? new Date(new Date(item.Created).getTime() + 5.5 * 60 * 60 * 1000)
                                  .toISOString()
                                  .replace('T', ' , ')
                                  .slice(0, 18)
                            : "",
                    });
                });

                this.setState({ TimeTrackingData: AllData }, () => {
                    $('#myTable').DataTable({
                        order: [[6, 'desc']],
                    });
                });

                console.log(this.state.TimeTrackingData);
                this.setState({ FilteredListData: AllData });
            } else {
                this.setState({ TimeTrackingData: '' }, () => {
                    $('#myTable').DataTable({
                        order: [[6, 'desc']],
                    });
                });
            }
        } catch (error) {
            console.error(error);
        }
  };


  private updateDatesForFilter = (startEnd, Date) => {
    if (startEnd == "Start") {
      this.setState({ StartDate: Date });
    }
    else {
      this.setState({ EndDate: Date });
    }
  }

  private filterData = async () => {
    if ((this.state.StartDate && this.state.EndDate) || (this.state.SelectedDepartment || this.state.SelectedEmployeeID)) {
      var tbl = $('#myTable').DataTable();
      tbl.destroy();

      let startDate = moment(this.state.StartDate).format("MM/DD/YYYY");
      let endDate = moment(this.state.EndDate).format("MM/DD/YYYY");
      let SelectedDepartment = this.state.SelectedDepartment;
      let SelectedEmployee = this.state.SelectedEmployeeID;

      let filteredData = this.state.FilteredListData.filter((x) => {
        let ConvertInDate = x.Date.slice(0, 10);
        let Department = x.Department;
        let EmployeeName = x.EmployeeName;
        let crea = moment(ConvertInDate).format("MM/DD/YYYY");
        //return moment(moment(crea)).isSameOrAfter(moment(startDate)) && moment(moment(crea)).isSameOrBefore(moment(endDate)) || (Department) == (SelectedDepartment) || (EmployeeName) == (SelectedEmployee);

        if (this.state.StartDate && this.state.EndDate && this.state.SelectedDepartment && this.state.SelectedEmployeeID) {
          return moment(moment(crea)).isSameOrAfter(moment(startDate)) && moment(moment(crea)).isSameOrBefore(moment(endDate)) && (Department) == (SelectedDepartment) && SelectedEmployee.includes(EmployeeName);
        }
        else if (this.state.StartDate && this.state.EndDate && this.state.SelectedDepartment) {
          return moment(moment(crea)).isSameOrAfter(moment(startDate)) && moment(moment(crea)).isSameOrBefore(moment(endDate)) && (Department) == (SelectedDepartment);
        }
        else if (this.state.StartDate && this.state.EndDate && this.state.SelectedEmployeeID) {
          return moment(moment(crea)).isSameOrAfter(moment(startDate)) && moment(moment(crea)).isSameOrBefore(moment(endDate)) && SelectedEmployee.includes(EmployeeName);
        }
        else if (this.state.SelectedDepartment && this.state.SelectedEmployeeID) {
          return (Department) == (SelectedDepartment) && SelectedEmployee.includes(EmployeeName);
        }
        else if (this.state.StartDate && this.state.EndDate) {
          return moment(moment(crea)).isSameOrAfter(moment(startDate)) && moment(moment(crea)).isSameOrBefore(moment(endDate));
        }
        else if (this.state.SelectedDepartment) {
          return (Department) == (SelectedDepartment);
        }
        else if (this.state.SelectedEmployeeID) {
          // return (EmployeeName) == (SelectedEmployee);
          return SelectedEmployee.includes(EmployeeName);
        }
      });

      this.setState({ TimeTrackingData: filteredData },
        () => {
          $('#myTable').DataTable({
            order: [[6, 'desc']],
          }
          );
        }
      );
    }
    else {
      alert("Please select the appropriate date range or filter option.");
    }
  }

  private resetFilter = async () => {
    this.setState({ StartDate: "", EndDate: "", SelectedDepartment: "", SelectedDepartmentID: "", SelectedEmployee: "", SelectedEmployeeID: "" });
    var tbl = $('#myTable').DataTable();
    tbl.destroy();

    this.setState({ TimeTrackingData: this.state.FilteredListData }, () => {
      $('#myTable').DataTable({
        order: [[6, 'desc']],
      });
    });
  }

  private saveExcel = async () => {
    const workbook = new Excel.Workbook();

    if (this.state.TimeTrackingData.length > 0) {
      try {
        const fileName = 'TimeSheet Details' + "_" + moment().format("DD/MM/YYYY");
        const worksheet = workbook.addWorksheet();

        // add worksheet columns
        // each columns contains header and its mapping key from data
        worksheet.columns = colums;

        // updated the font for first row.
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF002e6d' } // Blue color#002e6d
        };

        // loop through all of the columns and set the alignment with width.
        // worksheet.columns.forEach(column => {
        //   column.width = 20;
        //   column.alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        // });

        worksheet.columns = [
          { width: 25 }, { width: 25 }, { width: 15 }, { width: 20 }, { width: 80 }, { width: 25 }, { width: 20 }

        ];
        worksheet.getColumn(1).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(2).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(3).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(4).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(5).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(6).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };
        worksheet.getColumn(7).alignment = { horizontal: 'left', wrapText: true, vertical: 'middle' };

        const oddRowColor = 'FFFFFF'; // Lighter shade
        const evenRowColor = 'fbfbfb'; // Darker shade
        const borderColor = 'aaaaaa'; // Dark border color FFBFDEF7

        // Loop through data and add each one to worksheet
        this.state.TimeTrackingData.forEach((singleData: any, index: number) => {
          const row = worksheet.addRow(singleData);

          // Set fill color based on odd or even row
          const fillColor = index % 2 === 0 ? { argb: oddRowColor } : { argb: evenRowColor };
          row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: fillColor
          };

          // Set border color for each cell in the row
          row.eachCell((cell, colNumber) => {
            cell.border = {
              top: { style: 'thin', color: { argb: borderColor } },
              left: { style: 'thin', color: { argb: borderColor } },
              bottom: { style: 'thin', color: { argb: borderColor } },
              right: { style: 'thin', color: { argb: borderColor } },
            };
          });
        });

        // write the content using writeBuffer
        const buf = await workbook.xlsx.writeBuffer();

        // download the processed file
        saveAs(new Blob([buf]), `${fileName}.xlsx`);
      } catch (error) {
        console.error('Something Went Wrong', error.message);
      }
    } else {
      alert("Please select the appropriate date range below and click the select button to display the data in the selected date range. To download the chosen data, click the 'download data' option.");
    }
  }


  public triggerFlow = (postURL, data) => {

    if (this.state.EmailID.length > 0) {
      this.setState({ HideDialog: true, EmailID: "" })

      const mail = this.state.EmailID
      const data1 = JSON.stringify({ data, mail })
      const body: string = data1;

      const requestHeaders: Headers = new Headers();
      requestHeaders.append('Content-type', 'application/json');

      const httpClientOptions: IHttpClientOptions = {
        body: body,
        headers: requestHeaders
      };

      return this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, httpClientOptions).then((response) => {
        console.log("Flow Triggered Successfully...");
        this.setState({ MsgDialog: false })
      }).catch(error => {
        console.log(error);
        // swal.close();
      });

    }
    else {
      alert("Please Add the recipient's email address")
    }
  }

  public onItemChange = (event: React.FormEvent<HTMLDivElement>, item: any): void => {

    if (item) {
      this.setState({
        SelectedEmployeeID:
          item.selected
            ? [...this.state.SelectedEmployeeID, item.key as string]
            : this.state.SelectedEmployeeID.filter(key => key !== item.key),
      });
    }

  }
}

