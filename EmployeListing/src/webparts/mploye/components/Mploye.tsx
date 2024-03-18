import * as React from 'react';
import styles from './Mploye.module.scss';
import { IMployeProps } from './IMployeProps';
import { escape } from '@microsoft/sp-lodash-subset';
  import { NormalPeoplePicker, IPersonaProps, IPersona } from '@fluentui/react/lib/';
  import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
  import "@pnp/sp/site-groups/web";
  import { sp } from "@pnp/sp";
  import "@pnp/sp/webs";
  import "@pnp/sp/lists";
  import "@pnp/sp/items";
  import { DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react/lib/DetailsList';
  import { IconButton } from '@fluentui/react/lib/Button';
  import { Dialog, DialogType, DialogFooter, TextField, Dropdown, IDropdownOption, DatePicker,SelectionMode} from '@fluentui/react/lib/';
  import "@pnp/sp/sputilities";
import { IEmailProperties } from "@pnp/sp/sputilities";
import "@pnp/sp/site-users/web";

  export default class Emplisting extends React.Component<IMployeProps, { items: any[],itemID:number, searchQuery: string, isSortedDescending: boolean, showAddDialog: boolean,showEditDialog:boolean, showConfirmation:boolean,deleteitemID: number, newEmployeeName: string, newEmployeeDOB: Date | null, newEmployeeExperience: string, allDepartments: IDropdownOption[], selectedPeople: IPersonaProps[],selectedDepartmentId: React.ReactText,isAdmin:boolean,ManagerId:number }> {
    constructor(props: IMployeProps) {
      super(props);
      this.onPeoplePickerChange = this.onPeoplePickerChange.bind(this);
      sp.setup({
        spfxContext: this.props.spfxContext,    
      });
      this.state = {
        ManagerId:null,
        isAdmin:false,
        items: [],
        itemID: null, 
        selectedPeople: [],
        searchQuery: '',
        isSortedDescending: false,
        showAddDialog: false,
        showEditDialog: false,
        showConfirmation:false,
        deleteitemID: 0,
        newEmployeeName: '',
        newEmployeeDOB: null,
        newEmployeeExperience: '',
        allDepartments: [],
        selectedDepartmentId: ''
      };
    }

    public render(): React.ReactElement<IMployeProps> {
      const columns: IColumn[] = [
        { key: 'column0', name: 'Actions', fieldName: 'actions', minWidth: 80, maxWidth: 80, isResizable: true, onRender: (item) => this.renderActionsColumn(item) },
        { key: 'column1', name: 'Name', fieldName: 'Name', minWidth: 80, maxWidth: 80, isResizable: true, isSorted: true, isSortedDescending: this.state.isSortedDescending, onColumnClick: this.handleColumnClick },
        { key: 'column2', name: 'DOB', fieldName: 'DOB', minWidth: 120, maxWidth: 120, isResizable: true },
        { key: 'column3', name: 'Experience', fieldName: 'Experience', minWidth: 80, maxWidth: 80, isResizable: true },
        { key: 'column4', name: 'DepartmentName', fieldName: 'DepartmentName', minWidth: 120, maxWidth: 120, isResizable: true },
        { key: 'column5', name: 'Manager', fieldName: 'Manager', minWidth: 80, maxWidth: 80, isResizable: true }
      ];

      const { items, itemID, searchQuery,showConfirmation, showAddDialog,showEditDialog, newEmployeeName, newEmployeeDOB, newEmployeeExperience, allDepartments,selectedPeople,isAdmin,ManagerId } = this.state;

      return (
        <div>
          <div>
            <input type="text" value={searchQuery} onChange={this.handleSearchInputChange} />
            <button onClick={this.handleSearch}>Search</button>
            {/* <button onClick={this.openAddEmployeeDialog}>Add Employee</button> */}
            {isAdmin ? (
            <button onClick={this.openAddEmployeeDialog}>Add Employee</button>
            ) : (
              null
            )}
          </div>
          <DetailsList
            items={items}
            columns={columns}
            selectionMode={SelectionMode.none} 
    
            layoutMode={DetailsListLayoutMode.fixedColumns}
            // selectionPreservedOnEmptyClick={true }
          />
          <Dialog
            hidden={!showAddDialog}
            onDismiss={this.closeAddEmployeeDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Add New Employee',
              closeButtonAriaLabel: 'Close'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { maxWidth: 450 } }
            }}
          >
            <TextField label="Name" value={newEmployeeName} onChange={this.handleNewEmployeeNameChange} />
            <DatePicker
              label="DOB"
              value={newEmployeeDOB}
              onSelectDate={this.handleNewEmployeeDOBChange}
              allowTextInput={true}
            />
            <TextField label="Experience" value={newEmployeeExperience} onChange={this.handleNewEmployeeExperienceChange} />
            <Dropdown
              label="Department"
              options={allDepartments}
              onChange={this.handleNewEmployeeDepartmentChange}
              selectedKey={this.state.selectedDepartmentId}
              placeholder="Select department"
            />
            
            <PeoplePicker
                  context={this.props.spfxContext}
                  personSelectionLimit={1}
                  titleText="Project Manager"
                  groupName={""} // Leave this blank in case you want to filter from all users
                  placeholder="Enter a name or email address"
                  onChange={this.onPeoplePickerChange}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                  required={true}
                />

                
            <DialogFooter>
              <button className="styles.button" onClick={this.addEmployee}>Add</button>
              <button className="styles.button" onClick={this.closeAddEmployeeDialog}>Cancel</button>
            </DialogFooter>
          </Dialog>

              
          <Dialog
            hidden={!showEditDialog}
            onDismiss={this.closeEditEmployeeDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Edit Employee',
              closeButtonAriaLabel: 'Close'
            }}
            modalProps={{
              isBlocking: false,
              styles: { main: { maxWidth: 450 } }
            }}
          >
            <TextField label="Name" value={newEmployeeName} onChange={this.handleNewEmployeeNameChange} />
            <DatePicker
              label="DOB"
              value={newEmployeeDOB}
              onSelectDate={this.handleNewEmployeeDOBChange}
              allowTextInput={true}
            />
            <TextField label="Experience" value={newEmployeeExperience} onChange={this.handleNewEmployeeExperienceChange} />
            <Dropdown
              label="Department"
              options={allDepartments}
              onChange={this.handleNewEmployeeDepartmentChange}
              selectedKey={this.state.selectedDepartmentId}
              placeholder="select Department"
            />
            
            <PeoplePicker
                  context={this.props.spfxContext}
                  personSelectionLimit={1}
                  titleText="Project Manager"
                  groupName={""} // Leave this blank in case you want to filter from all users
                  placeholder="Enter a name or email address"
                  onChange={this.onPeoplePickerChange}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                  required={true}
                />

            <DialogFooter>
              <button onClick={this.openConfirmation}>Edit</button>
              <button onClick={this.closeEditEmployeeDialog}>Cancel</button>
            </DialogFooter>
          </Dialog>

          <Dialog
          hidden={!showConfirmation}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Are you sure you waant to update details ? ',
            closeButtonAriaLabel: 'Close'
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: { maxWidth: 450 } }
          }}
          >
            <DialogFooter>
              <button onClick={this.saveEditedData}>Yes</button>
              <button onClick={this.closeConfirmation}>Cancle</button>
            </DialogFooter>
          </Dialog>
        </div>
      );
    }

    public componentDidMount = () => {
      this.getListItems();
    }
    
    public onPeoplePickerChange(selected) {
      this.setState({ selectedPeople : selected });
      console.log('selected people id at selection :',selected);
    }  

    private getListItems = async () => {
      try {
        console.log('---------------------------------------------------------------------------------------');
        console.log('in the get list');
        const { selectedPeople }=this.state;

        console.log('get list selected people ' ,selectedPeople);
        const empdata = await sp.web.lists.getByTitle("Emplist").items.select("ID", "Name", "DOB", "Experience", "DepartmentNId","Manager/Title","Manager/Id","Manager/UserName","Manager/EMail").expand('Manager').getAll();
        console.log('empdata get list : ',empdata);

        const manager=await sp.web.lists.getByTitle("Emplist").items.select("Manager/Title","Manager/Id").expand('Manager').get();
        console.log('Manager :',manager);

        const depdata = await sp.web.lists.getByTitle("Department").items.select('Id', 'DepartmentN').getAll();

        const mergedData = empdata.map((employee,index) => {

          console.log(employee);
          console.log(index);

          const commondata = depdata.filter(dep => employee.DepartmentNId === dep.Id);
          const departmentName = commondata.map(data=>data.DepartmentN);  
          
          return {
            Name: employee.Name,
            DOB: employee.DOB,
            Experience: employee.Experience,
            DepartmentName: departmentName,
            ID: employee.ID,
            Manager:employee.Manager.Title,
          };
        });

        console.log('mereged data :',mergedData);
        this.setState({ items: mergedData });
        const departmentOptions = depdata.map(dep => ({
          key: dep.Id,
          text: dep.DepartmentN
        }));
        this.setState({ allDepartments: departmentOptions });
      } catch (error) {
        console.log("error line 245 : ",error);
      }
    // console.log("Email Sent!");
    let currentuser = await sp.web.currentUser();
    // console.log('current userr :',currentuser)
    const groups = await sp.web.siteGroups();
    // console.log('groups :',groups);
    const groupName = "Agroup";
    let grp = await sp.web.siteGroups.getByName(groupName)();
    // console.log("Admin group :",grp)
    let grpusers = await sp.web.siteGroups.getByName(groupName).users();
    // console.log("Admin group users :",grpusers)
    const isAdmin = grpusers.some(user=>user.Email==currentuser.Email);
    if(isAdmin){
      // console.log('yes you are admin!')
      this.setState({ isAdmin : true });
    }    
   }
    
    private handleSearchInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
      this.setState({ searchQuery: event.target.value });
    }
    
    private handleSearch = () => {
      const { items, searchQuery } = this.state;
      if (searchQuery.trim() === '') {
        this.getListItems();
      } else {
        const filteredItems = items.filter(item => item.Name.toLowerCase().includes(searchQuery.toLowerCase()));
        this.setState({ items: filteredItems });
      }
    }

  private saveEditedData = async (event: React.MouseEvent<HTMLButtonElement, MouseEvent>) => {
    // console.log('clicked');
    // console.log(item.ID)
    await this.editEmployee(this.state.itemID);
  }

    private handleColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
      const { items, isSortedDescending } = this.state;
      const sortedItems = [...items].sort((a, b) => {
        if (column.key === 'column1') { // Only sort if the clicked column is 'Name'
          if (isSortedDescending) {
            return a[column.fieldName].toLowerCase() > b[column.fieldName].toLowerCase() ? 1 : -1;
          } else {
            return a[column.fieldName].toLowerCase() < b[column.fieldName].toLowerCase() ? 1 : -1;
          }
        } else {
          return isSortedDescending ? b[column.fieldName] > a[column.fieldName] ? 1 : -1 : a[column.fieldName] > b[column.fieldName] ? 1 : -1;
        }
      });

      this.setState({ items: sortedItems, isSortedDescending: !isSortedDescending });
    }

    private renderActionsColumn = (item: any): JSX.Element => {
      return (
        <>
        <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this.deleteItem(item.ID)}/>
        <IconButton iconProps={{ iconName: 'Edit' }} onClick={()=>this.openEditEmployeeDialog(item.ID)}/>
        </>
      );
    }

    private openAddEmployeeDialog = () => {
      this.setState({ showAddDialog: true });
    }

    private openConfirmation = () => {
      this.setState({ showConfirmation: true });
    }


    private closeConfirmation = () => {
      this.setState({ showConfirmation: false });
    }

    private openEditEmployeeDialog = async(itemID) => {

      console.log("at opening of edit dialog : ",itemID);
      this.setState({ showEditDialog: true,itemID:itemID });
      try{
        const existingItem = await sp.web.lists.getByTitle("Emplist").items.getById(itemID).get();
        console.log(existingItem);
        this.setState({
          newEmployeeName: existingItem.Name,
          newEmployeeDOB: null,
          newEmployeeExperience: existingItem.Experience,
          selectedDepartmentId: existingItem.DepartmentNId,
          ManagerId:existingItem.ManagerId,
        });
      }
      catch(err){
        console.log(err);
      }
    }

    private closeAddEmployeeDialog = () => {
      this.setState({
        newEmployeeName: '',
        newEmployeeDOB: null,
        newEmployeeExperience: '',
        selectedDepartmentId: '',
        showAddDialog: false
    });

    }

    private closeEditEmployeeDialog = () => {
      this.setState({ showEditDialog: false });
    }

    private handleNewEmployeeNameChange = (event: React.FormEvent<HTMLInputElement>, newValue?: string) => {
      this.setState({ newEmployeeName: newValue || '' });
    }

    private handleNewEmployeeDOBChange = (date: Date | null | undefined) => {
      this.setState({ newEmployeeDOB: date || null });
    }

    private handleNewEmployeeExperienceChange = (event: React.FormEvent<HTMLInputElement>, newValue?: string) => {
      this.setState({ newEmployeeExperience: newValue || '' });
    }

    private handleNewEmployeeDepartmentChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
      if (option) {
        this.setState({ selectedDepartmentId: option.key }); // Update selected department id
      }
    }

    private addEmployee = async () => {
    
      const { newEmployeeName, newEmployeeDOB, newEmployeeExperience,ManagerId,selectedPeople } = this.state;
      console.log('selected people at add employee : ',selectedPeople); 
      
      console.log(newEmployeeName, newEmployeeDOB, newEmployeeExperience);
      try {
        await sp.web.lists.getByTitle("Emplist").items.add({
          Name: newEmployeeName,
          DOB: newEmployeeDOB,
          Experience: newEmployeeExperience,
          DepartmentNId: this.state.selectedDepartmentId,
          ManagerId:this.state.selectedPeople[0].id,
        });   
         console.log('item added');
        this.getListItems(); 
        this.closeAddEmployeeDialog();

        this.setState({
          newEmployeeName: '',  
          newEmployeeDOB: null,
          newEmployeeExperience: '',
          selectedDepartmentId: '',
        });
        const emailProps: IEmailProperties = {
          To:["meet@DesireInfoweb74.onmicrosoft.com"],
          Subject: "This email is about the Sharepoint task",
          Body: "Hello. <b>data added</b>",
          AdditionalHeaders: {
              "content-type": "text/html"
          }
      };
      
      await sp.utility.sendEmail(emailProps);
      console.log("Email has been sent successfully!");
      } catch (error) {
        console.log("error line 408 : ",error);
      }
    }

  
    private editEmployee = async (itemID: number) => {
      try {
          const existingItem = await sp.web.lists.getByTitle("Emplist").items.getById(itemID).get();
          const { newEmployeeName, newEmployeeDOB, newEmployeeExperience, selectedDepartmentId } = this.state;
          // Update the existing item with new values
          await sp.web.lists.getByTitle("Emplist").items.getById(itemID).update({
              Name: newEmployeeName || existingItem.Name,
              DOB: newEmployeeDOB || existingItem.DOB,
              Experience: newEmployeeExperience || existingItem.Experience,
              DepartmentNId: selectedDepartmentId || existingItem.DepartmentNId,
              ManagerId:this.state.selectedPeople[0].id || existingItem.Manager.Id,
          });

          // Refresh the list and close the edit dialog
          this.getListItems();
          this.closeEditEmployeeDialog();
          this.closeConfirmation();
          // Clear the form fields
          this.setState({
              newEmployeeName: '',
              newEmployeeDOB: null,
              newEmployeeExperience: '',
              selectedDepartmentId: '',
              ManagerId:null,
          });
      } catch (error) {
          console.log("error line 434 :",error);
      }
  }

    private deleteItem = async (itemID: number) => {
      try {
        await sp.web.lists.getByTitle("Emplist").items.getById(itemID).delete();
        this.getListItems();
      } catch (error) {
        console.log("error line 443 : ",error); 
      }
    }
  }