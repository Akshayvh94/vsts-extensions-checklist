import "../css/App.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { Fabric } from "OfficeFabric/Fabric";
import { Checkbox } from 'OfficeFabric/Checkbox';
import { IconButton } from "OfficeFabric/components/Button";
import { TextField } from "OfficeFabric/TextField";
import { autobind } from "OfficeFabric/Utilities";
import {Icon} from "OfficeFabric/Icon";
import {Pivot, PivotItem} from "OfficeFabric/Pivot";
import { MessageBar, MessageBarType } from 'OfficeFabric/MessageBar';

import * as WitExtensionContracts  from "TFS/WorkItemTracking/ExtensionContracts";
import { WorkItemFormService, IWorkItemFormService } from "TFS/WorkItemTracking/Services";
import * as Utils_Array from "VSS/Utils/Array";
import * as Utils_String from "VSS/Utils/String";

import {AutoResizableComponent} from "VSTS_Extension/Components/WorkItemControls/AutoResizableComponent";
import {ExtensionDataManager} from "VSTS_Extension/Utilities/ExtensionDataManager";
import {Loading} from "VSTS_Extension/Components/Common/Loading";
import {InputError} from "VSTS_Extension/Components/Common/InputError";

interface IChecklistProps {
}

interface IChecklistState {
    privateDataModel: IExtensionDataModel;
    sharedDataModel: IExtensionDataModel;
    isLoaded: boolean;
    isNewWorkItem: boolean;
    itemText: string;
    inputError: string;
    saveError: boolean;
    isPersonalView: boolean;
}

interface IChecklistItem {
    id: string;
    text: string;
    checked: boolean;
}

interface IExtensionDataModel {
    id: string;
    __etag: number;
    items: IChecklistItem[];
}

export class Checklist extends AutoResizableComponent<IChecklistProps, IChecklistState> {
    constructor(props: IChecklistProps, context?: any) {
        super(props, context);

        VSS.register(VSS.getContribution().id, {
            onLoaded: (args: WitExtensionContracts.IWorkItemLoadedArgs) => {
                this._refreshItems(false);
            },
            onUnloaded: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
                this.setState({...this.state, items: []});
            },
            onRefreshed: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
                this._refreshItems(true);
            },
            onSaved: (args: WitExtensionContracts.IWorkItemChangedArgs) => {
                if (this.state.isNewWorkItem) {
                    this._refreshItems(true);
                }                
            }
        } as WitExtensionContracts.IWorkItemNotificationListener);

        this.state = {
            privateDataModel: null,
            sharedDataModel: null,
            isLoaded: false,
            isNewWorkItem: false,
            itemText: "",
            inputError: "",
            saveError: false,
            isPersonalView: true
        }
    }

    public render(): JSX.Element {
        if (!this.state.isLoaded) {
            return <Loading />;
        }
        else if(this.state.isNewWorkItem) {
            return <MessageBar messageBarType={MessageBarType.info}>You need to save the workitem before working with checklist.</MessageBar>;
        }
        else {
            let currentModel = this.state.isPersonalView ? this.state.privateDataModel : this.state.sharedDataModel;

            return (
                <Fabric className="fabric-container">
                    <div className="container">
                        <Pivot initialSelectedIndex={this.state.isPersonalView ? 0 : 1} onLinkClick={this._onPivotChange}>
                            <PivotItem linkText="Personal" itemKey="personal" />
                            <PivotItem linkText="Shared" itemKey="shared" />
                        </Pivot>
                        
                        { 
                            this.state.saveError && 
                            (
                                <MessageBar messageBarType={MessageBarType.error}>
                                    The current version of checklist doesn't match the version of checklist in this workitem. Please refresh the workitem to get the latest Checklist data.
                                </MessageBar>
                            )
                        }

                        <div className="checklist-items">
                            { 
                                (currentModel == null || currentModel.items == null || currentModel.items.length == 0)
                                && 
                                <MessageBar messageBarType={MessageBarType.info}>
                                    No checklist items yet.
                                </MessageBar>
                            }
                            { 
                                (currentModel != null && currentModel.items != null && currentModel.items.length > 0) 
                                && 
                                this._renderCheckListItems(currentModel.items)
                            }
                            
                        </div>
                        <div className="add-checklist-items">
                            <IconButton className="add-icon" iconProps={{iconName: "Add"}} title="Add item" onClick={this._onAddListItem} />
                            <input
                                type="text" 
                                value={this.state.itemText}
                                onChange={this._onItemTextChange} 
                                onKeyUp={e => {
                                    if (e.keyCode === 13) {
                                        this._onAddListItem();
                                    }                                    
                                }}
                                />                            
                        </div>
                        { this.state.inputError && (<InputError error={this.state.inputError} />) }
                    </div>                    
                </Fabric>
            );
        }
    }

    @autobind
    private _onPivotChange(item: PivotItem) {
        this.setState({...this.state, isPersonalView: item.props.itemKey === "personal", itemText: "", inputError: ""});
    }

    @autobind
    private _onItemTextChange(e: React.ChangeEvent<HTMLInputElement>) {        
        this.setState({...this.state, itemText: e.target.value, inputError: this._getItemTextError(e.target.value)});
    }

    @autobind
    private _getItemTextError(value: string): string {
        if (value.length > 128) {
            return `The length of the title should less than 128 characters, actual is ${value.length}.`
        }
        return "";
    }

    @autobind
    private async _onAddListItem() {
        if (this.state.itemText && this.state.itemText.trim()) {
            const workItemFormService = await WorkItemFormService.getService();
            const workItemId = await workItemFormService.getId();
            let newModel: IExtensionDataModel;

            if (this.state.isPersonalView) {
                newModel = this.state.privateDataModel ? {...this.state.privateDataModel} : {id: `${workItemId}`, __etag: 0, items: []};
            }
            else {
                newModel = this.state.sharedDataModel ? {...this.state.sharedDataModel} : {id: `${workItemId}`, __etag: 0, items: []};
            }

            newModel.items = (newModel.items || []).concat({id: `${Date.now()}`, text: this.state.itemText, checked: false});

            try {
                newModel = await ExtensionDataManager.writeDocument<IExtensionDataModel>("CheckListItems", newModel, this.state.isPersonalView);

                if (this.state.isPersonalView) {
                    this.setState({...this.state, itemText: "", inputError: "", privateDataModel: newModel, saveError: false});
                }
                else {
                    this.setState({...this.state, itemText: "", inputError: "", sharedDataModel: newModel, saveError: false});
                }
            }
            catch (e) {
                this.setState({...this.state, saveError: true});
            }        
        }        
    }    

    private _renderCheckListItems(items: IChecklistItem[]): React.ReactNode {
        return items.map((item: IChecklistItem, index: number) => {
            return (
                <div className="checklist-item" key={`${index}`}>
                    <Checkbox 
                        className="checkbox"
                        label={item.text}
                        checked={item.checked}
                        onChange={(ev: React.FormEvent<HTMLElement>, isChecked: boolean) => this._onCheckboxChange(item.id, isChecked) } />         

                    <IconButton className="delete-item-button" iconProps={{iconName: "Delete"}} title="Delete item" onClick={() => this._onDeleteItem(item.id)} />
                </div>
            );
        });
    }

    @autobind
    private async _onDeleteItem(itemId: string) {
        let currentModel = this.state.isPersonalView ? this.state.privateDataModel : this.state.sharedDataModel;

        let newModel =  {...currentModel};
        Utils_Array.removeWhere(newModel.items, (item: IChecklistItem) => Utils_String.equals(item.id, itemId, true));

        try {
            newModel = await ExtensionDataManager.writeDocument<IExtensionDataModel>("CheckListItems", newModel, this.state.isPersonalView);

            if (this.state.isPersonalView) {
                this.setState({...this.state, privateDataModel: newModel, saveError: false});
            }
            else {
                this.setState({...this.state, sharedDataModel: newModel, saveError: false});
            }
        }
        catch (e) {
            this.setState({...this.state, saveError: true});
        } 
    }

    @autobind
    private async _onCheckboxChange(itemId: string, isChecked: boolean) {
        let currentModel = this.state.isPersonalView ? this.state.privateDataModel : this.state.sharedDataModel;

        let newModel =  {...currentModel};
        let index = Utils_Array.findIndex(newModel.items, (item: IChecklistItem) => Utils_String.equals(item.id, itemId, true));
        if (index !== -1) {
            newModel.items[index].checked = isChecked;
        }

        try {
            newModel = await ExtensionDataManager.writeDocument<IExtensionDataModel>("CheckListItems", newModel, this.state.isPersonalView);

            if (this.state.isPersonalView) {
                this.setState({...this.state, privateDataModel: newModel, saveError: false});
            }
            else {
                this.setState({...this.state, sharedDataModel: newModel, saveError: false});
            }
        }
        catch (e) {
            this.setState({...this.state, saveError: true});
        }
    }

    private async _refreshItems(hotReset: boolean) {
        if (!hotReset) {
            this.setState({...this.state, privateDataModel: null, sharedDataModel: null, isLoaded: false, isNewWorkItem: false});
        }        

        const workItemFormService = await WorkItemFormService.getService();
        const isNew = await workItemFormService.isNew();

        if (isNew) {
            this.setState({...this.state, privateDataModel: null, sharedDataModel: null, isLoaded: true, error: "", saveError: false, itemText: "", isNewWorkItem: true});
        }
        else {
            const workItemId = await workItemFormService.getId();
            let models: IExtensionDataModel[] = await Promise.all([
                ExtensionDataManager.readDocument<IExtensionDataModel>("CheckListItems", `${workItemId}`, null, true),
                ExtensionDataManager.readDocument<IExtensionDataModel>("CheckListItems", `${workItemId}`, null, false)
            ]);

            this.setState({...this.state, privateDataModel: models[0], sharedDataModel: models[1], isLoaded: true, isNewWorkItem: false, error: "", saveError: false, itemText: ""});
        }
    }
}

export function init() {
    ReactDOM.render(<Checklist />, $("#ext-container")[0]);
}