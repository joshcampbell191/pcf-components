import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { ITag } from 'office-ui-fabric-react/lib/Pickers';
import { TagPickerBase, ITagPickerProps } from './TagPicker';
import 'whatwg-fetch';

declare const Xrm: any;

// https://docs.microsoft.com/en-us/powerapps/developer/common-data-service/entity-metadata
enum EntityMetadataProperties {
	EntitySetName = "EntitySetName",
	PrimaryIdAttribute = "PrimaryIdAttribute",
	PrimaryNameAttribute  = "PrimaryNameAttribute"
}

export class TagPickerComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private context: ComponentFramework.Context<IInputs>;
	private notifyOutputChanged: () => void;
	private theContainer: HTMLDivElement;

	private selectedItems: ITag[];

	private props: ITagPickerProps = {
		onChange: this.onChange.bind(this),
		onEmptyInputFocus: this.onEmptyInputFocus.bind(this),
		onResolveSuggestions: this.onResolveSuggestions.bind(this)
	}

	private relatedEntity: string;
	private relationshipEntity: string;
	private relationshipName: string;

	private entityMetadata: ComponentFramework.PropertyHelper.EntityMetadata;
	private relatedEntityMetadata: ComponentFramework.PropertyHelper.EntityMetadata;

	private entityId?: string;
	private entityType: string;

	private get idAttribute(): string { return this.relatedEntityMetadata ? this.relatedEntityMetadata[EntityMetadataProperties.PrimaryIdAttribute] : ""; }
	private get nameAttribute(): string { return this.relatedEntityMetadata ? this.relatedEntityMetadata[EntityMetadataProperties.PrimaryNameAttribute] : ""; }

	/**
	 * Empty constructor.
	 */
	constructor()
	{
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this.context = context;
		this.notifyOutputChanged = notifyOutputChanged;
		this.props.selectedItems = this.selectedItems = [];
		this.theContainer = container;

		this.relatedEntity = this.context.parameters.relatedEntity.raw || "";
		this.relationshipEntity = this.context.parameters.relationshipEntity.raw || "";
		this.relationshipName = this.context.parameters.relationshipName.raw || "";

		this.entityId = (<any>this.context).page.entityId;
		this.entityType =  (<any>this.context).page.entityTypeName;

		this.loadMetadata().then(() => {
			return this.getRelatedEntities();
		})
		.then(entities => {
			return this.getTags(entities);
		})
		.then(tags => {
			this.props.selectedItems = this.selectedItems = tags;
			this.updateView(context);
		});
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		if (context.updatedProperties.includes("tags"))
			console.log("tags", context.parameters.tags.raw);

		ReactDOM.render(
			React.createElement(
				TagPickerBase,
				this.props
			),
			this.theContainer
		);
	}

	private loadMetadata(): Promise<ComponentFramework.PropertyHelper.EntityMetadata> {
		const entityName: string = (<any>this.context).page.entityTypeName;
		return Promise.all([
			this.context.utils.getEntityMetadata(entityName).then(value => this.entityMetadata = value),
			this.context.utils.getEntityMetadata(this.relatedEntity).then(value => this.relatedEntityMetadata = value)
		]);
	}

	private getRelatedEntities(): Promise<ComponentFramework.WebApi.Entity[]> {
		const options = `?$filter=${this.entityType}id eq ${this.entityId}`;
		return this.context.webAPI.retrieveMultipleRecords(this.relationshipEntity, options).then(
			results => { return results.entities; }
		);
	}

	private getTags(entities: ComponentFramework.WebApi.Entity[]): Promise<ITag[]> {
		if (entities.length < 1) {
			return Promise.resolve([]);
		}

		const promises = [];
		for(let entity of entities) {
			const relatedEntityId = entity[this.idAttribute];
			const options = `?$select=${this.idAttribute},${this.nameAttribute}`;
			promises.push(this.context.webAPI.retrieveRecord(this.relatedEntity, relatedEntityId, options));
		}

		return Promise.all(promises).then(
			results => {
				return results!.map(result => ({ key: result[this.idAttribute], name: result[this.nameAttribute] }));
			}
		);
	}

	private onEmptyInputFocus(selectedItems?: ITag[]): Promise<ITag[]> {
		return this.searchTags();
	}

	private onResolveSuggestions(filter: string, selectedItems?: ITag[]): Promise<ITag[]> {
		return this.searchTags(filter);
	}

	private searchTags(filter?: string): Promise<ITag[]> {
		let options = `?$select=${this.idAttribute},${this.nameAttribute}&$orderby=${this.nameAttribute} asc`;

		if (filter)
			options = `${options}&$filter=contains(${this.nameAttribute},'${filter}')`;

		return this.context.webAPI.retrieveMultipleRecords(this.relatedEntity, options).then(
			results => {
				if (results.entities.length < 1)
					return [];

				return results.entities.map(item => ({ key: item[this.idAttribute], name: item[this.nameAttribute] }));
			}
		);
	}

	private onChange(items?: ITag[]) : void {
		const promises: Promise<Response>[] = [];

		const itemsAdded = items?.filter(item => !this.selectedItems.some(selectedItem => selectedItem.key === item.key)) || [];
		for(let item of itemsAdded) {
			promises.push(this.associateItem(item));
		}

		const itemsRemoved = this.selectedItems.filter(selectedItem => !items?.some(item => item.key === selectedItem.key));
		for (let item of itemsRemoved) {
			promises.push(this.dissasociateItem(item));
		}

		Promise.all(promises).then(
			results => {
				this.props.selectedItems = this.selectedItems = items || [];
				this.notifyOutputChanged();
			}
		);
	}

	private associateItem(item: ITag): Promise<Response> {
		const clientUrl: string = (<any>Xrm).Utility.getGlobalContext().getClientUrl();

		const entityCollectionName = this.entityMetadata[EntityMetadataProperties.EntitySetName];
		const entityId: string = (<any>this.context).page.entityId;
		const payload = { "@odata.id" : `${clientUrl}/api/data/v9.1/${entityCollectionName}(${entityId})` };

		const relatedEntityCollectionName: string = this.relatedEntityMetadata[EntityMetadataProperties.EntitySetName];

		return window.fetch(`${clientUrl}/api/data/v9.1/${relatedEntityCollectionName}(${item.key})/${this.relationshipName}/$ref`, {
			method: "POST",
			headers: {
				"Content-Type": "application/json; charset=utf-8",
				"Accept": "application/json",
				"OData-MaxVersion": "4.0",
				"OData-Version": "4.0"
			},
			body: JSON.stringify(payload)
		});
	}

	private dissasociateItem(item: ITag): Promise<Response> {
		const clientUrl: string = (<any>Xrm).Utility.getGlobalContext().getClientUrl();

		const entityCollectionName = this.entityMetadata[EntityMetadataProperties.EntitySetName];
		const entityId: string = (<any>this.context).page.entityId;

		return window.fetch(`${clientUrl}/api/data/v9.1/${entityCollectionName}(${entityId})/${this.relationshipName}(${item.key})/$ref`, {
			method: "DELETE",
			headers: {
				"Content-Type": "application/json; charset=utf-8",
				"Accept": "application/json",
				"OData-MaxVersion": "4.0",
				"OData-Version": "4.0"
			}
		});
	}

	/**
	 * It is called by the framework prior to a control receiving new data.
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {
			tags: this.props.selectedItems!.map(items => items.key).join(",")
		};
	}

	/**
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		ReactDOM.unmountComponentAtNode(this.theContainer);
	}
}