import * as React from 'react';

import { TagPicker, ITag } from 'office-ui-fabric-react/lib/Pickers';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

initializeIcons();

export interface ITagPickerProps {
  selectedItems?: ITag[],
  onChange?: (items?: ITag[]) => void;
  onEmptyInputFocus?: (selectedItems?: ITag[]) => Promise<ITag[]>;
	onResolveSuggestions?: (filter: string, selectedItems?: ITag[]) => Promise<ITag[]>;
}

export interface ITagPickerState extends React.ComponentState, ITagPickerProps {
}

export class TagPickerBase extends React.Component<ITagPickerProps, ITagPickerState> {
  constructor(props: ITagPickerProps) {
    super(props);

    this.state = {
      selectedItems: props.selectedItems || []
    };
  }

  public componentWillReceiveProps(newProps: ITagPickerState): void {
    this.setState(newProps);
  }

  public render(): JSX.Element {
    const { selectedItems } = this.state;

    return (
      <div className={"tagPickerComponent"}>
        <TagPicker
          removeButtonAriaLabel="Remove"
          selectedItems={selectedItems}
          onChange={this._onChange}
          onResolveSuggestions={this._onResolveSuggestions}
          onEmptyInputFocus={this._onEmptyInputFocus}
          getTextFromItem={this._getTextFromItem}
          pickerSuggestionsProps={{
            suggestionsHeaderText: 'Suggested Tags',
            noResultsFoundText: 'No Tags Found'
          }}
          resolveDelay={300}
          inputProps={{
            'aria-label': 'Tag Picker'
          }}
        />
      </div>
    );
  }

  private _getTextFromItem(item: ITag): string {
    return item.name;
  }

  private _onChange = (items?: ITag[]): void => {
    this.setState(
      (prevState: ITagPickerState): ITagPickerState => {
        prevState.selectedItems = items;
        return prevState;
      }
    );

    if (this.props.onChange)
      this.props.onChange(items);
  }

  private _onResolveSuggestions = (filter: string,  selectedItems?: ITag[] | undefined): Promise<ITag[]> => {
    if (this.props.onResolveSuggestions)
      return this.props.onResolveSuggestions(filter, selectedItems);

    return Promise.resolve([]);
  };

  private _onEmptyInputFocus = (selectedItems?: ITag[] | undefined): Promise<ITag[]> => {
    if (this.props.onEmptyInputFocus)
      return this.props.onEmptyInputFocus(selectedItems);

    return Promise.resolve([]);
  };
}