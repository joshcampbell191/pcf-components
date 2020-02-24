import * as React from 'react';

import { TagPicker, IBasePicker, ITag } from 'office-ui-fabric-react/lib/Pickers';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';

initializeIcons(/* optional base url */);

const rootClass = mergeStyles({
  maxWidth: 500
});

const _testTags: ITag[] = [
  'black',
  'blue',
  'brown',
  'cyan',
  'green',
  'magenta',
  'mauve',
  'orange',
  'pink',
  'purple',
  'red',
  'rose',
  'violet',
  'white',
  'yellow'
].map(item => ({ key: item, name: item }));

export interface ITagPickerDemoPageProps {
  tags?: string;  
	resolveSuggestions: (filter: string) => Promise<ITag[]>;
}

export interface ITagPickerDemoPageState {
  isPickerDisabled?: boolean;
}

export class TagPickerBasicExample extends React.Component<ITagPickerDemoPageProps, ITagPickerDemoPageState> {
  // All pickers extend from BasePicker specifying the item type.
  private _picker = React.createRef<IBasePicker<ITag>>();

  constructor(props: ITagPickerDemoPageProps) {
    super(props);

    this.state = {
      isPickerDisabled: false
    };
  }

  public render() {
    return (
      <div className={rootClass}>        
        Filter items on selected: This picker will show already-added suggestions but will not add duplicate tags.
        <TagPicker
          removeButtonAriaLabel="Remove"
          componentRef={this._picker}
          onResolveSuggestions={this._onFilterChangedNoFilter}
          onItemSelected={this._onItemSelected}
          getTextFromItem={this._getTextFromItem}
          pickerSuggestionsProps={{
            suggestionsHeaderText: 'Suggested Tags',
            noResultsFoundText: 'No Color Tags Found'
          }}
          resolveDelay={300}
          disabled={this.state.isPickerDisabled}
          inputProps={{
            onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
            onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called'),
            'aria-label': 'Tag Picker'
          }}
        />
      </div>
    );
  }

  private _getTextFromItem(item: ITag): string {
    return item.name;
  }

  private _onFilterChangedNoFilter = (filter: string,  selectedItems?: ITag[] | undefined): Promise<ITag[]> => {
    return this.props.resolveSuggestions(filter);
  };

  private _onItemSelected = (selectedItem?: ITag | undefined): ITag | null => {
    if (this._picker.current && this._listContainsDocument(selectedItem, this._picker.current.items)) {
      return null;
    }
    return selectedItem ?? null;
  };

  private _listContainsDocument(tag?: ITag | undefined, tagList?: ITag[]) {
    if (!tag || !tagList || !tagList.length || tagList.length === 0) {
      return false;
    }
    return tagList.filter(compareTag => compareTag.key === tag.key).length > 0;
  }
}