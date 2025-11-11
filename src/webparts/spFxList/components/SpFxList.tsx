import * as React from 'react';
import styles from './SpFxList.module.scss';
import type { ISpFxListProps } from './ISpFxListProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { 
  FluentProvider,
  List, 
  ListItem, 
  Spinner,
  webLightTheme
} from '@fluentui/react-components';

interface IListItem {
  Id: number;
  Title?: string;
  [key: string]: any;
}

interface ISpFxListState {
  items: IListItem[];
  loading: boolean;
  error: string | null;
  listTitle: string;
}

export default class SpFxList extends React.Component<ISpFxListProps, ISpFxListState> {
  constructor(props: ISpFxListProps) {
    super(props);
    console.log('SpFxList constructor called with props:', props);
    this.state = {
      items: [],
      loading: false,
      error: null,
      listTitle: ''
    };
  }

  public componentDidMount(): void {
    console.log('SpFxList componentDidMount called');
    if (this.props.selectedListId) {
      this._loadListItems();
    } else {
      console.log('No selectedListId, skipping load');
    }
  }

  public componentDidUpdate(prevProps: ISpFxListProps): void {
    if (prevProps.selectedListId !== this.props.selectedListId) {
      this._loadListItems();
    }
  }

  private _loadListItems(): void {
    if (!this.props.selectedListId) {
      this.setState({ items: [], error: null, listTitle: '' });
      return;
    }

    this.setState({ loading: true, error: null });

    // First get the list title
    const listUrl = `${this.props.webUrl}/_api/web/lists(guid'${this.props.selectedListId}')?$select=Title`;
    
    this.props.spHttpClient.get(listUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Failed to load list: ${response.status}`);
        }
        return response.json();
      })
      .then((listData: any) => {
        // Handle both response formats
        const title = listData.Title || (listData.d && listData.d.Title) || '';
        this.setState({ listTitle: title });

        // Then get the list items - fetch common fields
        const itemsUrl = `${this.props.webUrl}/_api/web/lists(guid'${this.props.selectedListId}')/items?$top=100&$select=Id,Title,FileLeafRef,Modified,Created,Author/Title,Editor/Title&$expand=Author,Editor&$orderby=Id desc`;
        
        return this.props.spHttpClient.get(itemsUrl, SPHttpClient.configurations.v1);
      })
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Failed to load items: ${response.status}`);
        }
        return response.json();
      })
      .then((data: any) => {
        // Handle both odata=verbose (d.results) and odata=nometadata (value) formats
        let items: IListItem[] = [];
        
        if (data && data.d && data.d.results) {
          items = data.d.results;
        } else if (data && data.value) {
          items = data.value;
        }

        console.log('List items loaded:', items.length, items);
        this.setState({ items, loading: false, error: null });
      })
      .catch((error: Error) => {
        console.error('Error loading list items:', error);
        this.setState({ 
          items: [], 
          loading: false, 
          error: error.message || 'Failed to load list items' 
        });
      });
  }

  private _renderItemContent = (item: IListItem): React.ReactNode => {
    const displayName = item.Title || item.FileLeafRef || `Item ${item.Id}`;
    const modified = item.Modified ? new Date(item.Modified).toLocaleDateString() : '';
    const author = item.Author?.Title || '';
    
    return (
      <div style={{ padding: '8px 0' }}>
        <div style={{ fontWeight: '600', marginBottom: '4px', fontSize: '14px' }}>
          {displayName}
        </div>
        {(modified || author) && (
          <div style={{ fontSize: '12px', color: '#666' }}>
            {modified && <span style={{ marginRight: '12px' }}>Modified: {modified}</span>}
            {author && <span>By: {author}</span>}
          </div>
        )}
      </div>
    );
  }

  public render(): React.ReactElement<ISpFxListProps> {
    const { items, loading, error, listTitle } = this.state;
    console.log('SpFxList render - items:', items.length, 'loading:', loading, 'error:', error);
    console.log('Props:', this.props);

    const content = (() => {
      if (!this.props.selectedListId) {
        console.log('Rendering: No list selected message');
        return (
          <div className={styles.spFxList} style={{ padding: '20px' }}>
            <div style={{ backgroundColor: '#f3f2f1', padding: '16px', borderRadius: '4px' }}>
              <strong>Please select a list or document library from the web part properties.</strong>
            </div>
          </div>
        );
      }

      if (loading) {
        return (
          <div className={styles.spFxList}>
            {/* @ts-expect-error - Fluent UI v9 Spinner typing issue */}
            <Spinner label="Loading items..." size="medium" />
          </div>
        );
      }

      if (error) {
        return (
          <div className={styles.spFxList}>
            <div style={{ color: 'red', padding: '16px', backgroundColor: '#fef0f0', borderRadius: '4px' }}>
              Error: {error}
            </div>
          </div>
        );
      }

      return (
        <div className={styles.spFxList}>
          {listTitle && (
            <h2 style={{ marginBottom: '16px' }}>{listTitle}</h2>
          )}
          {items.length === 0 ? (
            <div style={{ padding: '16px', backgroundColor: '#f3f2f1', borderRadius: '4px' }}>
              No items found in this list.
            </div>
          ) : (
            <div style={{ border: '1px solid #e1dfdd', borderRadius: '4px', overflow: 'hidden' }}>
              {/* @ts-expect-error - Fluent UI v9 List typing issue */}
              <List>
                {items.map((item) => (
                  // @ts-expect-error - Fluent UI v9 ListItem typing issue
                  <ListItem key={item.Id.toString()}>
                    {this._renderItemContent(item)}
                  </ListItem>
                ))}
              </List>
            </div>
          )}
        </div>
      );
    })();

    return (
      // @ts-expect-error - Fluent UI v9 FluentProvider typing issue
      <FluentProvider theme={webLightTheme}>
        {content}
      </FluentProvider>
    );
  }
}
