import * as React from 'react';
import styles from './SpFxList.module.scss';
import type { ISpFxListProps } from './ISpFxListProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { 
  FluentProvider,
  List, 
  ListItem, 
  Spinner,
  webLightTheme,
  DataGrid,
  DataGridHeader,
  DataGridBody,
  DataGridRow,
  DataGridCell
} from '@fluentui/react-components';
import { ArrowUp16Regular, ArrowDown16Regular, Folder16Regular, Document16Regular, ChevronRight16Regular } from '@fluentui/react-icons';

interface IListItem {
  Id: number;
  Title?: string;
  FileLeafRef?: string;
  FSObjType?: number;
  ContentType?: string;
  ServerRelativeUrl?: string;
  [key: string]: any;
}

type SortDirection = 'asc' | 'desc' | null;
type SortColumn = 'Title' | 'Modified' | 'Created' | 'Author' | 'Editor' | null;

interface ISpFxListState {
  items: IListItem[];
  loading: boolean;
  error: string | null;
  listTitle: string;
  baseTemplate: number;
  sortColumn: SortColumn;
  sortDirection: SortDirection;
  currentFolderPath: string;
  folderPathHistory: string[];
  columnWidths: { [key: string]: number };
  resizingColumn: string | null;
}

export default class SpFxList extends React.Component<ISpFxListProps, ISpFxListState> {
  constructor(props: ISpFxListProps) {
    super(props);
    console.log('SpFxList constructor called with props:', props);
    this.state = {
      items: [],
      loading: false,
      error: null,
      listTitle: '',
      baseTemplate: 100,
      sortColumn: null,
      sortDirection: null,
      currentFolderPath: '',
      folderPathHistory: [],
      columnWidths: {
        Title: 300,
        Modified: 150,
        Created: 150,
        Author: 200,
        Editor: 200
      },
      resizingColumn: null
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
      // Reset folder path when list changes
      this.setState({ 
        currentFolderPath: '', 
        folderPathHistory: [] 
      }, () => {
        this._loadListItems();
      });
    }
  }

  private _loadListItems(folderPath?: string): void {
    if (!this.props.selectedListId) {
      this.setState({ 
        items: [], 
        error: null, 
        listTitle: '', 
        baseTemplate: 100, 
        sortColumn: null, 
        sortDirection: null,
        currentFolderPath: '',
        folderPathHistory: []
      });
      return;
    }

    this.setState({ loading: true, error: null });

    // First get the list title and base template
    const listUrl = `${this.props.webUrl}/_api/web/lists(guid'${this.props.selectedListId}')?$select=Title,BaseTemplate,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    
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
        const baseTemplate = listData.BaseTemplate || (listData.d && listData.d.BaseTemplate) || 100;
        const rootFolderUrl = listData.RootFolder?.ServerRelativeUrl || (listData.d && listData.d.RootFolder?.ServerRelativeUrl) || '';
        
        this.setState({ listTitle: title, baseTemplate });

        const isDocumentLibrary = baseTemplate === 101;
        const targetFolderPath = folderPath !== undefined ? folderPath : this.state.currentFolderPath;
        
        if (isDocumentLibrary) {
          // For document libraries, get folders and files from the current folder
          // Update current folder path in state if it was passed as parameter
          if (folderPath !== undefined) {
            this.setState({ currentFolderPath: folderPath });
          }
          return this._loadDocumentLibraryItems(rootFolderUrl, targetFolderPath)
            .then((items: IListItem[]) => {
              console.log('Document library items loaded:', items.length, items);
              this.setState({ items, loading: false, error: null });
            });
        } else {
          // For regular lists, get items as before
          const itemsUrl = `${this.props.webUrl}/_api/web/lists(guid'${this.props.selectedListId}')/items?$top=100&$select=Id,Title,FileLeafRef,Modified,Created,Author/Title,Editor/Title&$expand=Author,Editor&$orderby=Id desc`;
          return this.props.spHttpClient.get(itemsUrl, SPHttpClient.configurations.v1)
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
            });
        }
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

  private _loadDocumentLibraryItems(rootFolderUrl: string, folderPath: string): Promise<IListItem[]> {
    const fullFolderPath = folderPath 
      ? `${rootFolderUrl}/${folderPath}`.replace(/\/+/g, '/')
      : rootFolderUrl;
    
    // Get folders and files from the current folder
    const folderUrl = `${this.props.webUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(fullFolderPath)}')`;
    const foldersUrl = `${folderUrl}/Folders?$select=Name,ServerRelativeUrl,TimeLastModified,ItemCount`;
    const filesUrl = `${folderUrl}/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Modified,Created,Author/Title,Editor/Title&$expand=Author,Editor`;
    
    // Fetch both folders and files in parallel
    return Promise.all([
      this.props.spHttpClient.get(foldersUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (!response.ok) {
            // If folder doesn't exist or is empty, return empty array
            return { d: { results: [] }, value: [] };
          }
          return response.json();
        })
        .then((data: any) => {
          const folders = data.d?.results || data.value || [];
          return folders.map((folder: any) => ({
            Id: 0, // Folders don't have item IDs in this context
            Title: folder.Name,
            FileLeafRef: folder.Name,
            FSObjType: 1, // 1 = Folder
            ServerRelativeUrl: folder.ServerRelativeUrl,
            Modified: folder.TimeLastModified,
            IsFolder: true
          }));
        })
        .catch(() => []), // Return empty array on error
      
      this.props.spHttpClient.get(filesUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (!response.ok) {
            return { d: { results: [] }, value: [] };
          }
          return response.json();
        })
        .then((data: any) => {
          const files = data.d?.results || data.value || [];
          return files.map((file: any, index: number) => ({
            Id: index + 10000, // Use index-based ID for files
            Title: file.Name,
            FileLeafRef: file.Name,
            FSObjType: 0, // 0 = File
            ServerRelativeUrl: file.ServerRelativeUrl,
            Modified: file.Modified || file.TimeLastModified,
            Created: file.Created,
            Author: file.Author,
            Editor: file.Editor,
            IsFolder: false
          }));
        })
        .catch(() => []) // Return empty array on error
    ]).then(([folders, files]) => {
      // Combine folders and files, folders first
      const allItems = [...folders, ...files];
      return allItems;
    });
  }

  private _handleSort = (column: SortColumn): void => {
    const { sortColumn, sortDirection } = this.state;
    
    let newDirection: SortDirection = 'asc';
    if (sortColumn === column && sortDirection === 'asc') {
      newDirection = 'desc';
    } else if (sortColumn === column && sortDirection === 'desc') {
      newDirection = null;
    }

    this.setState({ 
      sortColumn: newDirection ? column : null, 
      sortDirection: newDirection 
    });
  }

  private _getSortedItems = (): IListItem[] => {
    const { items, sortColumn, sortDirection } = this.state;
    
    // Ensure items is always an array
    if (!items || !Array.isArray(items)) {
      return [];
    }
    
    if (!sortColumn || !sortDirection) {
      return items;
    }

    const sorted = [...items].sort((a, b) => {
      let aValue: any;
      let bValue: any;

      switch (sortColumn) {
        case 'Title':
          aValue = (a.Title || a.FileLeafRef || '').toLowerCase();
          bValue = (b.Title || b.FileLeafRef || '').toLowerCase();
          break;
        case 'Modified':
          aValue = a.Modified ? new Date(a.Modified).getTime() : 0;
          bValue = b.Modified ? new Date(b.Modified).getTime() : 0;
          break;
        case 'Created':
          aValue = a.Created ? new Date(a.Created).getTime() : 0;
          bValue = b.Created ? new Date(b.Created).getTime() : 0;
          break;
        case 'Author':
          aValue = (a.Author?.Title || '').toLowerCase();
          bValue = (b.Author?.Title || '').toLowerCase();
          break;
        case 'Editor':
          aValue = (a.Editor?.Title || '').toLowerCase();
          bValue = (b.Editor?.Title || '').toLowerCase();
          break;
        default:
          return 0;
      }

      if (aValue < bValue) return sortDirection === 'asc' ? -1 : 1;
      if (aValue > bValue) return sortDirection === 'asc' ? 1 : -1;
      return 0;
    });

    return sorted;
  }

  private _renderSortIcon = (column: SortColumn): React.ReactNode => {
    const { sortColumn, sortDirection } = this.state;
    
    if (sortColumn !== column) {
      return null;
    }

    return sortDirection === 'asc' 
      ? <ArrowUp16Regular style={{ marginLeft: '4px', verticalAlign: 'middle' }} />
      : <ArrowDown16Regular style={{ marginLeft: '4px', verticalAlign: 'middle' }} />;
  }

  private _handleResizeStart = (columnKey: string, e: React.MouseEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    this.setState({ resizingColumn: columnKey });
    
    const startX = e.clientX;
    const startWidth = this.state.columnWidths[columnKey] || 200;
    let rafId: number | null = null;
    let pendingWidth: number = startWidth;
    
    const updateWidth = (newWidth: number): void => {
      pendingWidth = newWidth;
      if (rafId === null) {
        rafId = requestAnimationFrame(() => {
          this.setState({
            columnWidths: {
              ...this.state.columnWidths,
              [columnKey]: pendingWidth
            }
          });
          rafId = null;
        });
      }
    };
    
    const handleMouseMove = (moveEvent: MouseEvent) => {
      const diff = moveEvent.clientX - startX;
      const newWidth = Math.max(100, startWidth + diff); // Minimum width of 100px
      updateWidth(newWidth);
    };
    
    const handleMouseUp = () => {
      if (rafId !== null) {
        cancelAnimationFrame(rafId);
        // Apply final width
        this.setState({
          columnWidths: {
            ...this.state.columnWidths,
            [columnKey]: pendingWidth
          },
          resizingColumn: null
        });
      } else {
        this.setState({ resizingColumn: null });
      }
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };
    
    document.addEventListener('mousemove', handleMouseMove, { passive: true });
    document.addEventListener('mouseup', handleMouseUp);
  }

  private _getColumnDefinitions = () => {
    const { columnWidths } = this.state;
    return [
      { columnId: 'Title', width: columnWidths['Title'] || 300, minWidth: 100 },
      { columnId: 'Modified', width: columnWidths['Modified'] || 150, minWidth: 100 },
      { columnId: 'Created', width: columnWidths['Created'] || 150, minWidth: 100 },
      { columnId: 'Author', width: columnWidths['Author'] || 200, minWidth: 100 },
      { columnId: 'Editor', width: columnWidths['Editor'] || 200, minWidth: 100 }
    ];
  }

  private _renderDataGridHeader = (): React.ReactNode => {
    const { columnWidths, resizingColumn } = this.state;
    const columns = [
      { key: 'Title', label: 'Title' },
      { key: 'Modified', label: 'Modified' },
      { key: 'Created', label: 'Created' },
      { key: 'Author', label: 'Author' },
      { key: 'Editor', label: 'Editor' }
    ] as Array<{ key: SortColumn; label: string }>;

    return (
<<<<<<< HEAD
      // @ts-expect-error - Fluent UI v9 DataGridHeader typing issue with React 17
      <DataGridHeader>
        {columns.map((col, index) => {
          const columnKey = col.key || '';
          const width = col.key ? (columnWidths[col.key] || 200) : 200;
          return (
            // @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17
            <DataGridCell
              // @ts-expect-error - Fluent UI v9 DataGridCell key prop typing issue with React 17
              key={columnKey}
              // @ts-expect-error - Fluent UI v9 DataGridCell style prop typing issue with React 17
=======
      <thead>
        <tr style={{ 
          backgroundColor: '#f3f2f1', 
          borderBottom: '2px solid #e1dfdd',
          cursor: 'pointer'
        }}>
          {columns.map(col => (
            <th
              key={col.key}
              onClick={() => this._handleSort(col.key)}
>>>>>>> parent of 148480a (Switched to HTML based list)
              style={{
                width: `${width}px`,
                minWidth: '100px',
                position: 'relative',
                padding: '12px 16px',
                textAlign: 'left',
                fontWeight: '600',
                fontSize: '14px',
                userSelect: 'none',
                cursor: resizingColumn === columnKey ? 'col-resize' : 'pointer',
                backgroundColor: '#f3f2f1',
                borderRight: index < columns.length - 1 ? '1px solid #e1dfdd' : 'none'
              }}
              // @ts-expect-error - Fluent UI v9 DataGridCell onClick prop typing issue with React 17
              onClick={() => col.key && this._handleSort(col.key)}
            >
              {col.label}
              {this._renderSortIcon(col.key)}
              {index < columns.length - 1 && (
                <div
                  onMouseDown={(e) => this._handleResizeStart(columnKey, e)}
                  style={{
                    position: 'absolute',
                    right: 0,
                    top: 0,
                    bottom: 0,
                    width: '8px',
                    cursor: 'col-resize',
                    backgroundColor: resizingColumn === columnKey ? '#0078d4' : 'transparent',
                    zIndex: 1,
                    userSelect: 'none'
                  }}
                  onMouseEnter={(e) => {
                    if (resizingColumn !== columnKey) {
                      e.currentTarget.style.backgroundColor = '#c8c6c4';
                    }
                  }}
                  onMouseLeave={(e) => {
                    if (resizingColumn !== columnKey) {
                      e.currentTarget.style.backgroundColor = 'transparent';
                    }
                  }}
                />
              )}
            </DataGridCell>
          );
        })}
      </DataGridHeader>
    );
  }

  private _renderDataGridRow = (item: IListItem): React.ReactNode => {
    const { columnWidths } = this.state;
    const displayName = item.Title || item.FileLeafRef || `Item ${item.Id}`;
    const modified = item.Modified ? new Date(item.Modified).toLocaleDateString() : '';
    const created = item.Created ? new Date(item.Created).toLocaleDateString() : '';
    const author = item.Author?.Title || '';
    const editor = item.Editor?.Title || '';

    return (
      // @ts-expect-error - Fluent UI v9 DataGridRow typing issue with React 17
      <DataGridRow 
        key={item.Id}
        style={{
          borderBottom: '1px solid #e1dfdd',
          transition: 'background-color 0.1s'
        }}
        onMouseEnter={(e: React.MouseEvent<HTMLElement>) => {
          e.currentTarget.style.backgroundColor = '#faf9f8';
        }}
        onMouseLeave={(e: React.MouseEvent<HTMLElement>) => {
          e.currentTarget.style.backgroundColor = 'transparent';
        }}
      >
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Title'] || 300}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {displayName}
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Modified'] || 150}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {modified}
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Created'] || 150}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {created}
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Author'] || 200}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {author}
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Editor'] || 200}px`,
          minWidth: '100px',
          padding: '12px 16px',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {editor}
        </DataGridCell>
      </DataGridRow>
    );
  }

<<<<<<< HEAD
  private _renderDocumentLibraryDataGridRow = (item: IListItem, index: number): React.ReactNode => {
    const { columnWidths } = this.state;
    const displayName = item.Title || item.FileLeafRef || `Item ${item.Id}`;
    const modified = item.Modified ? new Date(item.Modified).toLocaleDateString() : '';
    const created = item.Created ? new Date(item.Created).toLocaleDateString() : '';
    const author = item.Author?.Title || '';
    const editor = item.Editor?.Title || '';
    const isFolder = this._isFolder(item);

    return (
      // @ts-expect-error - Fluent UI v9 DataGridRow typing issue with React 17
      <DataGridRow 
        key={`${item.Id}-${index}`}
        style={{
          borderBottom: '1px solid #e1dfdd',
          transition: 'background-color 0.1s'
        }}
        onMouseEnter={(e: React.MouseEvent<HTMLElement>) => {
          e.currentTarget.style.backgroundColor = '#faf9f8';
        }}
        onMouseLeave={(e: React.MouseEvent<HTMLElement>) => {
          e.currentTarget.style.backgroundColor = 'transparent';
        }}
      >
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Title'] || 300}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {isFolder ? (
              <>
                <Folder16Regular style={{ color: '#0078d4', flexShrink: 0 }} />
                <span
                  onClick={() => this._navigateToFolder(displayName)}
                  style={{
                    cursor: 'pointer',
                    fontWeight: '600',
                    fontSize: '14px',
                    color: '#0078d4',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                    flex: 1
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.textDecoration = 'underline';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.textDecoration = 'none';
                  }}
                >
                  {displayName}
                </span>
              </>
            ) : (
              <>
                <Document16Regular style={{ color: '#666', flexShrink: 0 }} />
                <a
                  href={this._getDocumentUrl(item)}
                  target="_blank"
                  rel="noopener noreferrer"
                  onClick={(e) => {
                    e.stopPropagation();
                  }}
                  style={{
                    fontWeight: '600',
                    fontSize: '14px',
                    color: '#0078d4',
                    textDecoration: 'none',
                    cursor: 'pointer',
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                    flex: 1
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.textDecoration = 'underline';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.textDecoration = 'none';
                  }}
                >
                  {displayName}
                </a>
              </>
            )}
          </div>
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Modified'] || 150}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {modified}
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Created'] || 150}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {created}
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Author'] || 200}px`,
          minWidth: '100px',
          padding: '12px 16px',
          borderRight: '1px solid #e1dfdd',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {author}
        </DataGridCell>
        {/* @ts-expect-error - Fluent UI v9 DataGridCell typing issue with React 17 */}
        <DataGridCell style={{ 
          width: `${columnWidths['Editor'] || 200}px`,
          minWidth: '100px',
          padding: '12px 16px',
          overflow: 'hidden',
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap'
        }}>
          {editor}
        </DataGridCell>
      </DataGridRow>
    );
  }

=======
>>>>>>> parent of 148480a (Switched to HTML based list)
  private _navigateToFolder = (folderName: string): void => {
    const { currentFolderPath, folderPathHistory } = this.state;
    const newPath = currentFolderPath 
      ? `${currentFolderPath}/${folderName}` 
      : folderName;
    
    const newHistory = [...folderPathHistory, currentFolderPath];
    
    this.setState({ 
      currentFolderPath: newPath,
      folderPathHistory: newHistory
    }, () => {
      this._loadListItems(newPath);
    });
  }

  private _navigateToPath = (pathIndex: number): void => {
    const { currentFolderPath } = this.state;
    const pathParts = currentFolderPath ? currentFolderPath.split('/').filter(p => p) : [];
    
    if (pathIndex < 0) {
      // Navigate to root
      this.setState({ 
        currentFolderPath: '',
        folderPathHistory: []
      }, () => {
        this._loadListItems('');
      });
    } else {
      // Navigate to a specific folder in the path
      const targetParts = pathParts.slice(0, pathIndex + 1);
      const targetPath = targetParts.join('/');
      const newHistory: string[] = [];
      
      // Build history from path parts
      let accumulatedPath = '';
      targetParts.forEach((part, idx) => {
        if (idx < targetParts.length - 1) {
          accumulatedPath = accumulatedPath ? `${accumulatedPath}/${part}` : part;
          newHistory.push(accumulatedPath);
        }
      });
      
      this.setState({ 
        currentFolderPath: targetPath,
        folderPathHistory: newHistory
      }, () => {
        this._loadListItems(targetPath);
      });
    }
  }

  private _isFolder = (item: IListItem): boolean => {
    return item.FSObjType === 1 || item.IsFolder === true || item.ContentType === 'Folder';
  }

  private _renderBreadcrumb = (): React.ReactNode => {
    const { currentFolderPath, listTitle } = this.state;
    
    const pathParts = currentFolderPath ? currentFolderPath.split('/').filter(p => p) : [];
    const breadcrumbItems: Array<{ name: string; path: string; index: number }> = [
      { name: listTitle || 'Root', path: '', index: -1 }
    ];

    // Build breadcrumb from path parts
    let accumulatedPath = '';
    pathParts.forEach((part, index) => {
      accumulatedPath = accumulatedPath ? `${accumulatedPath}/${part}` : part;
      breadcrumbItems.push({ name: part, path: accumulatedPath, index });
    });

    // Only show breadcrumb if we're in a subfolder
    if (pathParts.length === 0) {
      return null;
    }

    return (
      <div style={{ 
        marginBottom: '16px', 
        padding: '12px 16px', 
        backgroundColor: '#f3f2f1', 
        borderRadius: '4px',
        display: 'flex',
        alignItems: 'center',
        flexWrap: 'wrap',
        gap: '4px'
      }}>
        {breadcrumbItems.map((item, idx) => (
          <React.Fragment key={idx}>
            {idx > 0 && (
              <ChevronRight16Regular style={{ margin: '0 4px', color: '#666' }} />
            )}
            <span
              onClick={() => {
                if (item.index === -1) {
                  // Navigate to root
                  this.setState({ 
                    currentFolderPath: '',
                    folderPathHistory: []
                  }, () => {
                    this._loadListItems('');
                  });
                } else {
                  this._navigateToPath(item.index);
                }
              }}
              style={{
                cursor: 'pointer',
                color: '#0078d4',
                textDecoration: 'none',
                padding: '4px 8px',
                borderRadius: '2px',
                transition: 'background-color 0.1s'
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.backgroundColor = '#e1dfdd';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.backgroundColor = 'transparent';
              }}
            >
              {item.name}
            </span>
          </React.Fragment>
        ))}
      </div>
    );
  }

  private _getDocumentUrl = (item: IListItem): string => {
    if (item.ServerRelativeUrl) {
      // Construct the full URL to open the document
      return `${this.props.webUrl}${item.ServerRelativeUrl}`;
    }
    // Fallback: construct URL from file name (shouldn't happen if ServerRelativeUrl is properly set)
    const fileName = item.FileLeafRef || item.Title || '';
    return `${this.props.webUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(fileName)}`;
  }

  private _renderItemContent = (item: IListItem): React.ReactNode => {
    const displayName = item.Title || item.FileLeafRef || `Item ${item.Id}`;
    const modified = item.Modified ? new Date(item.Modified).toLocaleDateString() : '';
    const author = item.Author?.Title || '';
    const isFolder = this._isFolder(item);
    
    return (
      <div 
        style={{ 
          padding: '8px 0',
          cursor: isFolder ? 'pointer' : 'default'
        }}
        onClick={() => {
          if (isFolder) {
            this._navigateToFolder(displayName);
          }
        }}
        onMouseEnter={(e) => {
          if (isFolder) {
            e.currentTarget.style.backgroundColor = '#faf9f8';
          }
        }}
        onMouseLeave={(e) => {
          if (isFolder) {
            e.currentTarget.style.backgroundColor = 'transparent';
          }
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
          {isFolder ? (
            <Folder16Regular style={{ color: '#0078d4', flexShrink: 0 }} />
          ) : (
            <Document16Regular style={{ color: '#666', flexShrink: 0 }} />
          )}
          {isFolder ? (
            <div style={{ fontWeight: '600', fontSize: '14px', flex: 1 }}>
              {displayName}
            </div>
          ) : (
            <a
              href={this._getDocumentUrl(item)}
              target="_blank"
              rel="noopener noreferrer"
              onClick={(e) => {
                // Prevent event bubbling to parent div
                e.stopPropagation();
              }}
              style={{
                fontWeight: '600',
                fontSize: '14px',
                flex: 1,
                color: '#0078d4',
                textDecoration: 'none',
                cursor: 'pointer'
              }}
              onMouseEnter={(e) => {
                e.currentTarget.style.textDecoration = 'underline';
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.textDecoration = 'none';
              }}
            >
              {displayName}
            </a>
          )}
        </div>
        {(modified || author) && (
          <div style={{ fontSize: '12px', color: '#666', marginLeft: '24px' }}>
            {modified && <span style={{ marginRight: '12px' }}>Modified: {modified}</span>}
            {author && <span>By: {author}</span>}
          </div>
        )}
      </div>
    );
  }

  public render(): React.ReactElement<ISpFxListProps> {
    const { items, loading, error, listTitle } = this.state;
    console.log('SpFxList render - items:', items?.length || 0, 'loading:', loading, 'error:', error);
    console.log('State:', this.state);
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
            {/* @ts-expect-error - Fluent UI v9 Spinner typing issue with React 17 */}
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

      const { baseTemplate, currentFolderPath, items } = this.state;
      const isDocumentLibrary = baseTemplate === 101;
      const sortedItems = this._getSortedItems();

      // Ensure items is an array
      const itemsArray = Array.isArray(items) ? items : [];
      const sortedItemsArray = Array.isArray(sortedItems) ? sortedItems : [];

      console.log('Rendering content - itemsArray.length:', itemsArray.length, 'sortedItemsArray.length:', sortedItemsArray.length, 'isDocumentLibrary:', isDocumentLibrary);

      return (
        <div className={styles.spFxList}>
          {listTitle && (
            <h2 style={{ marginBottom: '16px' }}>{listTitle}</h2>
          )}
          {isDocumentLibrary && this._renderBreadcrumb()}
          {itemsArray.length === 0 ? (
            <div style={{ padding: '16px', backgroundColor: '#f3f2f1', borderRadius: '4px' }}>
              {isDocumentLibrary && currentFolderPath 
                ? 'No items found in this folder.' 
                : 'No items found in this list.'}
            </div>
          ) : isDocumentLibrary ? (
<<<<<<< HEAD
            // Document Library - render with Fluent UI v9 DataGrid with sortable and resizable columns
            <div style={{ border: '1px solid #e1dfdd', borderRadius: '4px', overflow: 'hidden', minHeight: '200px' }}>
              <div style={{ padding: '8px', backgroundColor: '#faf9f8', borderBottom: '1px solid #e1dfdd' }}>
                <strong>Document Library Items ({sortedItemsArray.length})</strong>
              </div>
              {/* @ts-expect-error - Fluent UI v9 DataGrid typing issue with React 17 */}
              <DataGrid 
                // @ts-expect-error - Fluent UI v9 DataGrid items prop typing issue with React 17
                items={sortedItemsArray}
                // @ts-expect-error - Fluent UI v9 DataGrid columns prop typing issue with React 17
                columns={this._getColumnDefinitions()}
              >
                {this._renderDataGridHeader()}
                {/* @ts-ignore - Fluent UI v9 DataGridBody typing issue with React 17 - expects RowRenderFunction but we're using array of rows */}
                <DataGridBody {...({} as any)}>
                  {sortedItemsArray.length > 0 ? sortedItemsArray.map((item, index) => this._renderDocumentLibraryDataGridRow(item, index)) : <div style={{ padding: '16px' }}>No items to display</div>}
                </DataGridBody>
              </DataGrid>
=======
            // Document Library - render with folder navigation
            <div style={{ border: '1px solid #e1dfdd', borderRadius: '4px', overflow: 'hidden' }}>
              {/* @ts-expect-error - Fluent UI v9 List typing issue with React 17 */}
              <List>
                {items.map((item, index) => (
                  // @ts-expect-error - Fluent UI v9 ListItem typing issue with React 17
                  <ListItem key={`${item.Id}-${index}`}>
                    {this._renderItemContent(item)}
                  </ListItem>
                ))}
              </List>
>>>>>>> parent of 148480a (Switched to HTML based list)
            </div>
          ) : (
            // Regular List - render with Fluent UI v9 DataGrid with sortable and resizable columns
            <div style={{ border: '1px solid #e1dfdd', borderRadius: '4px', overflow: 'hidden', minHeight: '200px' }}>
              <div style={{ padding: '8px', backgroundColor: '#faf9f8', borderBottom: '1px solid #e1dfdd' }}>
                <strong>List Items ({sortedItemsArray.length})</strong>
              </div>
              {/* @ts-expect-error - Fluent UI v9 DataGrid typing issue with React 17 */}
              <DataGrid 
                // @ts-expect-error - Fluent UI v9 DataGrid items prop typing issue with React 17
                items={sortedItemsArray}
                // @ts-expect-error - Fluent UI v9 DataGrid columns prop typing issue with React 17
                columns={this._getColumnDefinitions()}
              >
                {this._renderDataGridHeader()}
                {/* @ts-ignore - Fluent UI v9 DataGridBody typing issue with React 17 - expects RowRenderFunction but we're using array of rows */}
                <DataGridBody {...({} as any)}>
                  {sortedItemsArray.length > 0 ? sortedItemsArray.map((item) => this._renderDataGridRow(item)) : <div style={{ padding: '16px' }}>No items to display</div>}
                </DataGridBody>
              </DataGrid>
            </div>
          )}
        </div>
      );
    })();

    return (
      // @ts-expect-error - Fluent UI v9 FluentProvider typing issue with React 17
      <FluentProvider theme={webLightTheme}>
        {content}
      </FluentProvider>
    );
  }
}
