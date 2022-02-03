//Importing files and objects creation
import * as React from 'react';
import styles from './DocsNodeAdmin.module.scss';
import { IDocsNodeAdminProps } from './IDocsNodeAdminProps';
import { Selection, SelectionMode, IColumn, ColumnActionsMode } from 'office-ui-fabric-react/lib/DetailsList';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import DatabaseConfiguration from './DatabaseConfiguration';
import { IDetailsListDocumentsState, IDocument } from './IDetailsListDocumentsStates';
import { CommandBarButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, DropdownMenuItemType } from 'office-ui-fabric-react/lib/Dropdown';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import constant from './Constant';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Nav } from 'office-ui-fabric-react/lib/Nav';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';

//Created objects of other files
const DC: DatabaseConfiguration = new DatabaseConfiguration();
//Declared and initialized global variables for See more functionalities 
const _INTERVAL_DELAY = 100;
export default class DocsNodeAdmin extends React.Component<IDocsNodeAdminProps, IDetailsListDocumentsState> {

  //Declaration and initialization of variables
  private _selection: Selection;
  public _SlidesItemArray = [];
  public _ImagesItemArray = [];
  public _TextSnippetItemArray = [];
  public _CategoriesItemArray = [];
  public _CategoriesItemsData = [];
  public _CategoriesDrpDwnItemArray = [];
  public _ParentCategoriesItemArray = [];
  public _SearchResult = [];
  private _lastIndexWithData: number;
  public _ITEMS_COUNT = 0;
  public _items = [];
  public _defaultRadioBtnCategoryType = '';
  //Props initialization
  constructor(props) {
    super(props);

    //Creating columns for Slides,Image,Text snippet List
    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'File Type',
        className: styles.fileIconCell,
        iconClassName: styles.fileIconHeaderIcon,
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item: IDocument) => {
          return (item.FileType != '' ? <img src={item.FileType} className={styles.fileIconImg} /> : <Icon iconName='TextDocument' />);
        }
      },
      {
        key: 'column2',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <Link href={item.LinkURL} target='_blank'>{item.Title}</Link>;
        }
      },
      {
        key: 'column3',
        name: constant.CategoryName,
        fieldName: constant.CategoryName,
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.Category}</span>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Edit',
        fieldName: 'edit',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editItem(item)} />
        }
      },
      {
        key: 'column5',
        name: 'Delete',
        fieldName: 'delete',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._showDialog(item)} />;
        }
      }
    ];

    //Creating columns for Category List
    const CategoryColumns: IColumn[] = [      
      {
        key: 'column1',
        name: constant.CategoryName,
        fieldName: constant.CategoryName,
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.Category}</span>;
        }
      },
      {
        key: 'column2',
        name: constant.CategoryParentId,
        fieldName: constant.CategoryParentId,
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.ParentCategory}</span>;
        },
        isPadded: true
      },
      {
        key: 'column3',
        name: constant.CategoryType,
        fieldName: constant.CategoryType,
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.CategoryType}</span>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Edit',
        fieldName: 'edit',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editItem(item)} />
        }
      },
      {
        key: 'column5',
        name: 'Delete',
        fieldName: 'delete',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._showDialog(item)} />;
        }
      }
    ];


    //Selection of N numbers of rows
    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });

    //Initializing state
    this.state = {
      items: [],
      columns: columns,
      categoryColumns: CategoryColumns,
      selectionDetails: '',
      isModalSelection: true,
      isCompactMode: false,
      CategaryDropDown: [],
      isDataLoaded: false,
      tabValue: '',
      defaultKey: '',
      showPanel: false,
      defaultNavSelectdKey: '',
      itemDiscription: '',
      itemTitle: '',
      itemName: '',
      addOrEditBtn: false,
      listName: '',
      itemId: null,
      renderKey: '',
      hideDialog: true,
      deleteItem: [],
      showList: true,
      categoryListForm: false,
      newCategoryTitle: '',
      radioBtnCategory: ''
    };

    //Binding methods to get current context inside method
    this._onChangeText = this._onChangeText.bind(this);
    this._onLoadData = this._onLoadData.bind(this);
    this._refreshCategories = this._refreshCategories.bind(this);
    this._getDocsNodeImageItems = this._getDocsNodeImageItems.bind(this);
    this._getDocsNodeSlidesItems = this._getDocsNodeSlidesItems.bind(this);
    this._getDocsNodeTextSnippetItems = this._getDocsNodeTextSnippetItems.bind(this);
    this._onHandleChange = this._onHandleChange.bind(this);
    this._showPanel = this._showPanel.bind(this);
    this._addNewData = this._addNewData.bind(this);
    this._onLinkClickAssets = this._onLinkClickAssets.bind(this);
    this._editItem = this._editItem.bind(this);
    this._deleteItem = this._deleteItem.bind(this);
    this._EmptyList = this._EmptyList.bind(this);
    this._SelectCategory = this._SelectCategory.bind(this);
    this._showDialog = this._showDialog.bind(this);
    this._onChangeCategoryType = this._onChangeCategoryType.bind(this);
    this._getDocsNodeCategoryItems = this._getDocsNodeCategoryItems.bind(this);
  }

  //This method is invoked immediately after a component is mounted
  //This will load data from a remote endpoint
  componentDidMount() {
    //DC.getUpdate();
    //Check the All Default List and Library are exist 
    //If not, then this would create  new one
    DC._checkListExistsOrNot().then(async (data) => {

      //Check the columns of All Default List created are exist 
      //If not, then this would create columns for that List or Library
      await DC._checkForColumnExistence();
      var arry = { name: constant.SlidesDisplayName };

      //By default select Slide section in Manage Assets
      this._onLinkClickAssets('', arry);
    });
  }

  //Await Refresh Category List
  public async _refreshCategories(){
    var mainArray = await DC._getDocsNodeCategoriesName();
    this._CategoriesDrpDwnItemArray = mainArray.DocsNodeCategoriesArrayItems;    
    this._CategoriesItemArray = [{ key: 0, text: constant.AllSlideName }].concat(this._CategoriesDrpDwnItemArray);
    this._CategoriesDrpDwnItemArray = [{ key: 'Header', text: 'Select Category', itemType: DropdownMenuItemType.Header }].concat(this._CategoriesDrpDwnItemArray);
  }

  //This function is to get all items from DocsNodeSlide Library
  public async _getDocsNodeSlidesItems() {

    //Await to get all items from DocsNodeSlide Library
    this._SlidesItemArray = await DC._getDocsNodeSlidesName();
    this._SearchResult = this._SlidesItemArray;

    //Empty List
    await this._EmptyList();

    //Await to get all Categories from DocsNodeCategory List
    await this._refreshCategories();

    //This function is for lazy loading of items in List component
    this._onLoadData(this._SlidesItemArray.length,this._SlidesItemArray);

    //Binding the data
    this.setState({
      CategaryDropDown: this._CategoriesItemArray,
      defaultKey: '',
      renderKey: '',
      tabValue: constant.SlidesDisplayName,
      defaultNavSelectdKey: 'key1',
      listName: constant.DocsNodeSlidesName,
      showPanel: false,
      selectionDetails: 'No items selected',
      showList: true,
      categoryListForm: false
    });
  }

  //This function to get all items from Picture Library
  public async _getDocsNodeImageItems() {
    
    //Empty List
    await this._EmptyList();
    
    //Await to get all Categories from DocsNodeCategory List
    await this._refreshCategories();

    //Calling the api to get all items from DocsNodePicture Library
    DC._getDocsNodePictureName().then((responseData) => {
      
      this._ImagesItemArray = responseData;
      this._SearchResult = responseData;

      //This function is for lazy loading of items in List component
      this._onLoadData(this._ImagesItemArray.length,this._ImagesItemArray);

      //Binding the data
      this.setState({
        //items: this._ImagesItemArray,
        CategaryDropDown: this._CategoriesItemArray,
        defaultKey: '',
        renderKey: '',
        tabValue: constant.ImagesDisplayName,
        defaultNavSelectdKey: 'key2',
        listName: constant.DocsNodePictureName,
        showPanel: false,
        selectionDetails: 'No items selected',
        showList: true,
        categoryListForm: false
      });
    });
  }

  //This function is to get all items from Text Snippet List 
  public async _getDocsNodeTextSnippetItems() {

    //Empty List
    await this._EmptyList();

    //Await to get all Categories from DocsNodeCategory List
    await this._refreshCategories();

    //Calling the api to get all items from DocsNodeTextSnippet List
    DC._getDocsNodeTextSnippetName().then((responseData) => {
      this._TextSnippetItemArray = responseData;
      this._SearchResult = responseData;
      
      //This function is for lazy loading of items in List component
      this._onLoadData(this._TextSnippetItemArray.length, this._TextSnippetItemArray);

      //Binding the data
      this.setState({
        //items: this._TextSnippetItemArray,
        CategaryDropDown: this._CategoriesItemArray,
        defaultKey: '',
        renderKey: '',
        tabValue: constant.TextSnippetDisplayName,
        defaultNavSelectdKey: 'key3',
        listName: constant.DocsNodeTextName,
        showPanel: false,
        selectionDetails: 'No items selected',
        showList: true,
        categoryListForm: false
      });
    });
  }

  public async _getDocsNodeCategoryItems() {

    //Empty List
    await this._EmptyList(); 

    //Calling the api to get all items from DocsNodeCategory List
    DC._getDocsNodeCategoriesName().then((responseData)=>{
      
      this._CategoriesItemsData = responseData.DocsNodeCategoriesItemsData;
      this._SearchResult = responseData.DocsNodeCategoriesItemsData;
      this._ParentCategoriesItemArray = [{ key: 'Header', text: 'Select Parent Category', itemType: DropdownMenuItemType.Header }].concat(responseData.DocsNodeParentCategoriesArrayItems);

      //This function is for lazy loading of items in List component
      this._onLoadData(this._CategoriesItemsData.length,this._CategoriesItemsData);

      //Binding the data
      this.setState({
        //items: this._CategoriesItemsData,
        showList: false,
        defaultNavSelectdKey: 'key4',
        categoryListForm: true,
        showPanel: false,
        tabValue: constant.CategoryDisplayName,
        renderKey: '',
        defaultKey: '',
        listName: constant.DocsNodeCategoriesName
      });
    });    
  }

  public _EmptyList() {
    this.setState({
      items: [] ,
      isDataLoaded: false     
    });
  }

  //Following method is for filter the category list from the displaying items in List Component
  public _onHandleChange(event) {
    var result = [];
    var flag = false;
    const { tabValue } = this.state;
    if (tabValue == constant.SlidesDisplayName) {
      this._SlidesItemArray.map((items) => {
        if (items.Category == event.text) {
          result.push(items);
          flag = true;
        } else if (event.text == constant.AllSlideName) {
          result = this._SlidesItemArray;
          flag = true;
        }
      });
      if (flag == false) {
        result = [];
      }
    } else if (tabValue == constant.ImagesDisplayName) {
      this._ImagesItemArray.map((items) => {
        if (items.Category == event.text) {
          result.push(items);
          flag = true;
        } else if (event.text == constant.AllSlideName) {
          result = this._ImagesItemArray;
          flag = true;
        }
      });
      if (flag == false) {
        result = [];
      }
    } else {
      this._TextSnippetItemArray.map((items) => {
        if (items.Category == event.text) {
          result.push(items);
          flag = true;
        } else if (event.text == constant.AllSlideName) {
          result = this._TextSnippetItemArray;
          flag = true;
        }
      });
      if (flag == false) {
        result = [];
      }
    }
    this._SearchResult = result;
    this.setState({
      items: result,
      defaultKey: event.key
    });
  }

  //Get panel open
  public _showPanel() {
    this._defaultRadioBtnCategoryType = '';
    this.setState({
      showPanel: true,
      itemTitle: '',
      itemDiscription: '',
      addOrEditBtn: false,
      itemName: '',
      newCategoryTitle: '',
      renderKey: '',
      radioBtnCategory: ''
    });
  }

  public _SelectCategory(event) {
    this.setState({
      renderKey: event.key
    });
  }

  //This method is for rendering result
  public render() {

    //Creating constants for state varaibles
    const {
      columns,
      categoryColumns,
      items,
      selectionDetails,
      itemTitle,
      itemDiscription,
      isModalSelection,
      isDataLoaded,
      CategaryDropDown,
      defaultKey,
      showPanel,
      defaultNavSelectdKey,
      itemName,
      addOrEditBtn,
      renderKey,
      showList,
      categoryListForm,
      deleteItem,
      hideDialog,
      newCategoryTitle,
      radioBtnCategory
    } = this.state;

    //Will return the UI of webpart
    return (<div>
      <div className={styles.topNav}>
        <div className={styles.container}>
          <div className={styles.ProductIcon}>
            <img src={String(require('../images/logo.png'))} />
          </div>
          <div className={styles.productName}>
            <span>DocsNode Templates</span>
          </div>
        </div>
      </div>
      <div className={styles.container}>
        <div className={styles.leftNav}>
          <div className={styles.manageAssets}>
            <Nav
              selectedKey={defaultNavSelectdKey}
              onLinkClick={this._onLinkClickAssets}
              groups={[
                {
                  name: 'Manage Assets',
                  links: [
                    { name: constant.SlidesDisplayName, key: 'key1', url: '', iconProps: { iconName: 'Boards' } },
                    { name: constant.ImagesDisplayName, key: 'key2', url: '', iconProps: { iconName: 'FileImage' } },
                    { name: constant.TextSnippetDisplayName, key: 'key3', url: '', iconProps: { iconName: 'ContextMenu' } }
                  ]
                },
                {
                  name: 'Manage Assets Categories',
                  links: [
                    { name: constant.CategoryDisplayName, url: '', key: 'key4',iconProps:{iconName: 'List'} },
                  ]
                }
              ]}
            />
          </div>
        </div>
        <div className={styles.rightContent}>
          {showList && <div><div className={styles.ribbonTop}>
            <div className={styles.addNewBtn}>
              <CommandBarButton
                iconProps={{ iconName: 'Add' }}
                text="Add New"
                onClick={this._showPanel}
              />
            </div>
            <div className={styles.searchBox}>
              <SearchBox
                placeholder="Search"
                onEscape={ev => {
                  console.log('Custom onEscape Called');
                }}
                onClear={ev => {
                  console.log('Custom onClear Called');
                }}
                onChange={newValue => this._onChangeText(newValue)}
              />
            </div>
            <div className={styles.shortBy}>
              <Dropdown
                placeHolder="Select options"
                label='Categories : '
                selectedKey={defaultKey}
                options={CategaryDropDown}
                onChanged={this._onHandleChange}
              />
            </div>
          </div>
            <div>            
              <ShimmeredDetailsList
                setKey="items"
                items={items!}
                columns={columns}
                selectionMode={isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                enableShimmer={!isDataLoaded}
                listProps={{ renderedWindowsAhead: 0, renderedWindowsBehind: 0 }}
                isHeaderVisible={true}
                selection={this._selection}
                selectionPreservedOnEmptyClick={true}
                enterModalSelectionOnTouch={true}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              />
            </div>
          </div>}
          {categoryListForm && <div><div className={styles.ribbonTop}>
            <div className={styles.addNewBtn}>
              <CommandBarButton
                iconProps={{ iconName: 'Add' }}
                text="Add New"
                onClick={this._showPanel}
              />
            </div>
          </div>
            <div>
              <ShimmeredDetailsList
                setKey="items"
                items={items!}
                columns={categoryColumns}
                selectionMode={isModalSelection ? SelectionMode.multiple : SelectionMode.none}
                enableShimmer={!isDataLoaded}
                listProps={{ renderedWindowsAhead: 0, renderedWindowsBehind: 0 }}
                isHeaderVisible={true}
                selection={this._selection}
              />
            </div>
          </div>}
          <Dialog
            hidden={hideDialog}
            onDismiss={this._closeDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: deleteItem.Title,
              subText: 'Sure you want to delete this item?'
            }}
          >
            <DialogFooter>
              <PrimaryButton onClick={() => this._deleteItem(deleteItem)} text="Delete" />
              <DefaultButton onClick={this._closeDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
      </div>
      <div id="mySidenavAdd" className={styles.sidenav}>
        <div className={styles.headerSidenav}>
          <Panel
            isOpen={showPanel}
            onDismiss={this._hidePanel}
            type={PanelType.medium}
            headerText={addOrEditBtn == true ? itemName : 'New Item'}
            onRenderFooterContent={this._onRenderFooterContent}
            isFooterAtBottom={true}
          >
            <div className={styles.bodySidenav}>
              {showList && <div className={styles.NewIteamForm}>
                {this.state.tabValue != constant.TextSnippetDisplayName ? addOrEditBtn == true ? null : <div className={styles.chooseFile}>
                  <Label required>Choose file :</Label>
                  <TextField
                    className={styles.FormInput}
                    accept=".ppt,.pptx,image/*"
                    type="file" id='inputTypeFiles'
                    onGetErrorMessage={this._getErrorMessage}
                    validateOnLoad={false} />
                </div> : null}
                <div className={styles.FormInput}>
                  <Label required>Title :</Label>
                  <TextField id='titleID' placeholder="Enter value here" underlined
                    value={itemTitle}
                    onChanged={this._handleChangeTitle}
                    onGetErrorMessage={this._getErrorMessage}
                    validateOnLoad={false}
                  />
                </div>
                <div className={styles.FormInput}>
                  <Label>Description :</Label>
                  <TextField
                    id='discription'
                    multiline
                    autoAdjustHeight
                    placeholder="Enter value here"
                    underlined value={itemDiscription}
                    onChanged={this._handleChangeDiscript} />
                </div>
                <div className={styles.FormInput}>
                  <Label>Category :</Label>
                  <Dropdown
                    placeHolder="Select options"
                    selectedKey={renderKey}
                    options={this._CategoriesDrpDwnItemArray}
                    onChanged={this._SelectCategory}
                  />
                </div>
                {/* <div className={styles.FormInput}>
                  <label>Description <span>*</span></label>
                  <button className="modal-trigger" data-modal="modal-name"><i className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i></button>
                </div>                 */}
              </div>}
              {categoryListForm &&
                <div className={styles.NewIteamForm}>
                  <div className={styles.FormInput}>
                    <Label required>Category :</Label>
                    <TextField id='newCategoryName' placeholder="Enter new category here" underlined
                      value={newCategoryTitle}
                      onChanged={this._handleChangeCategoryTitle}
                      onGetErrorMessage={this._getErrorMessage}
                      validateOnLoad={false}
                    />
                  </div>
                  <div className={styles.FormInput}>
                    <Label>Parent Category(Optional) :</Label>
                    <Dropdown
                      placeHolder="Select options"                      
                      defaultSelectedKey={renderKey}
                      options={this._ParentCategoriesItemArray}
                      onChanged={this._SelectCategory}
                    />
                  </div> 
                  <div className={styles.FormInput}>
                    <ChoiceGroup                      
                      selectedKey ={radioBtnCategory}
                      options={[
                        {
                          key: constant.SlidesDisplayName,
                          text: constant.SlidesDisplayName
                        },
                        {
                          key: constant.ImagesDisplayName,
                          text: constant.ImagesDisplayName
                        },
                        {
                          key: constant.TextSnippetDisplayName,
                          text: constant.TextSnippetDisplayName
                        }
                      ]}
                      onChange={this._onChangeCategoryType}
                      label="Category Type :"
                      required={true}                      
                    />
                  </div>                                   
                </div>}
            </div>
          </Panel>
        </div>
      </div>
    </div>
    );
  }

  //Show selected section in Managed Assest menu
  public _onLinkClickAssets = (event, item) => {
    if (item.name == constant.SlidesDisplayName) {
      this._getDocsNodeSlidesItems();
    } else if (item.name == constant.ImagesDisplayName) {
      this._getDocsNodeImageItems();
    } else if (item.name == constant.TextSnippetDisplayName) {
      this._getDocsNodeTextSnippetItems();
    }
    else {
      this._getDocsNodeCategoryItems();
    }
  }

  //Show/Hide panel of add and edit item
  private _hidePanel = (): void => {
    this.setState({ showPanel: false, addOrEditBtn: false });
  };

  //Input change text in Title text 
  _handleChangeTitle = (event) => {
    this.setState({
      itemTitle: event,
    });
  };

  //Input change text in Discription text
  _handleChangeDiscript = (event) => {
    this.setState({
      itemDiscription: event,
    });
  };

  //Input change text in category text
  _handleChangeCategoryTitle = (event) => {
    this.setState({
      newCategoryTitle: event,
    });
  };

  //Show dialog box for validation of deleting item
  private _showDialog(item: any) {
    this.setState({
      hideDialog: false,
      deleteItem: item
    });
  }

  //Close the dialog box
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  };

  //Rendering Save and cancel button at footer in panel bar
  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton className={styles.footerbtn} onClick={this._addNewData}>Save</PrimaryButton>
        <DefaultButton className={styles.footerbtn} onClick={this._hidePanel}>Cancel</DefaultButton>
      </div>
    );
  };

  //Search the item from Displaying items List Component
  private _onChangeText = (text: string): void => {

    //Binding the search data
    this.setState({
      items: text ? this._SearchResult.filter(i => i.Title.toLowerCase().indexOf(text) > -1) : this._SearchResult
    });
  };

  private _onChangeCategoryType(event, option: any) {
    this._defaultRadioBtnCategoryType = option.key;
    this.setState({
      radioBtnCategory: option.key
    });
  }

  //Following function is use to add new item or edit the existing item 
  public _addNewData() {
    var titleValue = '';
    var discriptionValue = '';
    var newFile = '';
    var ParentLevel = 0;
    if (this.state.listName == constant.DocsNodeCategoriesName) {
      titleValue = document.getElementById('newCategoryName')['value'];

      if (titleValue != undefined && this._defaultRadioBtnCategoryType != '') {
        
        if(this.state.renderKey != ''){
          ParentLevel = ParentLevel + 1;
        }
        //Add and Edit list Item
        DC._updateListItem(titleValue, discriptionValue, this.state.renderKey, this.state.listName, this.state.addOrEditBtn, this.state.itemId, this._defaultRadioBtnCategoryType,ParentLevel).then((dataResult) => {

          //Rebind the data
          this._rebindData();
        });
      }else{
        alert('Filled required fields!!');
      }
      
    } else if (this.state.listName != constant.DocsNodeTextName) {
      titleValue = document.getElementById('titleID')['value'];
      discriptionValue = document.getElementById('discription')['value'];
      if (this.state.addOrEditBtn) {
        //Edit library item
        DC._uploadFiles('', titleValue, discriptionValue, this.state.renderKey, this.state.itemName, this.props.context, this.state.listName).then((dataResult) => {

          //Rebind the data
          this._rebindData();
        });
      } else {
        //Add library Item
        var flag = false;
        newFile = document.getElementById('inputTypeFiles')['files'][0];
        if (this.state.tabValue == constant.SlidesDisplayName) {
          //Validation for Slide Library only pptx files are uploaded
          if (newFile['name'].includes('.pptx') || newFile['name'].includes('.ppt')) {
            flag = true;
          } else {
            alert('Upload Presenation file with format .pptx');
            flag = false;
          }
        } else {
          //Validation for Picture Library only image files are uploaded
          if (newFile['name'].includes('.png') || newFile['name'].includes('.jpeg') || newFile['name'].includes('.jpg')) {
            flag = true;
          } else {
            alert('Upload Image with format .png,.jpg,.jpeg');
            flag = false;
          }
        }
        if (flag != false) {
          if (this.state.itemTitle.length > 0 && newFile != undefined) {
            DC._uploadFiles(newFile, titleValue, discriptionValue, this.state.renderKey, '', this.props.context, this.state.listName).then((dataResult) => {

              //Rebind the data
              this._rebindData();
            });
          }else{
            alert('Filled required fields!!');
          }
        }
      }
    } else {
      titleValue = document.getElementById('titleID')['value'];
      discriptionValue = document.getElementById('discription')['value'];

      //Add and Edit list Item
      DC._updateListItem(titleValue, discriptionValue, this.state.renderKey, this.state.listName, this.state.addOrEditBtn, this.state.itemId, this._defaultRadioBtnCategoryType,ParentLevel).then((dataResult) => {

        //Rebind the data
        this._rebindData();
      });
    }
  }

  //This function is use to get item from library which is going to be edited
  public _editItem(item: any) {
    //Calling the item to get edited
    DC._getLibraryItemToEdit(item, this.state.listName).then((itemData) => {
      this._defaultRadioBtnCategoryType = itemData[0].CategoryType;
      this.setState({
        itemTitle: itemData[0].Title,
        newCategoryTitle: itemData[0].Name,
        itemDiscription: itemData[0].Discription,
        itemName: itemData[0].Name,
        itemId: item.Key,
        radioBtnCategory: itemData[0].CategoryType,
        showPanel: true,
        renderKey: itemData[0].CategoryKey,
        addOrEditBtn: true
      });
    });
  }

  //This function is use to delete item from List and Library
  public _deleteItem(item: any) {

    //Close Dialog
    this._closeDialog();

    //Calling the api to delete item
    DC._deleteListItem(item, this.state.listName).then((itemData) => {

      //Rebind the data
      this._rebindData();
    });
  }

  //This function is use call for rebinding the data after add,edit and delete any edit
  public _rebindData() {
    if (this.state.tabValue == constant.SlidesDisplayName) {
      this._getDocsNodeSlidesItems();
    } else if (this.state.tabValue == constant.ImagesDisplayName) {
      this._getDocsNodeImageItems()
    } else if (this.state.tabValue == constant.TextSnippetDisplayName) {
      this._getDocsNodeTextSnippetItems();
    } else {
      this._getDocsNodeCategoryItems();
    }
  }

  //Following function is use to load the data  synchronously(Lazy loading)
  private _loadData = (): void => {

    //Set time inteval to display data
    setInterval(() => {
      const randomQuantity: number = Math.floor(Math.random() * 10) + 1;
      const itemsCopy = this.state.items!.slice(0);
      itemsCopy.splice(
        this._lastIndexWithData,
        randomQuantity,
        ...this._items.slice(this._lastIndexWithData, this._lastIndexWithData + randomQuantity)
      );
      this._lastIndexWithData += randomQuantity;
      this.setState({
        items: itemsCopy
      });
    }, _INTERVAL_DELAY);
  };

  //This function is use to divide N number of data in parts to display in List Component
  private _onLoadData = (itemLength,itemData): void => {
    
    this._ITEMS_COUNT = itemLength;
    this._items = itemData;
    
    let items = [];
    if (this._ITEMS_COUNT > 10) {
      //dividing data ramdomly 
      const randomQuantity: number = Math.floor(Math.random() * 10) + 1;
      items = this._items.slice(0, randomQuantity).concat(new Array(this._ITEMS_COUNT - randomQuantity));
      this._lastIndexWithData = randomQuantity;

      //Lazy Loading
      this._loadData();
      this.setState({
        isDataLoaded: true,
        items: items
      });
    } else {
      this.setState({
        isDataLoaded: true,
        items: this._items
      });
    }

  };

  //This function is use for sorting data on clicking header of the column
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    //reorganize the data in acending or decending
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  };

  //This function is use to acending or decending the data
  public _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  //This function is use for selection N Number of rows
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  //This function is use to show error message if user didn't input text in Title of Choose file to upload
  private _getErrorMessage = (value: string): string => {
    return value.length > 0 ? '' : 'Required *';
  };
}
