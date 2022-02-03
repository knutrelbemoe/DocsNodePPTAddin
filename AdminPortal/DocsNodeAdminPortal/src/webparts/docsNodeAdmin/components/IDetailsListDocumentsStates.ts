//State for DocsNode Temaplates
export interface IDetailsListDocumentsState {
  columns: any;
  categoryColumns: any;
  items: IDocument[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  CategaryDropDown: any;
  isDataLoaded: boolean;
  tabValue: string;
  defaultKey: string;
  showPanel: boolean;
  defaultNavSelectdKey: string;
  itemDiscription: any;
  itemTitle: any;
  itemName: any;
  addOrEditBtn: boolean;
  listName: string;
  itemId: number;
  renderKey: string;
  hideDialog: boolean;
  deleteItem: any;
  showList: boolean;
  categoryListForm: boolean;
  newCategoryTitle: string;
  radioBtnCategory: string;
}
//State for columns of list
export interface IDocument {
  Title: string;
  Category: string;
  FileType: any;
  LinkURL: string;
  Key: number;
  CategoryType: string;
  ParentCategory: string;
}