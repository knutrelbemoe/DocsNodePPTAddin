//Importing files and objects creation
import CommonUtility from "./CommonUtility";
import constant from "./Constant";
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from "@microsoft/sp-http";

//Creating object of existing files
const CU: CommonUtility = new CommonUtility();

//Array for List and Library is to be created by default
let dbConfigArr = [];
dbConfigArr.push({ listname: constant.DocsNodeCategoriesName, type: constant.NewList, BaseTemplate: 100 });
dbConfigArr.push({ listname: constant.DocsNodeSlidesName, type: constant.NewLibrary, BaseTemplate: 101 });
dbConfigArr.push({ listname: constant.DocsNodeTextName, type: constant.NewList, BaseTemplate: 100 });
dbConfigArr.push({ listname: constant.DocsNodePictureName, type: constant.NewLibrary, BaseTemplate: 109 });

export default class DatabaseConfiguration {

    //Check weather List or Library is exist or Not
    public async _checkListExistsOrNot() {
        try {
            for (var i = 0; i < dbConfigArr.length; i++) {
                await CU._getRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/Web/Lists?$filter=title eq ('" + dbConfigArr[i].listname + "')")
                    .then(async (data: any) => {
                        if (data.d.results.length > 0) {
                            console.log("List exists");
                        }
                        else {
                            var spDefaultMetadata = {
                                List: JSON.stringify({
                                    BaseTemplate: dbConfigArr[i].BaseTemplate,
                                    __metadata: { type: "SP.List" },
                                    Title: dbConfigArr[i].listname
                                }),
                                Document: JSON.stringify({
                                    __metadata: { type: "SP.List" },
                                    AllowContentTypes: true,
                                    BaseTemplate: dbConfigArr[i].BaseTemplate,
                                    ContentTypesEnabled: true,
                                    Title: dbConfigArr[i].listname
                                })
                            };
                            if (dbConfigArr[i].type == constant.NewList) {
                                //Creating New List
                                await this._createNewListOrLibrary(spDefaultMetadata.List);
                            }
                            else {
                                //Creating New Library
                                await this._createNewListOrLibrary(spDefaultMetadata.Document);
                            }
                        }
                    });
            }
            return ("Success");
        } catch (error) {
            console.log("checkListExistsOrNot: " + error);
            return ("Fail");
        }
    }

    //Creating new List or Library 
    public async _createNewListOrLibrary(listDetails: any) {
        try {
            //Await for POST request
            await CU._postRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/", listDetails, '')
                .then((data: any) => {
                    console.log("New List Created");
                });
        } catch (error) {
            console.log("createNewList: " + error);
        }
    }

    //Creating Column Array for List and Library
    public async _checkForColumnExistence() {
        try {
            for (var i = 0; i < dbConfigArr.length; i++) {
                switch (dbConfigArr[i].listname) {
                    case constant.DocsNodeCategoriesName:
                        let docsNodeCatColm = [];
                        docsNodeCatColm.push({ listName: dbConfigArr[i].listname, columnName: constant.CategoryName, FieldTypeKind: 2, EnforceUniqueValues: true, Indexed: true });
                        docsNodeCatColm.push({ listName: dbConfigArr[i].listname, columnName: constant.CategoryParentId, FieldTypeKind: 9, EnforceUniqueValues: false, Indexed: false });
                        docsNodeCatColm.push({ listName: dbConfigArr[i].listname, columnName: constant.CategoryType, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodeCatColm.push({ listName: dbConfigArr[i].listname, columnName: constant.CategoryLevel, FieldTypeKind: 9, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeCatColm);
                        break;
                    case constant.DocsNodeSlidesName:
                        let docsNodeSlideColm = [];
                        docsNodeSlideColm.push({ listName: dbConfigArr[i].listname, columnName: constant.SlidesCategoryName, FieldTypeKind: 7, EnforceUniqueValues: false, Indexed: false });
                        docsNodeSlideColm.push({ listName: dbConfigArr[i].listname, columnName: constant.SlidesDiscriptionName, FieldTypeKind: 3, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeSlideColm);
                        break;
                    case constant.DocsNodePictureName:
                        let docsNodeImgColm = [];
                        docsNodeImgColm.push({ listName: dbConfigArr[i].listname, columnName: constant.ImageCategoryName, FieldTypeKind: 7, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeImgColm);
                        break;
                    case constant.DocsNodeTextName:
                        let docsNodetxtColm = [];
                        docsNodetxtColm.push({ listName: dbConfigArr[i].listname, columnName: constant.TextSnippetName, FieldTypeKind: 3, EnforceUniqueValues: false, Indexed: false });
                        docsNodetxtColm.push({ listName: dbConfigArr[i].listname, columnName: constant.TextCategoryName, FieldTypeKind: 7, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodetxtColm);
                        break;
                    default:
                        break;
                }
            }
        } catch (error) {
            console.log("checkForColumnExistence: " + error);
        }
    }

    //Check weather Column exist or not in List or Library
    public async _checkColumn(columnData: any) {
        try {
            for (var i = 0; i < columnData.length; i++) {
                var flag = false;
                //GET request
                await CU._getRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/Web/Lists/GetByTitle('" + columnData[i].listName + "')/Fields")
                    .then(async (data: any) => {
                        if (data.d.results.length > 0) {
                            for (var j = 0; j < data.d.results.length; j++) {
                                if (data.d.results[j].Title == columnData[i].columnName) {
                                    console.log("Column exists");
                                    flag = true;
                                }
                            }
                            if (flag == false) {
                                //Creating new column
                                await this._createNewColumn(columnData[i]);
                            }
                        }
                    });
            }
        } catch (error) {
            console.log("checkColumn:" + error);
        }
    }
    
    //Create New column for List or Library
    public async _createNewColumn(columnData: any) {
        try {
            var colData = null;
            colData = JSON.stringify({
                __metadata: { 'type': 'SP.Field' },
                Description: 'Created From DocsNode',
                Title: columnData.columnName,
                FieldTypeKind: columnData.FieldTypeKind,
                EnforceUniqueValues: columnData.EnforceUniqueValues,
                Indexed: columnData.Indexed
            });
            if(columnData.columnName == constant.TextSnippetName){
                colData = JSON.stringify({
                    '__metadata': { 'type': 'SP.FieldMultiLineText' },
                    'Description': 'Created From DocsNode',
                    //SchemaXml: '<Field Type="Note" DisplayName="Feedback" Required="FALSE" EnforceUniqueValues="FALSE" Indexed="FALSE" NumLines="6" UnlimitedLengthInDocumentLibrary="TRUE" RichText="TRUE" RichTextMode="FullHtml" IsolateStyles="TRUE" Sortable="FALSE" StaticName="Feedback" Name="Feedback" ColName="ntext2" RowOrdinal="0" />',
                    'Title': columnData.columnName,
                    'FieldTypeKind': columnData.FieldTypeKind,
                    EnforceUniqueValues: columnData.EnforceUniqueValues,
                    Indexed: columnData.Indexed,                    
                     //'RestrictedMode': true,
                     'RichText': true,
                    //  'NumLines':6,
                    // 'RichTextMode':"FullHtml",
                     // 'IsolateStyles':true,
                     //  'AppendOnly':false
                    // RichTextMode: "FullHtml",
                    // //RestrictedMode: true,
                    // //Required:false,
                    // AppendOnly:false,
                    // Type:"Note",
                    // DisplayName:columnData.columnName,
                    // Format:"Dropdown", 
                    // IsolateStyles:true,
                    // Name:columnData.columnName, 
                    // NumLines:"6" , ID:"{728bd392-a9e0-4942-b7a9-ea3d09695g3g}", StaticName:"TextSnippet", ColName:"ntext2", RowOrdinal:"0" 
                });
            }
            if (columnData.FieldTypeKind != 7) {
                //POST request
                await CU._postRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/GetByTitle('" + columnData.listName + "')/Fields", colData, '')
                    .then((data: any) => {
                        console.log("New Column Created");
                    },(error)=>{
                        console.log(error);
                    });
            } else {
                //Creating lookup column in List or Library
                await this._createLookedUpCol(columnData);
            }

        } catch (error) {
            console.log("createNewColumn: " + error);
        }
    }

    //Create New lookup column in List or Library
    public async _createLookedUpCol(columnData: any) {
        try {
            //Get GUID of Existing List or Library
            await this._getListGUID(constant.DocsNodeCategoriesName).then(async (listId) => {
                var JSONVAR = "{ 'parameters': { '__metadata': { 'type': 'SP.FieldCreationInformation' }, 'FieldTypeKind': 7,'Title': '" + columnData.columnName + "', 'LookupListId': '" + listId + "' ,'LookupFieldName': '" + constant.CategoryName + "' } }";
                //POST request
                await CU._postRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/GetByTitle('" + columnData.listName + "')/Fields/addfield", JSONVAR, '')
                    .then((data: any) => {
                        console.log("New Column Created");
                    });
            });
        } catch (error) {
            console.log('createLookedUpCol : ' + error);
        }
    }

    //Get GUID of List or Library
    public _getListGUID(listName) {
        try {
            return fetch(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists?$filter=title eq ('" + listName + "')",
                {
                    headers: { Accept: 'application/json;odata=verbose' }, credentials: "same-origin"
                }).then((response) => {
                    return response.json();
                }, (errorJson) => {
                    console.log("checkListIfExistsOrNot: " + errorJson);
                }).then((data) => {
                    return data.d.results[0].Id;
                }).catch((error) => {
                    console.log("checkLibraryIfExistsOrNot: " + error);
                });
        }
        catch (error) {
            console.log("_getListGUID: " + error);
        }
    }

    //This function is use get all items from DocsNodeSlide Library
    public _getDocsNodeSlidesName() {
        try {
            var DocsNodeSlidesArrayItems = [];
            //GET request
            return CU._getRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/Lists/getbytitle('" + constant.DocsNodeSlidesName + "')/items?$select=ID,LinkFilename,DocIcon,Title," + constant.SlidesCategoryName + "/Title," + constant.SlidesCategoryName + "/Category,Editor/Name,Editor/Title,FileLeafRef,FileRef,ContentTypeId,ContentType/Id,ContentType/Name,*&$expand=" + constant.SlidesCategoryName + ",File,ContentType,Editor/Id")
                .then((ResponseData) => {
                    console.log(ResponseData);
                    if (ResponseData.d.results.length > 0) {
                        for (var i = 0; i < ResponseData.d.results.length; i++) {
                            if (ResponseData.d.results[i].ContentType.Name != 'Folder') {
                                DocsNodeSlidesArrayItems.push({ Key: ResponseData.d.results[i].ID, Title: ResponseData.d.results[i].LinkFilename, Category: ResponseData.d.results[i].SlidesCategory.Category, FileType: this._getIcon(ResponseData.d.results[i].DocIcon), LinkURL: ResponseData.d.results[i].File.LinkingUri });
                            }
                        }
                    }
                    console.log(DocsNodeSlidesArrayItems);
                    return DocsNodeSlidesArrayItems;
                }, (responseError) => {
                    console.log('getDocsNodeSlidesName inside getRequest : ' + responseError);
                    return DocsNodeSlidesArrayItems;
                });
        }
        catch (error) {
            console.log('getDocsNodeSlidesName : ' + error);
        }
    }

    //This function is use get all items from DocsNodeTextSnippet List
    public _getDocsNodeTextSnippetName() {
        try {
            var DocsNodeTextSnippetArrayItems = [];
            //GET request
            return CU._getRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/Lists/getbytitle('" + constant.DocsNodeTextName + "')/items?$select=ID," + constant.TextSnippetName + ",Title," + constant.TextCategoryName + "/Title," + constant.TextCategoryName + "/Category&$expand=" + constant.TextCategoryName + "")
                .then((ResponseData) => {
                    if (ResponseData.d.results.length > 0) {
                        for (var i = 0; i < ResponseData.d.results.length; i++) {
                            DocsNodeTextSnippetArrayItems.push({ Key: ResponseData.d.results[i].ID, Title: ResponseData.d.results[i].Title, Category: ResponseData.d.results[i].TextCategory.Category, FileType: '' });
                        }
                    }
                    return DocsNodeTextSnippetArrayItems;
                }, (responseError) => {
                    console.log('getDocsNodeTextSnippetName inside getRequest : ' + responseError);
                    return DocsNodeTextSnippetArrayItems;
                });
        }
        catch (error) {
            console.log('getDocsNodeTextSnippetName : ' + error);
        }
    }

    //This function is use get all items from DocsNodeCategory List
    public _getDocsNodeCategoriesName() {
        try {
            var DocsNodeCategoriesArrayItems = [];
            var DocsNodeParentCategoriesArrayItems = [];
            var DocsNodeCategoriesItemsData = [];
            //GET request
            return CU._getRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/Lists/getbytitle('" + constant.DocsNodeCategoriesName + "')/items?$select=ID,LinkFilename,DocIcon,Title,Category,*")
                .then((ResponseData) => {
                    if (ResponseData.d.results.length > 0) {  
                        for (var i = 0; i < ResponseData.d.results.length; i++) {
                            if(ResponseData.d.results[i].ParentCategory == null){
                                DocsNodeParentCategoriesArrayItems.push({ key: ResponseData.d.results[i].ID, text: ResponseData.d.results[i].Category });
                            }
                        }                      
                        for (var i = 0; i < ResponseData.d.results.length; i++) {                            
                            DocsNodeCategoriesArrayItems.push({ key: ResponseData.d.results[i].ID, text: ResponseData.d.results[i].Category });                                                     
                        }

                        ResponseData.d.results.map((item)=>{
                            var flag = false;
                            DocsNodeParentCategoriesArrayItems.map((itemParent) => {
                                if (itemParent.key == item.ParentCategory)
                                {
                                    DocsNodeCategoriesItemsData.push({Key: item.ID, Title: item.Category,CategoryType:(item.CategoryType != undefined? item.CategoryType: null),ParentCategory:itemParent.text, Category: (item.Category != undefined ? item.Category : null), FileType: this._getIcon(item.DocIcon)});
                                    flag =true;
                                }
                            });
                            if(flag == false){
                                DocsNodeCategoriesItemsData.push({Key: item.ID, Title: item.Category,CategoryType:(item.CategoryType != undefined? item.CategoryType: null),ParentCategory:item.ParentCategory, Category: (item.Category != undefined ? item.Category : null), FileType: this._getIcon(item.DocIcon)});
                            }
                        });
                        // for(var j = 0; j < DocsNodeParentCategoriesArrayItems.length; j++){
                        //     if(ResponseData.d.results[i].ParentCategory == DocsNodeParentCategoriesArrayItems[1].key){
                        //         DocsNodeCategoriesItemsData.push({Key: ResponseData.d.results[i].ID, Title: ResponseData.d.results[i].Category,CategoryType:(ResponseData.d.results[i].CategoryType != undefined? ResponseData.d.results[i].CategoryType: null),ParentCategory:DocsNodeParentCategoriesArrayItems[1].key, Category: (ResponseData.d.results[i].Category != undefined ? ResponseData.d.results[i].Category : null), FileType: this._getIcon(ResponseData.d.results[i].DocIcon)});
                        //     }else{
                        //         DocsNodeCategoriesItemsData.push({Key: ResponseData.d.results[i].ID, Title: ResponseData.d.results[i].Category,CategoryType:(ResponseData.d.results[i].CategoryType != undefined? ResponseData.d.results[i].CategoryType: null),ParentCategory:DocsNodeParentCategoriesArrayItems[1].key, Category: (ResponseData.d.results[i].Category != undefined ? ResponseData.d.results[i].Category : null), FileType: this._getIcon(ResponseData.d.results[i].DocIcon)});
                        //     }                            
                        // }                        
                    }
                    return ({DocsNodeCategoriesArrayItems,DocsNodeParentCategoriesArrayItems,DocsNodeCategoriesItemsData});
                }, (responseError) => {
                    console.log('getDocsNodeCategoriesName inside getRequest : ' + responseError);
                    return ({DocsNodeCategoriesArrayItems,DocsNodeParentCategoriesArrayItems});
                });
        } catch (error) {
            console.log('getDocsNodeCategoriesName : ' + error);
        }
    }

    //This function is use get all items from DocsNodePicture Library
    public _getDocsNodePictureName() {
        try {
            var DocsNodePictureArrayItems = [];
            //GET request
            return CU._getRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/Lists/getbytitle('" + constant.DocsNodePictureName + "')/items?$select=ID,Title,LinkFilenameNoMenu,DocIcon,EncodedAbsUrl," + constant.ImageCategoryName + "/Title," + constant.ImageCategoryName + "/Category&$expand=" + constant.ImageCategoryName + "")
                .then((ResponseData) => {
                    if (ResponseData.d.results.length > 0) {
                        for (var i = 0; i < ResponseData.d.results.length; i++) {
                            DocsNodePictureArrayItems.push({ Key: ResponseData.d.results[i].ID, Title: ResponseData.d.results[i].LinkFilenameNoMenu, Category: (ResponseData.d.results[i].ImageCategory != undefined ? ResponseData.d.results[i].ImageCategory.Category : null), FileType: this._getIcon(ResponseData.d.results[i].DocIcon) });
                        }
                    }
                    return DocsNodePictureArrayItems;
                }, (responseError) => {
                    console.log('getDocsNodePictureName inside getRequest : ' + responseError);
                    return DocsNodePictureArrayItems;
                });
        } catch (error) {
            console.log('getDocsNodePictureName : ' + error);
        }
    }

    //Upload File or image and adding or editing item in Library
    public async _uploadFiles(uploadFileObj, titleValue, discriValue,key, filename, context ,ListDisplayName) {
        try {
            if (uploadFileObj != '') {
                var file = uploadFileObj;
                if (file != undefined || file != null) {
                    let spOpts: ISPHttpClientOptions = {
                        headers: {
                            "Accept": "application/json",
                            "Content-Type": "application/json"
                        },
                        body: file,
                        credentials: "same-origin"
                    };                    
                    var url = CU.tenantURL() + CU.siteCollectionPath + "/_api/Web/Lists/getByTitle('" + ListDisplayName + "')/RootFolder/Files/Add(url='" + file.name + "', overwrite=true)";
                    return context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then(async (response: SPHttpClientResponse) => {
                        return response.json();
                    }).then((responseJSON: JSON) => {
                        //updating columns values of this item
                        return this._updateLibraryColValue(responseJSON['Name'], titleValue, discriValue,key,ListDisplayName);
                    });
                }
            } else {
                //updating columns values of this item
                return await this._updateLibraryColValue(filename, titleValue, discriValue,key,ListDisplayName);
            }
        }
        catch (error) {
            console.log('uploadFiles : ' + error);
        }
    }

    //Updating column properties of existing item
    public _updateLibraryColValue(responseName, titleValue, discriValue,key,ListDisplayName) {
        try {
            var itemURL = '';            
            itemURL = CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/getbytitle('" + ListDisplayName + "')/items?$select=FileRef,ID,LinkFilename&$filter=substringof('" + responseName + "',FileRef)";      
            //GET request
            return CU._getRequest(itemURL).then((responseData) => {
                var url = CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/getbytitle('" + ListDisplayName + "')/items('" + responseData.d.results[0].ID + "')";
                var commomJSON = null;
                if(key == ''){
                    key = null;
                }
                switch(ListDisplayName){
                    case constant.DocsNodeSlidesName:
                        commomJSON = JSON.stringify({
                            __metadata: { 'type': 'SP.Data.DocsNodeSlidesItem' },
                            SlidesDiscription: discriValue,
                            Title: titleValue,
                            SlidesCategoryId: key
                        });
                        break;
                    case constant.DocsNodePictureName:
                        commomJSON = JSON.stringify({
                            __metadata: { 'type': 'SP.Data.DocsNodePictureItem' },
                            Title: titleValue,
                            ImageCategoryId:key,
                            Description:discriValue
                        });
                        break;                    
                    default:
                        break;
                }     
                //POST request           
                return CU._postRequest(url, commomJSON, 'MERGE').then((data) => {
                    return data;
                });
            });
        } catch (error) {
            console.log("postRequest: " + error);
        }
    }

    //Adding or editing List items
    public _updateListItem(titleValue, discriValue,key,ListDisplayName,flag,itemID,catgyType,ParentLevel){
        try {
            var xMethod = '';
            var url ='';
            var  commomJSON = '';
            if(flag){
                url = CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/getbytitle('" + ListDisplayName + "')/items('"+itemID+"')"; 
            }else{
                url = CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/getbytitle('" + ListDisplayName + "')/items";
            }    
            if(key == ''){
                key = null;
            }
            if (ListDisplayName == constant.DocsNodeTextName) {
                commomJSON = JSON.stringify({
                    __metadata: { 'type': 'SP.Data.DocsNodeTextListItem' },
                    Title: titleValue,
                    TextSnippet: discriValue,
                    TextCategoryId: key
                });
            }else{
                commomJSON = JSON.stringify({
                    __metadata: { 'type': 'SP.Data.DocsNodeCategoriesListItem' },
                    Title: titleValue,                    
                    Category: titleValue,
                    ParentCategory:key,
                    CategoryType:catgyType,
                    CategoryLevel: ParentLevel
                });
            }
            
            if(flag){
                xMethod = 'MERGE';
            }
            //POST request
            return CU._postRequest(url, commomJSON, xMethod).then((data) => {
                return data;
            });
        } catch (error) {
            console.log('_updateListItem : '+ error);
        }
    }

    //Get the item for editing for Library or List
    public _getLibraryItemToEdit(itemResult,listname) {
        try{
            //GET request
            return CU._getRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/getbytitle('" + listname + "')/items('" + itemResult.Key + "')?$select=ID,LinkFilename,DocIcon,Title,*")
            .then((responseData) => {
                var editItemArray = [];
                switch(listname){
                    case constant.DocsNodeSlidesName:
                        editItemArray.push({Name:responseData.d.LinkFilename, Discription: responseData.d.SlidesDiscription, Title:responseData.d.Title, CategoryKey: responseData.d.SlidesCategoryId  });
                        break;
                    case constant.DocsNodePictureName:
                        editItemArray.push({Name:responseData.d.LinkFilename, Discription: responseData.d.Description,Title:responseData.d.Title, CategoryKey: responseData.d.ImageCategoryId});
                        break;
                    case constant.DocsNodeTextName:
                        editItemArray.push({ Name: responseData.d.Title, Discription: responseData.d.TextSnippet, Title: responseData.d.Title, CategoryKey: responseData.d.TextCategoryId });
                        break;
                    case constant.DocsNodeCategoriesName:
                        editItemArray.push({ Name: responseData.d.Category, Title: responseData.d.Category, CategoryKey: responseData.d.ParentCategory, CategoryType:responseData.d.CategoryType });
                        break;
                    default:
                        break; 
                }
                return editItemArray;
            });
        }
        catch(error){
            console.log('_getLibraryItemToEdit : '+ error);
        }        
    }

    //Delete the item from Library or List
    public _deleteListItem(itemResult,listname) {
        //POST request
        return CU._postRequest(CU.tenantURL() + CU.siteCollectionPath + "/_api/web/lists/getbytitle('" + listname + "')/items('" + itemResult.Key + "')", '', 'DELETE')
            .then((responseData) => {
                return responseData;
            });
    }

    //Get icons for document 
    public _getIcon(IconType: string): string {
        if (IconType != null) {
            switch (IconType) {
                case 'docx':
                    return String(require('../images/icon-docx.png'));
                case 'xlsx':
                    return String(require('../images/icon-xlsx.png'));
                case 'pptx':
                    return String(require('../images/icon-ppt.png'));
                case 'pdf':
                    return String(require('../images/icon-pdf.png'));
                case 'png':
                    return String(require('../images/icon-img.png'));
                default:
                    return String(require('../images/icon-other.png'));
            }
        }
    }

    public getUpdate(){
        var colData = JSON.stringify({
            __metadata: { 'type': 'SP.FieldMultiLineText' },
            RichText: true,
            RichTextMode:"FullHtml" ,
            IsolateStyles:"TRUE"
        });

        CU._postRequest("https://binaryrepublik516.sharepoint.com/sites/contentTypeHub/_api/web/lists/getbytitle('DocsNodeText')/fields('6f51b1fa-63a8-4cb4-9895-8b594ff5ef04')",colData,'MERGE')
        .then((data)=>{
            console.log(data);
        })
    }
}