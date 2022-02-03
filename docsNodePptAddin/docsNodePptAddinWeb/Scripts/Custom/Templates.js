var SPURL = "";
var GraphAPIToken = '';
var SPToken = '';
var templateServerRelURL = "/sites/DocsNodeAdmin/";
var DocsNodePictureDisplayName = '';
var DocsNodeSlideDisplayName = '';
var DocsNodeConfigurationDisplayName = 'DocsNodeConfiguration';
var imageListSite = '';
var DocsNodeCategoryDisplayName = '';
var twoBreadCrumbArr = [];
var oneBreadCrumbArr = [];
var imageMaxID = 0;
var imageParentMaxID = 0;
var slideMaxID = 0;
var slideParentMaxID = 0;
var imageItem = 20;
var slideItem = 14;
var BRTemplatesJS = window.BRTemplatesJS || {};
var imageSearchQuery = "";
var slideSearchQuery = "";
var DocNodePowerPoint;
var TokenArray;
var ImageListGUID = "";

var snippetSearch = "";
var imgSearch = "";

var IMGResultCount = 0;
var IMGFolderReset = 0;

// Added by Amartya for image and text snippet URL (22-Jan-2021) [Start]
var currentImagePath = "";
var GET_IMAGE_TEXT_SNIPPET_LIBRARY_URL = "https://docsnode-functions.azurewebsites.net/api/GetImgTxtSnippetLibrary?code=OeJukPSR2Xn3WKbUcWZXp0GxPrxIcnerIODVOipFw2aeKZjFluq0wg==";
var CREATE_TEXT_SNIPPET_URL = "https://docsnode-functions.azurewebsites.net/api/CreateTextSnippet?code=yF4qztOpHM//yYTtFdLmr3RTbWE20alRFM/5Sbfj/8FZ922EEKnLoQ==";
var GET_TEMPLATEIMAGE_URL = "https://docsnode-functions.azurewebsites.net/api/GetTemplateImage?code=xU1rFnkQfz3/IU27vcnYD5Hs/KbQFIUjAF4kqWiKYi/EFHr6zBYNJA==";
var FOLDER_IMAGE_URL = "./../../Images/icons/folder-open.jpg";
// Added by Amartya for image and text snippet URL (22-Jan-2021)    [End]

(function () {
    "use strict";
    Office.initialize = function (reason) {
        $(document).ready(function () {

            $("body").on("click", ".clickme a", function () {
                $('#emptydataTwo').hide();
                $('#emptydataOne').hide();
                $('#ManageTabsContent').show();
                $("#WaitDialog").hide();
                $("#WaitDialog").hide();
                $('.clickme a').removeClass('activelink');
                $(this).addClass('activelink');
                var tabid = $(this).data('tag');
                $('.list').removeClass('active').addClass('hide');
                $('#' + tabid).addClass('active').removeClass('hide');
                if (tabid == 'two') {
                    $('#btnRefresh').show();
                    $('#secondMainContent').css('display', 'block');
                    $("#txtSearch").attr("placeholder", "Search Corporate Images");
                    $("#txtSearch").val(imgSearch);
                    $('#mainContent').css('display', 'none');
                    if (($("#two ul.categoryList").html() != "" || $("#two ul.categoryItems").html() != "") && ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == "0" || $("#hdnCategory").val() == undefined)) {
                        $("#two ul.categoryList").show();
                        $("#two ul.categoryItems").show();
                        $("#two ul.childCategoryList").hide();
                        $("#two ul.childCategoryItems").hide();
                    } else {

                        $('#txtSearch').val("");
                        slideSearchQuery = "";
                        imageSearchQuery = "";

                        imageMaxID = 0;
                        var $_category = $("#hdnCategory").val();
                        var $_categoryName = $("#hdnCategoryName").val();
                        if ($_category != '' && $_categoryName != '') {
                            LoadCategoryandImagesOnFolderClick($_category, $_categoryName, false);
                        } else {
                            LoadImageCategoryandItemsOnPageLoad();
                        }
                    }
                }
                if (tabid == 'one') {
                    $('#btnRefresh').show();
                    $('#secondMainContent').css('display', 'block');
                    $("#txtSearch").attr("placeholder", "Search Slide");
                    $("#txtSearch").val(snippetSearch);
                    $('#mainContent').css('display', 'none');
                    if (($("#one ul.categoryList").html() != "" || $("#one ul.categoryItems").html() != "") && ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == "0" || $("#hdnTextCategory").val() == undefined)) {
                        $("#one ul.categoryItems").show();
                        $("#one ul.categoryList").show();
                        $("#one ul.childCategoryList").hide();
                        $("#one ul.childCategoryItems").hide();
                    }
                    else {
                        $('#txtSearch').val("");
                        slideSearchQuery = "";
                        imageSearchQuery = "";
                        slideMaxID = 0;
                        var $_category = $("#hdnTextCategory").val();
                        var $_categoryName = $("#hdnTextCategoryName").val();
                        if ($_category != '' && $_categoryName != '') {
                            LoadCategoryandSlideOnFolderClick($_category, $_categoryName, false);
                        } else {
                            LoadCategoryandSlideOnFolderClick(null, '', false);
                        }
                    }
                }
                if (tabid == 'zero') {
                    $('#secondMainContent').css('display', 'none');
                    $('#mainContent').css('display', 'block');
                    $('#btnRefresh').hide();
                    $('#ManageTabsContent').hide();
                }
            });

            $("body").on("click", "#one .liTitle", function () {
                $("#txtSearch").val(snippetSearch);
                $('#emptydataOne').hide();
                var $_this = $(this).attr("attr-category");
                var catName = $(this).attr('attr-title');
                LoadCategoryandSlideOnFolderClick($_this, catName, true);
            });

            $("body").on("click", "#two .liTitle", function () {
                $('#emptydataTwo').hide();
                var $_this = $(this).attr("attr-category");
                var catName = $(this).attr('attr-title');
                LoadCategoryandImagesOnFolderClick($_this, catName, true);
            });

            var activeLi = localStorage.getItem("active");

            //on keydown, clear the countdown 
            $('#txtSearch').on('keypress', function (e) {
                if (e.which == 13) {
                    $('.list.SideContentBox.active').scrollTop(0);
                    setTimeout(function () {
                        if ($(".SideContentBox.active").attr("id") == "one") {
                            snippetSearch = $("#txtSearch").val();
                            var $_category = $("#hdnTextCategory").val();
                            if ($("#txtSearch").val().trim() != "") {
                                searchSlides($("#txtSearch").val().trim(), $_category);
                            }
                            else {
                                clearSearchText();
                            }
                        }
                        if ($(".SideContentBox.active").attr("id") == "two") {
                            var $_category = $("#hdnCategory").val();
                            imgSearch = $("#txtSearch").val();
                            if ($("#txtSearch").val().trim() != "") {
                                seachImages($("#txtSearch").val().trim(), $_category);
                            }
                            else {
                                clearSearchText();
                            }
                        }
                    }, 100);
                }
            });

            $("body").on("click", "#clearSearch", function () {
                $('.list.SideContentBox.active').scrollTop(0);
                setTimeout(function () {
                    refresh();
                }, 100);

            });
            //$('.list.SideContentBox').off('scroll');
            //$('.list.SideContentBox').on('scroll', function () {
            //    if ($(this).scrollTop() + $(this).innerHeight() >= $(this)[0].scrollHeight) {
            //        if ($('.list.SideContentBox.active').attr("id") == "two") // Image section
            //        {
            //            loadImagesOnScroll();
            //        }
            //        if ($('.list.SideContentBox.active').attr("id") == "one") // Image section
            //        {
            //            loadTextOnScroll();
            //        }
            //    }
            //});
            $('#emptydataTwo').hide();
            $("body").on("click", "#btnRefresh", function () {
                $('.list.SideContentBox.active').scrollTop(0);
                IMGFolderReset = 0;
                setTimeout(function () {
                    refresh();
                }, 100);

            });
        });
    };
})();
BRTemplatesJS.Config = function () {
    var utility = new BRDocsNodeJS.postTokens();
    TokenArray = utility.callFunction();
    SPToken = TokenArray[0];
    GraphAPIToken = TokenArray[1];
    SPURL = TokenArray[2];

    setInterval(function () {
        TokenArray = utility.callFunction();
        SPToken = TokenArray[0];
        GraphAPIToken = TokenArray[1];
        SPURL = TokenArray[2];
        console.log("Refresh Tokens template");
    }, 540001);
    GetConfigurations();
};
function GetConfigurations() {

    //DocsNodeCategoryDisplayName = localStorage.getItem('CategoryListGUID');
    //DocsNodePictureDisplayName = localStorage.getItem('ImageLibraryListGUID');
    //DocsNodeSlideDisplayName = localStorage.getItem('SlidesLibraryListGUID');
    //imageListSite = localStorage.getItem('ImageSourceListPath');
    //ImageListGUID = imgList[0].ConfigSourceListGUID;
    //LoadImageCategoryandItemsOnPageLoad();

    var url = SPURL + templateServerRelURL + "/_api/web/lists/getbytitle('DocsNodeConfiguration')/items?$select=ConfigAssestTitle,ConfigSourceList,ConfigSourceListGUID,ConfigSourceListPath,*";
    $.ajax({
        url: url,
        method: "GET",
        async: false,
        headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken }
    }).then(function (result) {

        var catList = _.filter(result.d.results, function (itm) { return itm.ConfigAssestTitle == 'Category List' });
        var textList = _.filter(result.d.results, function (itm) { return itm.ConfigAssestTitle == 'Text Snippet List' });
        var imgList = _.filter(result.d.results, function (itm) { return itm.ConfigAssestTitle == 'Images Library' });

        CategoryListGUID = catList[0].ConfigSourceListGUID;
        SnippetLibraryName = textList[0].ConfigSourceDisplayListName;
        TemplateLibraryDisplayName = textList[0].ConfigSourceListGUID;
        TextLibSourceListPath = textList[0].ConfigSourceListPath;
        ImageListGUID = imgList[0].ConfigSourceListGUID;
        imageListSite = imgList[0].ConfigSourceListPath;

        LoadImageCategoryandItemsOnPageLoad();

    }).fail(function (data) {
        console.log(JSON.stringify(data));
    });
}

function refresh() {
    if ($(".list.SideContentBox.active").attr("id") == "one" || $(".list.SideContentBox.active").attr("id") == "two") {
        $('#txtSearch').val("");
        $('#emptydataTwo').hide();
        $('#emptydataOne').hide();
        slideSearchQuery = "";
        imageSearchQuery = "";
        clearSearchText();
    }
}

// Added by Amartya for image loading (22-Jan-2021) [Start]

function loadImageAndTextFromLibrary(path, library) {
    var dfd = $.Deferred();
    if (SPToken) {
        let payload = {
            "tenant": ORG_TENANT.id, //"641333af-c280-4e39-8f0c-1a52f0be8dc7",
            "FolderPath": path,
            "SPOUrl": ORG_ROOT_WEB.webUrl,
            "Title": library
        }

        $.ajax({
            url: GET_IMAGE_TEXT_SNIPPET_LIBRARY_URL,
            beforeSend: function (request) {
                request.setRequestHeader("Accept", "application/json; odata=verbose");
            },
            dataType: "json",
            headers: {
                'Authorization': 'Bearer ' + SPToken,
            },
            type: "POST",
            contentType: "application/json",
            data: JSON.stringify(payload),
        }).done(function (response) {
            dfd.resolve(response);

        }).fail(function (error) {
            console.log('Error in loadImageAndTextFromLibrary :- ', error);
            dfd.reject(error.responseText);
        });
    }
    return dfd.promise();
}

function populateImageFolderView(path, isFolderClicked) {

    $("#WaitDialog").show();

    if (isFolderClicked) {
        currentImagePath += !path ? "/" : currentImagePath === "/" ? path : "/" + path;
    }
    else {
        currentImagePath = path;
    }

    var imageBreadCrumbArray = currentImagePath.split("/");
    imageBreadCrumbArray = imageBreadCrumbArray[imageBreadCrumbArray.length - 1] ?
        imageBreadCrumbArray : imageBreadCrumbArray.slice(0, 1);
    var imagePathSpan = "";

    imageBreadCrumbArray.map(function (eachFolder, index) {

        var folderPath = imageBreadCrumbArray.slice(0, index + 1).join("/");
        folderPath = folderPath ? folderPath : "/";

        if (imageBreadCrumbArray.length === 1) {

            imagePathSpan += "<span path='" + folderPath + "' class='breadcrumb-span'>" +
                "Home" +
                "</span>";
        }
        else if (index < imageBreadCrumbArray.length - 1) {

            imagePathSpan += "<span path='" +
                folderPath +
                "' class='breadcrumb-span'>" +
                (!eachFolder ? "<i class='ms-Icon ms-Icon--HomeSolid' aria-hidden='true'></i>" : eachFolder) +
                "<i class='ms-Icon ms-Icon--ChevronRightMed' aria-hidden='true'></i>" +
                "</span>";
        }
        else {
            imagePathSpan += "<span path='" +
                folderPath +
                "' class='breadcrumb-span'>" +
                eachFolder +
                "</span>";
        }
    })

    loadImageAndTextFromLibrary(currentImagePath, "Images Library").then(function (response) {
        $("#WaitDialog").show();
        var $_ul = $("#two ul.categoryItems");
        $_ul.html("");
        $_ul.addClass("boxOfTemplate")
        $_ul.append("<div class='breadcrumb-div'>" + imagePathSpan + "</div>");

        $('#two ul.categoryItems .breadcrumb-span').each(function (index, item) {
            $(this).unbind('click');
            $(this).bind({
                click: function () {
                    populateImageFolderView($(this).attr('path'), false);
                },
            });
        });


        if (response.Files.results.length == 0 && response.Folders.results.length == 0) {
            $_ul.append("<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>");
            IMGFolderReset = IMGFolderReset + 1;
        }
        else {
            // For Folders
            try {
                response.Folders.results
                    .filter(function (eachFolder) { return (eachFolder.Name !== "Forms") })
                    .map((folder, index) => {
                        // $_ul.append(
                        //     "<li title='" + (folder.Name) + "' class='liIteamImg'>" +
                        //     "<div class='liInnerImage liFolder'>" +
                        //     "<img class='folderImage' src='" + FOLDER_IMAGE_URL + "'></img>" +
                        //     "<div class='text'>" +
                        //     "<span>" + folder.Name + "</span>" +
                        //     "</div>" +
                        //     "</div>" +
                        //     "</a>" +
                        //     "</li>"
                        // );
                        if (folder.Name.indexOf("_") !== 0) {
                            $_ul.append(
                                "<li title='" + (folder.Name) + "' class='checkboxToggler' contenttypename='folder'>" +
                                "<div>" +
                                "<img class='docimgouterbox' src='" + FOLDER_IMAGE_URL + "'></img>" +
                                "<div class='text'>" +
                                "<span style='word-break: break-all'>" + folder.Name + "</span>" +
                                "</div>" +
                                "</div>" +
                                "</li>"
                            );
                        }
                    });

            } catch (error) {
                console.error("Error in Folder Rendering: ", error);
            }
            // For Text
            try {
                IMGResultCount = response.Files.results.length;
                IMGFolderReset = IMGFolderReset + 1;

                response.Files.results.forEach((image, index) => {
                    var imageURL = SPURL.concat(image.ServerRelativeUrl);
                    var filename = image.Name.substring(0, image.Name.lastIndexOf('.')) || image.Name;

                    toDataURL(imageURL, function (dataUrl) {
                        IMGResultCount--;
                        var base64result = "data:image/png;base64, " + dataUrl;
                        $_ul.append(
                            "<li title='" + (filename) +
                            "' class='checkboxToggler' " +
                            "onClick=\"insertImage('" + imageURL + "')\">" +
                            "<div class='box-img'>" +
                            "<img src='" + base64result + "' class='docimgouterbox'>" +
                            "</div>" +
                            `<span style="word-break: break-all" title="${filename}">${filename}</span>` +
                            "</li>");
                    }, IMGFolderReset);
                });

            } catch (error) {
                console.error("Error in File Rendering: ", error);

            }
            try {
                $('#two ul.categoryItems li').each(function (index, item) {
                    $(this).unbind('click');
                    $(this).bind({
                        click: function () {
                            populateImageFolderView(item.title, true);
                        },
                    });
                });

            } catch (error) {
                console.error("Error in Binding: ", error);

            }
        }
        $("#WaitDialog").hide();
    })

}

// Added by Amartya for image loading (22-Jan-2021) [End]

function LoadCategoryandSlidesOnPageLoad(filterCondition, catContainer, itemContainer, categoryId) {
    var $_ulList = catContainer;
    var $_ulItems = itemContainer;
    $('#emptydataOne').hide();
    new Promise(function (callback, reject) {
        var $_ul = catContainer;
        $('#WaitDialog').show();
        loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', filterCondition).then(function (data) {
            var results = data.results;
            if (results.length > 0) {
                $.each(data.results, function (index, value) {
                    var category = value.Category;
                    var str = "<li>" +
                        "<a href='#' class='liTitle' attr-title='" + category + "' attr-category='" + value.ID + "'>" +
                        "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                        "<span class='fileName' title='" + category + "'>" + (category.length > 20 ? category.substr(0, 20) + "..." : category) + "</span>" +
                        "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                        "</a >" +
                        "</li>"
                    $_ul.append(str);
                    $("#WaitDialog").hide();
                });
                $('#emptydataOne').hide();
            }
            callback(1);
        }, function (error) {
            var erroeMeg = JSON.parse(error.responseText);
            erroeMeg = erroeMeg["error"].message.value;
            docsTemplateList = "<div class='displayMessage'>" + erroeMeg + "\nNo Template Library found in DocsNode Admin Panel.</div>";
            $_ul.html(docsTemplateList);
            $("#WaitDialog").hide();
        });
    }).then(function (result) {
        var $_ul = itemContainer;
        $_ul.html("");
        loadDataFromSharePoint(DocsNodeSlideDisplayName, 'ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory', "(SlidesCategory eq " + categoryId + ")", "slide").then(function (data) {
            var results = data.results;
            var catLi = '';
            if (results.length > 0) {
                if (categoryId == null)
                    slideParentMaxID = _.max(_.pluck(results, "ID"));
                else
                    slideMaxID = _.max(_.pluck(results, "ID"));
                bindSlidesHTML(results, itemContainer);
            } else {
                if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
                    $('#emptydataOne').show();
                    $("#WaitDialog").hide();
                } else {
                    $('#emptydataOne').hide();
                }
            }
        });
    });
};

function clearSearchText() {
    if ($(".SideContentBox.active").attr("id") == "one") {
        snippetSearch = "";

        var $_category = $("#hdnTextCategory").val();
        var $_categoryName = $("#hdnTextCategoryName").val();
        $("#txtSearch").attr("placeholder", "Search Slides");
        $("#txtSearch").val('');
        if ($_category == "" || $_category == undefined || $_category == null) {
            slideParentMaxID = 0;
            $("#one ul.categoryItems").html("");
            $("#one ul.categoryList").html("");
            LoadCategoryandSlideOnFolderClick(null, "", false);
        }
        else {
            slideMaxID = 0;
            $("#one ul.childCategoryList").html("");
            $("#one ul.childCategoryItems").html("");
            LoadCategoryandSlideOnFolderClick($_category, $_categoryName, false);
        }
    }

    if ($(".SideContentBox.active").attr("id") == "two") {
        var $_category = $("#hdnCategory").val();
        var $_categoryName = $("#hdnCategoryName").val();
        imgSearch = '';
        $("#txtSearch").attr("placeholder", "Search Corporate Images");
        $("#txtSearch").val('');
        if ($_category == "" || $_category == undefined || $_category == null) {
            imageParentMaxID = 0;
            $("#two ul.categoryList").html("");
            $("#two ul.categoryItems").html("");
            LoadImageCategoryandItemsOnPageLoad(0);
        }
        else {
            imageMaxID = 0;
            $("#two ul.childCategoryList").html("");
            $("#two ul.childCategoryItems").html("");
            LoadCategoryandImagesOnFolderClick($_category, $_categoryName, false);
        }
    }
}

function searchSlides(searchParam, categoryId) {
    $("#WaitDialog").show();

    if (categoryId != "" && categoryId != "0") {

        slideMaxID = 0;
        $("#one ul.childCategoryList").html("");
        $("#one ul.childCategoryItems").html("");

        $("#one ul.categoryList").hide();
        $("#one ul.categoryItems").hide();
        $("#one ul.childCategoryList").hide();
        $("#one ul.childCategoryItems").show();

        var categoryList = [];
        categoryList.push(categoryId);

        loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Slides') and (ParentCategory eq " + categoryId + "))").then(function (data) {
            $.each(data.results, function (index, value) {
                categoryList.push(value.ID);
            });

            var searchString = "substringof('" + searchParam + "',Title) and ";
            var cateStr = "";
            if (categoryList.length > 1) {
                cateStr += "("
                for (var i = 0; i < categoryList.length; i++) {
                    if (i + 1 == categoryList.length)
                        cateStr += "(SlidesCategory/ID eq " + categoryList[i] + ")";
                    else
                        cateStr += "(SlidesCategory/ID eq " + categoryList[i] + ") or ";
                }
                cateStr += ")"

                slideSearchQuery = "(" + searchString + cateStr + ")";
            }
            else {
                slideSearchQuery = "(" + searchString + "(SlidesCategory/ID eq " + categoryId + "))";
            }
            slideSearchAjax($("#one ul.childCategoryItems"));
        });

    }
    else {
        slideSearchQuery = "(substringof('" + searchParam + "',Title))";
        slideParentMaxID = 0;
        $("#one ul.categoryList").html("");
        $("#one ul.categoryItems").html("");
        $("#one ul.categoryList").hide();
        $("#one ul.categoryItems").show();
        $("#one ul.childCategoryList").hide();
        $("#one ul.childCategoryItems").hide();
        slideSearchAjax($("#one ul.categoryItems"));
    }
}

function slideSearchAjax(container) {
    loadDataFromSharePoint(DocsNodeSlideDisplayName, 'ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory', slideSearchQuery, 'slide').then(function (data) {
        var results = data.results;
        if (results.length > 0) {
            if ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == "0" || $("#hdnTextCategory").val() == undefined) {
                slideParentMaxID = _.max(_.pluck(data.results, "ID"));
            }
            else
                slideMaxID = _.max(_.pluck(data.results, "ID"));
            bindSlidesHTML(results, container);
            $("#WaitDialog").hide();
        } else {
            if (container[0].childElementCount == 0) {
                $('#emptydataOne').show();
                $("#WaitDialog").hide();
            } else {
                $('#emptydataOne').hide();
            }
        }
    });
}

function seachImages(searchParam, categoryId) {

    $("#WaitDialog").show();
    if (categoryId != "" && categoryId != "0") {
        imageMaxID = 0;
        var itemSeachQuery;

        $(".childCategoryList").html("");
        $(".childCategoryItems").html("");

        $(".childCategoryList").show();
        $(".childCategoryItems").show();

        $(".categoryList").hide();
        $("#two ul.categoryItems").hide();
        var categoryList = [];
        categoryList.push(categoryId);
        if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
            loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (ParentCategory eq " + categoryId + "))", "image", true).then(function (data) {
                $.each(data.results, function (index, value) {
                    categoryList.push(value.ID);
                });
                var searchString = "substringof('" + searchParam + "',Title) and ";
                var cateStr = "";
                if (categoryList.length > 1) {
                    cateStr += "("
                    for (var i = 0; i < categoryList.length; i++) {
                        if (i + 1 == categoryList.length)
                            cateStr += "(ImageCategory/ID eq " + categoryList[i] + ")";
                        else
                            cateStr += "(ImageCategory/ID eq " + categoryList[i] + ") or ";
                    }
                    cateStr += ")"

                    itemSeachQuery = "(" + searchString + cateStr + ")";
                }
                else {
                    itemSeachQuery = "(" + searchString + "(ImageCategory/ID eq " + categoryId + "))";
                }
                //This is needed for lazy loading
                imageSearchQuery = itemSeachQuery;
                imageSearchAjax($("#two ul.childCategoryItems"));
            });
        }

    }
    else {

        var url = SPURL + imageListSite + "/_api/web/lists(guid'" + ImageListGUID + "')/items?$select=FileRef,FileLeafRef&$filter=substringof('" + searchParam + "',FileRef)";


        LoadSnippetFileConent(url).then(function (data) {

            var $_ul = $("#two ul.categoryItems");
            $_ul.html("");
            $_ul.addClass("boxOfTemplate")
            if (data.d.results.length > 0) {
                IMGResultCount = data.d.results.length;
                IMGFolderReset = IMGFolderReset + 1;
                $.map(data.d.results, function (image, index) {
                    var imageURL = SPURL + image.FileRef;
                    var filename = image.FileLeafRef.substring(0, image.FileLeafRef.lastIndexOf('.')) || image.FileLeafRef;
                    toDataURL(imageURL, function (dataUrl) {
                        IMGResultCount--;
                        var base64result = "data:image/png;base64, " + dataUrl;
                        $_ul.append(
                            "<li title='" + (filename) +
                            "' class='checkboxToggler' " +
                            "onClick=\"insertImage('" + imageURL + "')\">" +
                            "<div class='box-img'>" +
                            "<img src='" + base64result + "' class='docimgouterbox'>" +
                            "</div>" +
                            `<span style="word-break: break-all" title="${filename}">${filename}</span>` +
                            "</li>");

                    }, IMGFolderReset);
                });


        }
        else {

            $_ul.append("<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>");
                IMGFolderReset = IMGFolderReset + 1;
        }

        $("#WaitDialog").hide();
        //imageSearchQuery = "(substringof('" + searchParam + "',Title))";
        ////This is needed for lazy loading
        //imageParentMaxID = 0;
        //$("#two ul.categoryList").html("");
        //$("#two ul.categoryItems").html("");
        //$("#two ul.childCategoryList").hide();
        //$("#two ul.childCategoryItems").hide();
        //$("#two ul.categoryList").hide();
        //$("#two ul.categoryItems").show();
        //imageSearchAjax($("#two ul.categoryItems"));
        });
    }
}


    function LoadSnippetFileConent(fileUrl) {
        var dfd = $.Deferred();

        try {
            callAjaxGet(fileUrl).done(function (data) {
                dfd.resolve(data);
            });
        } catch (err) {
            dfd.reject(err);
        }

        return dfd;
    }
function imageSearchAjax(container) {
    loadDataFromSharePoint(ImageListGUID, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl', imageSearchQuery, "image", true).then(function (data) {

        if ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == "0" || $("#hdnCategory").val() == undefined) {
            imageParentMaxID = _.max(_.pluck(data.results, "ID"));
        }
        else
            imageMaxID = _.max(_.pluck(data.results, "ID"));

        if (data.results.length > 0) {
            $.each(data.results, function (index, image) {

                toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                    var base64result = "data:image/png;base64, " + dataUrl;
                    container.append("<li title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img src='" + base64result + "'>" +
                        "</div></a></li>");
                });
            });
        }
        else {
            container.append("<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>");
        }
        $("#WaitDialog").hide();
    });
}

function LoadImageCategoryandItemsOnPageLoad() {
    twoBreadCrumbArr = [];
    var $_breaCrumbUL = $("#two ul.breadcrumb");
    $_breaCrumbUL.html("");
    $("#two ul.breadcrumb").hide();
    $("#two ul.categoryList").html("");
    $("#two ul.categoryItems").html("");
    $("#two ul.categoryList").show();
    $("#two ul.categoryItems").show();
    var $_ulList = $('#two ul.categoryList');
    var $_ulItems = $('#two ul.categoryItems');

    $("#two ul.childCategoryList").hide();
    $("#two ul.childCategoryItems").hide();

    if (imageListSite.toLowerCase() == "/sites/docsnodeadmin" || imageListSite.toLowerCase() == "/sites/docsnodeadmin/") {
        //new Promise(function (callback, reject) {
        //    var $_ul = $("#two ul.categoryList");
        //    $("#WaitDialog").show();
        //    loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (CategoryLevel eq 0))").then(function (response) {
        //        var results = response.results
        //        if (results.length > 0) {
        //            $.each(results, function (index, value) {
        //                var category = value.Category;
        //                var str = "<li>" +
        //                    "<a href='#' class='liTitle' attr-title='" + category + "' attr-category='" + value.ID + "'>" +
        //                    "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
        //                    "<span class='fileName ' title='" + category + "'>" + (category.length > 20 ? category.substr(0, 20) + "..." : category) + "</span>" +
        //                    "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
        //                    "</a >" +
        //                    "</li>"
        //                $_ul.addClass("root");
        //                $_ul.append(str);
        //            });
        //        }
        //        $('#emptydataTwo').hide();
        //        $("#WaitDialog").hide();


        //        callback(1);
        //    }, function (error) {
        //        var erroeMeg = JSON.parse(error.responseText);
        //        erroeMeg = erroeMeg["error"].message.value;
        //        docsTemplateList = "<div class='displayMessage'>" + erroeMeg + "\nNo Template Library found in DocsNode Admin Panel.</div>";
        //        $_ul.html(docsTemplateList);
        //        $("#WaitDialog").hide();
        //    });
        //}).then(function (result) {
        //    var $_ul = $("#two ul.categoryItems");
        //    $_ul.html("");

        //    loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title,*&$expand=ImageCategory', "(ImageCategory eq null)", "image").then(function (response) {
        //        var results = response.results;
        //        if (results.length > 0) {
        //            imageParentMaxID = _.max(_.pluck(results, "ID"));
        //            $.each(results, function (index, image) {
        //                toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
        //                    var base64result = "data:image/png;base64, " + dataUrl;
        //                    $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
        //                        "</div></a></li>");
        //                    $("#WaitDialog").hide();
        //                });
        //            });
        //            $('#emptydataTwo').hide();
        //        } else {
        //            if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
        //                $('#emptydataTwo').show();
        //                $("#WaitDialog").hide();
        //            } else {
        //                $('#emptydataTwo').hide();
        //            }
        //        }
        //    });
        //});

        populateImageFolderView(currentImagePath, true);
    }
    else {

        populateImageFolderView(currentImagePath, true);

        // Commented by Amartya for image loading (22-Jan-2021) [Start]

        // var $_ul = $("#two ul.categoryItems");
        // $_ul.html("");
        // loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl', "", "image").then(function (response) {
        //     var results = response.results;
        //     if (results.length > 0) {
        //         imageParentMaxID = _.max(_.pluck(results, "ID"));
        //         $.each(results, function (index, image) {
        //             toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
        //                 var base64result = "data:image/png;base64, " + dataUrl;
        //                 $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
        //                     "</div></a></li>");
        //                 $("#WaitDialog").hide();
        //             });
        //         });
        //         $('#emptydataTwo').hide();
        //     } else {
        //         if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
        //             $('#emptydataTwo').show();
        //             $("#WaitDialog").hide();
        //         } else {
        //             $('#emptydataTwo').hide();
        //         }
        //     }
        // });
        
        // Commented by Amartya for image loading (22-Jan-2021) [End]
    }
}
function loadTextOnScroll() {

    var $_category = $("#hdnTextCategory").val();
    var $_ul = "";
    var filter = "";
    var cat = 0;

    if (slideSearchQuery != "") {
        filter = slideSearchQuery;
    }

    if ($_category == "" || $_category == undefined || $_category == null) {
        $_ul = $("#one ul.categoryItems");
    }
    else {
        cat = $_category;
        $_ul = $("#one ul.childCategoryItems");
    }

    getSlidestForLazyLoading(cat, filter).then(function (data) {
        if (cat == 0)
            slideParentMaxID = _.max(_.pluck(data.results, "ID"));
        else
            slideMaxID = _.max(_.pluck(data.results, "ID"));
        bindSlidesHTML(data.results, $_ul);
    });
}
function getSlidestForLazyLoading(categoryId, filter) {
    var dfd = $.Deferred();


    var cat = "";

    if (categoryId == 0)
        cat = "null";
    else
        cat = categoryId;

    var url = "";

    if (categoryId == 0)// It means parent level
    {
        if (filter == "")
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(SlidesCategory/ID eq " + cat + " and ID gt " + slideParentMaxID + ")&$top=" + slideItem;
        else
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(" + filter + " and ID gt " + slideParentMaxID + ")&$top=" + slideItem;
    }
    else {
        if (filter == "") {
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(SlidesCategory/ID eq " + cat + " and ID gt " + slideMaxID + ")&$top=" + slideItem;
        }
        else {
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + DocsNodeSlideDisplayName + "')/items?$select=ID,Title,LinkFilename,EncodedAbsUrl,SlidesDiscription,SlidesCategory/ID,SlidesCategory/Title&$expand=SlidesCategory&$filter=(" + filter + " and ID gt " + slideMaxID + ")&$top=" + slideItem;
        }
    }


    try {
        callAjaxGet(url).done(function (data) {
            dfd.resolve(data.d);
        });
    } catch (err) {
        console.log(err);
        dfd.reject(err);
    }

    return dfd;
}
function loadImagesOnScroll() {
    var $_category = $("#hdnCategory").val();
    var $_ul = "";
    var filter = "";
    if (imageSearchQuery != "") {
        filter = imageSearchQuery;
        if ($_category == "" || $_category == undefined || $_category == null) {
            $_ul = $("#two ul.categoryItems");
        }
        else {
            $_ul = $("#two ul.childCategoryItems");
        }
    }
    else {
        if ($_category == "" || $_category == undefined || $_category == null) {
            filter = "(ImageCategory eq null)";
            $_ul = $("#two ul.categoryItems");
        }
        else {
            filter = "(ImageCategory/ID eq " + $_category + ")";
            $_ul = $("#two ul.childCategoryItems");
        }
    }

    loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory', filter, "image").then(function (data) {
        var results = data.results;
        if (results.length > 0) {
            if ($_category == "" || $_category == undefined || $_category == null) {
                imageParentMaxID = _.max(_.pluck(data.results, "ID"));
            }
            else {
                imageMaxID = _.max(_.pluck(data.results, "ID"));
            }
            $.each(results, function (index, image) {
                toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                    var base64result = "data:image/png;base64, " + dataUrl;
                    $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
                        "</div></a></li>");
                });
            });
        }
    });
}

function LoadCategoryandImagesOnFolderClick($_this, $_categoryName, $_handleBreadcrumb) {
    $("#two ul.childCategoryList").html("");
    $("#two ul.childCategoryItems").html("");
    var $_category = $_this;
    //Added for Breadcrumb by Darshana
    if ($_handleBreadcrumb) {
        //console.log("Array:", twoBreadCrumbArr);
        if ($_category != 0) {
            twoBreadCrumbArr.push({ id: $_category, category: $_categoryName });
            if (twoBreadCrumbArr.length > 0) {
                $("#two ul.breadcrumb").show();
            }
        }
        var $_breaCrumbUL = $("#two ul.breadcrumb");
        $_breaCrumbUL.html("");
        // Append to breadcrumb
        for (var key in twoBreadCrumbArr) {
            if (twoBreadCrumbArr.hasOwnProperty(key)) {
                var catID = key == 0 ? 0 : twoBreadCrumbArr[key - 1].id;
                $_breaCrumbUL.append('<li title="' + twoBreadCrumbArr[key].category + '"><a class="ms-Breadcrumb-itemLink" data-id="' + key + '" attr-category="' + catID + '"><i class="ms-Breadcrumb-chevron ms-Icon ms-Icon--ChevronLeft"></i><label class="ms-Breadcrumb-itemLink">' + twoBreadCrumbArr[key].category + '</label></li>');
                $(document).off('click');
                $(document).on("click", "#two ul.breadcrumb li", function (event) {
                    $('#emptydataTwo').hide();
                    $("#txtSearch").val("");
                    imageSearchQuery = "";
                    $("#two ul.childCategoryList").html("");
                    $("#two ul.childCategoryItems").html("");
                    var index = parseInt($(this).find("a").attr("data-id"));
                    if (index < twoBreadCrumbArr.length) {
                        twoBreadCrumbArr = twoBreadCrumbArr.splice(0, twoBreadCrumbArr.length - (twoBreadCrumbArr.length - index));
                        if (twoBreadCrumbArr.length == 0) {
                            $("#two ul.breadcrumb").hide();
                        }
                        var sli = $("#two ul.breadcrumb li a").filter(function () {
                            return $(this).attr("data-id") >= index;
                        });
                        if ($(sli).length > 0) $(sli).parent().remove();
                        if (parseInt($(this).find("a").attr("data-id")) - 1 == -1) {
                            $("#hdnCategory").val("");
                            $("#hdnCategoryName").val("");

                            $("#two ul.categoryList").show();
                            $("#two ul.categoryItems").show();
                            $("#two ul.childCategoryList").hide();
                            $("#two ul.childCategoryItems").hide();

                        } else {
                            imageMaxID = 0;
                            $("#hdnCategory").val($(this).find("a").attr("attr-category"));
                            $("#hdnCategoryName").val(twoBreadCrumbArr[index - 1].category);
                            loadImagesByCategory($(this).find("a").attr("attr-category"));
                        }
                    }


                });
            }
        }
    }
    loadImagesByCategory($_category);
    $("#WaitDialog").show();
    $("#hdnCategory").val($_category);
    $("#hdnCategoryName").val($_categoryName);
}

function LoadCategoryandSlideOnFolderClick($_this, $_categoryName, $_handleBreadcrumb) {
    $("#one ul.childCategoryList").html("");
    $("#one ul.childCategoryItems").html("");
    var $_category = $_this;
    if ($_handleBreadcrumb) {
        if ($_category != 0) {
            oneBreadCrumbArr.push({ id: $_category, category: $_categoryName });
            if (oneBreadCrumbArr.length > 0) {
                $("#one ul.breadcrumb").show();
            }
        }
        var $_breaCrumbUL = $("#one ul.breadcrumb");
        $_breaCrumbUL.html("");
        for (var key in oneBreadCrumbArr) {
            if (oneBreadCrumbArr.hasOwnProperty(key)) {
                var catID = key == 0 ? 0 : oneBreadCrumbArr[key - 1].id;
                var cateTitle = oneBreadCrumbArr[key]["category"];
                $_breaCrumbUL.append('<li title="' + cateTitle + '"><a class="ms-Breadcrumb-itemLink" data-id="' + key + '" attr-category="' + catID + '"><i class="ms-Breadcrumb-chevron ms-Icon ms-Icon--ChevronLeft"></i><label class="ms-Breadcrumb-itemLink">' + oneBreadCrumbArr[key].category + '</label></li>');
                $(document).on("click", "#one ul.breadcrumb li", function (event) {
                    $('#emptydataOne').hide();
                    $("#txtSearch").val("");
                    slideSearchQuery = "";
                    $("#one ul.childCategoryList").html("");
                    $("#one ul.childCategoryItems").html("");
                    var index = parseInt($(this).find("a").attr("data-id"));
                    if (index < oneBreadCrumbArr.length) {
                        oneBreadCrumbArr = oneBreadCrumbArr.splice(0, oneBreadCrumbArr.length - (oneBreadCrumbArr.length - index));
                        if (oneBreadCrumbArr.length == 0) {
                            $("#one ul.breadcrumb").hide();
                        }
                        var sli = $("#one ul.breadcrumb li a").filter(function () {
                            return $(this).attr("data-id") >= index;
                        });
                        if ($(sli).length > 0) $(sli).parent().remove();
                        if (parseInt($(this).find("a").attr("data-id")) - 1 == -1) {
                            oneBreadCrumbArr = [];
                            var $_breaCrumbUL = $("#one ul.breadcrumb");
                            $_breaCrumbUL.html("");
                            $("#hdnTextCategory").val("");
                            $("#hdnTextCategoryName").val("");

                            $("#one ul.categoryList").show();
                            $("#one ul.categoryItems").show();
                            $("#one ul.childCategoryList").hide();
                            $("#one ul.childCategoryItems").hide();
                        } else {
                            slideMaxID = 0;
                            $("#hdnTextCategory").val($(this).find("a").attr("attr-category"));
                            $("#hdnTextCategoryName").val(oneBreadCrumbArr[index - 1].category);
                            loadcategoryandSlides($(this).find("a").attr("attr-category"));
                        }
                    }
                });
            }
        }
    }
    loadcategoryandSlides($_category);
    $("#WaitDialog").show();
    $("#hdnTextCategory").val($_category);
    $("#hdnTextCategoryName").val($_categoryName);
}

function loadcategoryandSlides(categoryId) {
    $('#emptydataOne').hide();
    $("#WaitDialog").show();

    var catContainer = "";
    var itemContainer = "";
    var filterCondition = "";

    if (categoryId == null) {

        catContainer = $("#one ul.categoryList");
        itemContainer = $("#one ul.categoryItems");

        $("#one ul.categoryItems").show();
        $("#one ul.categoryList").show();
        $("#one ul.childCategoryList").hide();
        $("#one ul.childCategoryItems").hide();

        filterCondition = "((CategoryType eq 'Slides') and (CategoryLevel eq 0))";
    }
    else {
        $("#one ul.childCategoryList").html("");
        $("#one ul.childCategoryItems").html("");

        catContainer = $("#one ul.childCategoryList");
        itemContainer = $("#one ul.childCategoryItems");

        $("#one ul.categoryItems").hide();
        $("#one ul.categoryList").hide();
        $("#one ul.childCategoryList").show();
        $("#one ul.childCategoryItems").show();

        filterCondition = "((CategoryType eq 'Slides') and (ParentCategory eq " + categoryId + "))";
    }
    LoadCategoryandSlidesOnPageLoad(filterCondition, catContainer, itemContainer, categoryId);
}

function loadImagesByCategory($_category) {
    imageMaxID = 0;

    $("#two ul.categoryList").hide();
    $("#two ul.categoryItems").hide();

    $("#two ul.childCategoryList").show();
    $("#two ul.childCategoryItems").show();
    var $_ulList = $("#two ul.childCategoryList");
    var $_ulItems = $("#two ul.childCategoryItems");

    var results = null;
    try {
        new Promise(function (callback, reject) {
            var $_ul = $("#two ul.childCategoryList");
            $_ul.html("");
            var isCategoryAvailable = false;
            if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
                loadDataFromSharePoint(DocsNodeCategoryDisplayName, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (ParentCategory eq " + $_category + "))").then(function (data) {
                    results = data.results;
                    if (results.length > 0) {
                        $.each(results, function (index, value) {
                            var cateTitle = value.Category;
                            var str = "<li>" +
                                "<a href='#' class='liTitle' attr-title='" + cateTitle + "' attr-category='" + value.ID + "'>" +
                                "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                                "<span class='fileName' title='" + cateTitle + "'>" + (cateTitle.length > 20 ? cateTitle.substr(0, 20) + "..." : cateTitle) + "</span>" +
                                "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                                "</a >" +
                                "</li>"
                            $_ul.append(str);
                        });
                        $("#WaitDialog").hide();
                        $('#emptydataTwo').hide();
                    }
                    callback(1);
                });
            }
        }).then(function (result) {
            var $_ul = $("#two ul.childCategoryItems");
            $_ul.html("");
            if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
                loadDataFromSharePoint(DocsNodePictureDisplayName, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory', "(ImageCategory/ID eq " + $_category + ")", "image").then(function (data) {
                    results = data.results;
                    if (results.length > 0) {
                        imageMaxID = _.max(_.pluck(results, "ID"));
                        $.each(results, function (index, image) {
                            toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                                var base64result = "data:image/png;base64, " + dataUrl;
                                $_ul.append("<li class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' src='" + base64result + "'>" +
                                    "</div></a></li>");
                                $("#WaitDialog").hide();
                            });
                        });
                        $('#emptydataTwo').hide();
                    } else {
                        if ($_ulList[0].childElementCount == 0 && $_ulItems[0].childElementCount == 0) {
                            $('#emptydataTwo').show();
                            $("#WaitDialog").hide();
                        } else {
                            $('#emptydataTwo').hide();
                        }
                    }
                });
            }
        });
    } catch (error) {
        console.log('loadImagesByCategory: ' + error);
    }

}

function LoadCategories(category_name, columns, filters) {
    var dfd = $.Deferred();

    var url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + ")";

    try {
        callAjaxGet(url).done(function (data) {
            dfd.resolve(data.d);
        });
    } catch (err) {
        dfd.reject(err);
    }
    return dfd;
}

function bindSlidesHTML(data, itemContainer) {
    $.each(data, function (index, slides) {
        var filename = slides.LinkFilename;
        var title = slides['Title'];
        var insertSlideHtml = 'insertSlide(\'' + slides.EncodedAbsUrl + '\')';
        itemContainer.append("<li title='" + title + "' onclick=" + insertSlideHtml + " class='liIteamImg'><a href='#' class='liInnerImage'><label class='lblId' style='display:none'>" + slides.Id +
            "</label><span class='SubSicon'></span><span class='SubScontent'><img src='" + SPURL + templateServerRelURL + "_layouts/15/getpreview.ashx?path=" + slides.EncodedAbsUrl + "&resolution=0" + "'/></span></a></li>");
    });
    $("#WaitDialog").hide();
    $('#emptydataOne').hide();
}

function insertSlide(url) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
        function (asyncResult) {
            var currentSlideIndex = asyncResult.value.slides[0].index;
            var presentation = new ActiveXObject("powerpoint.Application"); // Used to store a new presentation for each slide.
            presentation.ActivePresentation.Slides.InsertFromFile(url, currentSlideIndex, 1, 1);
        }
    );
}

//function toDataURL(fileURL, callback) {
//    try {
//        fileURL = fileURL.replace(SPURL, "");
//        var url = SPURL + imageListSite + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
//        var xhr = new window.XMLHttpRequest();
//        xhr.open("GET", url, true);
//        xhr.setRequestHeader("Accept", "application/json; odata=verbose");
//        xhr.setRequestHeader("Authorization", "Bearer " + SPToken);
//        //Now set response type
//        xhr.responseType = 'arraybuffer';
//        xhr.addEventListener('load', function () {
//            if (xhr.status === 200) {
//                filereader().then(function (imagedata) {
//                    var dataUrl = imagedata;
//                    var base64 = dataUrl.split(',')[1];
//                    callback(base64);
//                });
//            }
//            function filereader() {
//                var dfdImg = $.Deferred();
//                var sampleBytes = new Uint8Array(xhr.response);
//                var blob = new Blob([sampleBytes], { type: "image/jpeg" });
//                var reader = new FileReader();
//                reader.onload = function () {
//                    dfdImg.resolve(reader.result);
//                }
//                reader.readAsDataURL(blob);
//                return dfdImg.promise();
//            }
//        });
//        xhr.send();
//    } catch (error) {
//        console.log('toDataUrl: ' + error);
//    }
//}

function toDataURL(fileURL, callback, folderResetCount) {
    fileURL = fileURL.replace(SPURL, "");
    //var url = SPURL + imageListSite + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
    var url = SPURL + imageListSite + "/_api/web/GetFileByServerRelativeUrl('" + fileURL + "')";
    var payload = {
        "SPOUrl": ORG_ROOT_WEB.webUrl,
        "tenant": ORG_TENANT.id,
        "ImagePath": url
    };


    $.ajax({
        url: GET_TEMPLATEIMAGE_URL,
        beforeSend: function (request) {
            request.setRequestHeader("Accept", "application/json; odata=verbose");
            request.setRequestHeader('cache-control', 'max-age=3600');
        },
        dataType: "json",
        headers: {
            'Authorization': 'Bearer ' + SPToken,
        },
        type: "POST",
        contentType: "application/json",
        data: JSON.stringify(payload),
    }).done(function (response) {
        if (folderResetCount === IMGFolderReset) {
            callback(response);
        }
    });


    //var xhr = new window.XMLHttpRequest();
    //xhr.open("GET", url, true);
    //xhr.setRequestHeader("Accept", "application/json; odata=verbose");
    //xhr.setRequestHeader("Authorization", "Bearer " + SPToken);
    ////Now set response type
    //xhr.responseType = 'arraybuffer';
    //xhr.addEventListener('load', function () {
    //    if (xhr.status === 200) {
    //        var sampleBytes = new Uint8Array(xhr.response);
    //        var blob = new Blob([sampleBytes], { type: "image/jpeg" });
    //        var reader = new FileReader();
    //        reader.onload = function () {
    //            var dataUrl = reader.result;
    //            var base64 = dataUrl.split(',')[1];
    //            callback(base64);
    //        };
    //        reader.readAsDataURL(blob);
    //    }
    //})
    //xhr.send();
}

function loadDataFromSharePoint(category_name, columns, filters, type, search) {
    var dfd = $.Deferred();
    var url = "";

    // type is not undefined means lazy loading should be done.
    if (type != undefined) {
        if (type == "image") {
            if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
                //columns = 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory';
                //filters = '(ImageCategory eq null)';
                if (imageListSite.toLowerCase() == "/sites/docsnodeadmin") {
                    imageListSite = imageListSite + "/";
                }
            }
            else {
                columns = 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl';
                if (!search)
                    filters = '';
            }

            if ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == null || $("#hdnCategory").val() == undefined)// It means parent level
                url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + (filters ? (filters + " and") : '') + " ID gt " + imageParentMaxID + ")&$top=" + imageItem;
            else
                url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + (filters ? (filters + " and") : '') + " ID gt " + imageMaxID + ")&$top=" + imageItem;
        }
        else if (type == "slide") {
            if ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == null || $("#hdnTextCategory").val() == undefined)
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + " and ID gt " + slideParentMaxID + ")&$top=" + slideItem;
            else
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + " and ID gt " + slideMaxID + ")&$top" + slideItem;
        }
    }
    else {
        if (imageListSite.toLowerCase() != '/sites/docsnodeadmin/' && imageListSite.toLowerCase() != '/sites/docsnodeadmin') {
            url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns;
        }
        else {
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + ")";
        }
    }

    try {
        callAjaxGet(url).then(function (data) {
            dfd.resolve(data.d);
        }, function (error) {
            dfd.reject(error);
        });
    } catch (err) {
        dfd.reject(err);
    }

    return dfd;
}

function callAjaxGet(url) {
    var dfdGET = $.Deferred();
    try {
        if (SPToken) {
            $.ajax({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                success: function (data) {
                    dfdGET.resolve(data);
                },
                error: function (data) {
                    console.log(data);
                    dfdGET.reject(data);
                }
            });
        }
    } catch (error) {
        console.log("callAjaxGet: " + error);
        dfdGET.reject(error);
    }
    return dfdGET.promise();
}


function insertImage(fileURL) {
    onImageClickToDataURL(fileURL, function (base64result) {
        Office.context.document.setSelectedDataAsync(base64result, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log(asyncResult.error.message);
                }
            });
    });
}

function onImageClickToDataURL(fileURL, callback) {
    fileURL = fileURL.replace(SPURL, "");
    //var url = SPURL + imageListSite + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
    var url = SPURL + imageListSite + "/_api/web/GetFileByServerRelativeUrl('" + fileURL + "')";
    var payload = {
        "SPOUrl": ORG_ROOT_WEB.webUrl,
        "tenant": ORG_TENANT.id,
        "ImagePath": url
    };


    $.ajax({
        url: GET_TEMPLATEIMAGE_URL,
        beforeSend: function (request) {
            request.setRequestHeader("Accept", "application/json; odata=verbose");
            request.setRequestHeader('cache-control', 'max-age=3600');
        },
        dataType: "json",
        headers: {
            'Authorization': 'Bearer ' + SPToken,
        },
        type: "POST",
        contentType: "application/json",
        data: JSON.stringify(payload),
    }).done(function (response) {

        callback(response);

    });

}
