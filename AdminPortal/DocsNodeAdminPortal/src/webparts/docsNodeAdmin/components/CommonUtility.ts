
export default class CommonUtility {

    //Get Default Site Collection Path
    public siteCollectionPath = '/sites/contenttypehub';

    //Get Default Tenant Name from window url
    public tenantURL() {
        try {
            let currentPageURL = window.location.href;
            let tenantURL = currentPageURL.split('com/')[0];
            tenantURL = tenantURL + 'com';
            return tenantURL;
        } catch (error) {
            console.log("tenantURL: " + error);
        }
    }

    //This function to get data from List or Library on 'GET' request
    public _getRequest(url: string): any {
        const one = new Promise<any>((resolve, reject) => {
            try {
                return fetch(url, {
                    headers: { Accept: 'application/json;odata=verbose' },
                    credentials: "same-origin"
                }).then(function (response) {
                    if (response.status >= 200 && response.status < 400) {
                        resolve(response.json());
                    }
                    else {
                        reject("createBoard: ");
                    }
                }).catch(error =>
                    console.error("getRequest: " + error));
            } catch (error) {
                console.log("getRequest: " + error);
            }
        });
        return one;
    }

    //This function to post data in List or Library on 'POST' request
    public _postRequest(url: string, postBody, xMethod): any {
       // const one = new Promise<any>((resolve, reject) => {
            try {
                return this.getValues().then((token) => {
                    return fetch(url, {
                        headers: {
                            Accept: 'application/json;odata=verbose',
                            "Content-Type": 'application/json;odata=verbose',
                            "X-RequestDigest": token.d.GetContextWebInformation.FormDigestValue,
                            "X-Http-Method": xMethod,
                            'IF-MATCH': '*'                            
                        },
                        method: 'POST',
                        body: postBody,
                        credentials: "same-origin"
                    }).then((response) => {
                        if(response.status <= 204 && response.status >= 200){
                            return 'success';
                        }
                        else{
                            return response.json();
                        }                        
                    }).then((res)=>{
                        console.log(res);
                        return res;
                    }).catch(error => 
                        console.error("postRequest: " + error));
                }, (error) => {
                    console.log("getValues => postRequest" + error);
                });
            } catch (error) {
                console.log("postRequest: " + error);
            }
        //});
       // return one;
    }

    //This function is to get FormDigestValue for X-RequestDigest
    public getValues(): any {
        try {
            var url = this.tenantURL() + this.siteCollectionPath + "/_api/contextinfo";
            return fetch(url, {
                method: "POST",
                headers: { Accept: "application/json;odata=verbose" },
                credentials: "same-origin"
            })
                .then((response) => {
                    return response.json();
                });
        } catch (error) {
            console.log("getValues: " + error);
        }
    }
}