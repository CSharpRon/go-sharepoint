// Package sharepoint contains a simple interface for retrieving and updating list items from SharePoint Online.
//
// Usage:
// Each public function from this package requires that an Authentication object be passed in. This object must have all of the public fields instantiated.
//
// Example:
/* con := sharepoint.Connection {
        ClientID: "SuperSecretValue",
        TenantID: "0000000-0000000-000000-00001",
        ClientSecret: "SuperSecretValue",
        RefreshToken: "SuperSecretRefreshToken",
        URLHost: "https://contoso.sharepoint.com/teams/marketing",
        DomainHost: "contoso.sharepoint.com",
}
*/
package sharepoint

import (
        "encoding/json"
        "fmt"
        "io/ioutil"
        "net/http"
        "reflect"
        "strconv"
        "strings"
        "time"

        "github.com/fatih/structs"
        "github.com/mitchellh/mapstructure"
)

var (
        errNil        = fmt.Errorf("Unable to map response body to output: variable is nil")
        errMarshal    = fmt.Errorf("Unable to marshal the input to a JSON object")
        errStatusCode = fmt.Errorf("Unexpected status code returned from the SharePoint service")
        errUnmarshal  = fmt.Errorf("Unable to cast SharePoint response body to raw struct")
        errID         = fmt.Errorf("Missing or invalid ID passed in")
        errNotJSON    = fmt.Errorf("Item is not a valid json object")
)

// Connection holds the configuration settings for the SharePoint online connection.
type Connection struct {

        // ClientID holds the client id for the SharePoint connection
        ClientID string `json:"-"`

        // TenantID holds the tenant (or realm) id for the SharePoint connection
        TenantID string `json:"-"`

        // ClientSecret holds the client secret for the SharePoint connection
        ClientSecret string `json:"-"`

        // RefreshToken holds the refresh token for the SharePoint connection
        RefreshToken string `json:"-"`

        // URLHost holds the fqdn to the specific SharePoint group such as constoso.sharepoint.com/teams/marketing
        URLHost string `json:"-"`

        // DomainHost holds the fqdn to the SharePoint root URL such as constoso.sharepoint.com
        DomainHost string `json:"-"`

        securityToken Token
}

// Token holds onto the access information returned by the SharePoint Online credentialing service.
//
// It should not be instantiated by the user. Instead, it is created by this package once a successful token has been retrieved.
type Token struct {
        TokenType string `json:"token_type"`
        ExpiresIn string `json:"expires_in"`
        NotBefore string `json:"not_before"` // Unix Timestamp
        ExpiresOn string `json:"expires_on"` // Unix Timestamp
        Resource  string `json:"resource"`
        Token     string `json:"access_token"`
}

// Test attempts to retrieve an access token for the SharePoint connection configuration. If it succeeds, nil is returned. Otherwise, the error is returned.
func (c *Connection) Test() error {

        _, err := c.accessToken()

        return err
}

// GetListItems will create the appropriate GET request to SharePoint online from the Authentication, listName, and queryString object
// and map the result onto a new interface.
func (c *Connection) GetListItems(listName string, queryString string) (RawSharePointResponse, error) {

        // Standard endpoint format to get items from an SP list
        endpoint := fmt.Sprintf("%s/_api/web/lists/getbytitle('%s')/items?%s", c.URLHost, listName, queryString)

        return getItems(*c, endpoint)
}

// GetListItemByID populates the output interface with values returned from a sharepoint query where the itemID is specified.
//
// If the itemID is blank, all items from the list will be returned.
func (c *Connection) GetListItemByID(listName string, itemID string, queryString string) (RawSharePointResponse, error) {

        // Standard endpoint format to get items from an SP list
        endpoint := fmt.Sprintf("%s/_api/web/lists/getbytitle('%s')/items(%s)?%s", c.URLHost, listName, itemID, queryString)

        return getItems(*c, endpoint)
}

// InsertListItem adds a new Marshalled JSON object to the specified Sharepoint list.
func (c *Connection) InsertListItem(listName string, item interface{}, fields ...string) error {

        // Standard endpoint format to insert items into a list
        endpoint := fmt.Sprintf("%s/_api/web/lists/getbytitle('%s')/items", c.URLHost, listName)

        // The item passed in should have tags marked "sharepoint" to denote the names of the SharePoint fields that correspond to the interface values
        structs.DefaultTagName = "sharepoint"

        return insertListItem(*c, item, endpoint, fields...)

}

// UpdateListItem updates an item where a valid sharepoint id is passed in.
// The item variable passed in will be converted into a JSON object that will be sent in the PATCH request.
//
// By default, all of the "sharepoint" tags on the interface will be passed into the PATCH request.
// However, this can be overridden by explicitly specifying them in the fields... variable.
// In this way, the http response body is dynamically created
func (c *Connection) UpdateListItem(listName string, item interface{}, id int, fields ...string) error {

        if id == 0 {
                return errID
        }

        // Standard endpoint format to insert items into a list
        endpoint := fmt.Sprintf("%s/_api/web/lists/getbytitle('%s')/items('%d')", c.URLHost, listName, id)

        return updateListItem(*c, listName, item, endpoint, id, fields...)

}

// GetDocumentLibraryItems will go out the SharePoint site and return an array of items from the specified library. When the list returns empty, an emtpy array will be returned.
func (c *Connection) GetDocumentLibraryItems(folderRelativePath string) (RawSharePointResponse, error) {

        endpoint := fmt.Sprintf("%s/_api/web/GetFolderByServerRelativeUrl('%s')/files", c.URLHost, folderRelativePath)

        return getItems(*c, endpoint)
}

// DownloadDocumentLibraryFile returns the byte array pertaining to a file in the passed in document library. If the file is not available, an error is returned.
func (c *Connection) DownloadDocumentLibraryFile(folderRelativePath string, fileName string) ([]byte, error) {

        endpoint := fmt.Sprintf("%s/_api/web/GetFolderByServerRelativeUrl('%s')/files('%s')/$value", c.URLHost, folderRelativePath, fileName)

        return download(*c, endpoint)
}

// UploadDocumentLibraryFile performs a POST request to upload the specified file.
func (c *Connection) UploadDocumentLibraryFile(folderRelativePath string, fileName string, overwriteOnConflict bool, file []byte) error {

        overwriteFlag := "true"
        if !overwriteOnConflict {
                overwriteFlag = "false"
        }

        endpoint := fmt.Sprintf("%s/_api/web/GetFolderByServerRelativeUrl('%s')/files/add(url='%s',overwrite=%s)", c.URLHost, folderRelativePath, fileName, overwriteFlag)

        _, err := post(*c, endpoint, file)
        return err
}

// ToTimeHookFunc is a custom decoder for mapstructure which analyzes if the value being parsed can be converted to a time object
//
// If it can, then the conversion is made. If not, the default value is used.
func ToTimeHookFunc() mapstructure.DecodeHookFunc {
        return func(
                f reflect.Type,
                t reflect.Type,
                data interface{}) (interface{}, error) {
                if t != reflect.TypeOf(time.Time{}) {
                        return data, nil
                }

                switch f.Kind() {
                case reflect.String:
                        return time.Parse(time.RFC3339, data.(string))
                case reflect.Float64:
                        return time.Unix(0, int64(data.(float64))*int64(time.Millisecond)), nil
                case reflect.Int64:
                        return time.Unix(0, data.(int64)*int64(time.Millisecond)), nil
                default:
                        return data, nil
                }
                // Convert it by parsing
        }
}

// accessToken will instantiate or return the access token pertaining to the SharePoint connection.
// If no access token is returned, an error is returned.
func (c *Connection) accessToken() (string, error) {

        var expiration int64
        expiration, _ = strconv.ParseInt(c.securityToken.ExpiresOn, 10, 64)

        // If the token is empty or expired, generate it
        if c.securityToken.Token == "" || expiration == 0 || expiration <= time.Now().Unix() {

                var response Token

                // Azure AAC Endpoint
                url := "https://accounts.accesscontrol.windows.net/tokens/OAuth/2"

                payloadString := "grant_type=refresh_token" +
                        "&client_id=" + c.ClientID + "@" + c.TenantID +
                        "&client_secret=" + c.ClientSecret +
                        "&resource=00000003-0000-0ff1-ce00-000000000000/" + c.DomainHost + "@" + c.TenantID +
                        "&refresh_token=" + c.RefreshToken

                payload := strings.NewReader(payloadString)

                client := &http.Client{}
                req, err := http.NewRequest("GET", url, payload)

                if err != nil {
                        return "", err
                }

                req.Header.Add("Content-Type", "application/x-www-form-urlencoded")
                res, err := client.Do(req)
                if err != nil {
                        return "", err
                }

                if res.StatusCode != 200 {
                        return "", fmt.Errorf("Bad Request in retrieving access token: %s", res.Body)
                }

                body, err := ioutil.ReadAll(res.Body)
                defer res.Body.Close()

                json.Unmarshal(body, &response)
                c.securityToken = response
        }

        // Return just the token
        return c.securityToken.Token, nil
}

// baseType is a helper function to return the base type of the passed in variable t, which is useful if the type is an array.
func baseType(t reflect.Type, expected reflect.Kind) (reflect.Type, error) {

        base := t.Elem()
        if t.Elem().Kind() == reflect.Ptr {
                base = t.Elem().Elem()
        }

        if base.Kind() != expected {
                return nil, fmt.Errorf("expected %s but got %s", expected, t.Kind())
        }

        return base, nil
}

func download(c Connection, endpoint string) ([]byte, error) {

        // Execute the request and get the response body
        body, httpStatus, err := get(c, endpoint)
        if err != nil {
                return nil, err
        }

        // If the status code returned is not what was expected, throw an error
        if httpStatus != http.StatusOK {
                return nil, errStatusCode
        }

        // If a file is being downloaded, then the response body *is* the file contents
        return body, nil

}

// retreive the specified items from the SharePoint list and return them in a RawSharePointResponse object
func getItems(c Connection, endpoint string) (RawSharePointResponse, error) {

        // Read in the response to the raw struct container
        var rawSPResponse RawSharePointResponse

        // Execute the request and get the response body
        body, httpStatus, err := get(c, endpoint)
        if err != nil {
                return rawSPResponse, err
        }

        // If 404 returned, no items were found and we can return early
        if httpStatus == http.StatusNotFound {
                return rawSPResponse, nil
        }

        // The response should be a standard response body from SharePoint, where the data is contained in the Value field
        err = json.Unmarshal(body, &rawSPResponse)
        if err != nil {
                return rawSPResponse, errUnmarshal
        }

        // If the value field is nil, it is because a single item response was returned but was not decoded
        if rawSPResponse.Value == nil {

                var singleItem interface{}
                var values []interface{}

                json.Unmarshal(body, &singleItem)
                if err != nil {
                        return rawSPResponse, errUnmarshal
                }
                rawSPResponse = RawSharePointResponse{
                        Value: append(values, singleItem),
                }
        }

        return rawSPResponse, nil
}

func insertListItem(c Connection, item interface{}, endpoint string, fields ...string) error {

        var byteValue json.RawMessage
        var cleanSheet []byte
        var err error
        sheetMap := structs.Map(item)
        validMap := make(map[string]json.RawMessage)

        // If fields were passed in, apply the filter
        if len(fields) > 0 {

                // Append the corresponding value to the new map object
                for _, field := range fields {
                        byteValue, _ = json.Marshal(sheetMap[field])
                        validMap[field] = byteValue
                }

                cleanSheet, err = json.Marshal(validMap)
                if err != nil {
                        return errMarshal
                }

        } else {
                // Produce a cleaned up version of the object that was passed in
                for key, value := range sheetMap {

                        byteValue, _ = json.Marshal(value)
                        stringValue := string(byteValue)

                        // Ignore values that SharePoint will automatically create
                        if key == "Created" || key == "Modified" || key == "GUID" || key == "Attachments" {
                                continue
                        }

                        // Ignore IDs when they are set to 0
                        if strings.Contains(strings.ToLower(key), "id") && stringValue == "0" {
                                continue
                        }

                        validMap[key] = byteValue
                }

                cleanSheet, err = json.Marshal(validMap)
                if err != nil {
                        return errMarshal
                }

        }

        response, err := post(c, endpoint, cleanSheet)
        if err != nil {
                return err
        }

        // Only status code 201 indicates that the item was created
        if response.StatusCode != 201 {
                return errStatusCode
        }

        return nil
}

// ScanResponse converts a RawSharePointResponse (assuming it has data) into the data type of the output variable.
// This variable must be a pointer and must be an array since the type of RawSharePointResponse.Value is an array.
func (c *Connection) ScanResponse(rawSPResponse RawSharePointResponse, output interface{}) error {

        return scanResponse(rawSPResponse, output)
}

// scanResponse converts a RawSharePointResponse (assuming it has data) into the data type of the output variable.
// This variable must be a pointer and must be an array since the type of RawSharePointResponse.Value is an array.
func scanResponse(rawSPResponse RawSharePointResponse, output interface{}) error {

        // With a valid response, we can use reflection to map it to the output interface
        // Logic very heavily inspired from the encoding implementation of the github.com/jmoiron/sqlx package
        var vp reflect.Value

        // Get a pointer to the output
        value := reflect.ValueOf(output)

        // Error Checks
        if value.Kind() != reflect.Ptr {
                return fmt.Errorf("Must pass a pointer, not a value, to StructScan destination")
        }

        if value.IsNil() {
                return errNil
        }

        // Get the object that the pointer is holding
        direct := reflect.Indirect(value)
        slice, err := baseType(value.Type(), reflect.Slice)
        if err != nil {
                return fmt.Errorf("Unable to get slice of pointer object")
        }
        isPtr := slice.Elem().Kind() == reflect.Ptr

        // Get the base type for an element in the output (derefrences pointer if necessary)
        base := slice.Elem()
        if slice.Elem().Kind() == reflect.Ptr {
                base = slice.Elem().Elem()
        }

        for _, r := range rawSPResponse.Value {

                // A new instance of the output's base type
                vp = reflect.New(base)

                // Converts a single object in the response to a map
                responseMap := make(map[string]interface{})
                mapDecoder := mapstructure.DecoderConfig{
                        TagName:    "sharepoint",
                        Result:     vp.Interface(),
                        DecodeHook: mapstructure.ComposeDecodeHookFunc(ToTimeHookFunc())}

                // Marshal the object back into bytes and then convert it into a map
                singleItem, _ := json.Marshal(r)
                json.Unmarshal(singleItem, &responseMap)

                // Using the custom decoder which looks for the sharepoint struct tag, decode the valid to the vp object
                decoder, _ := mapstructure.NewDecoder(&mapDecoder)
                err = decoder.Decode(responseMap)

                // Now that we have a decoded object, append it to the pointer object for the output
                if isPtr {
                        direct.Set(reflect.Append(direct, vp))
                } else {
                        direct.Set(reflect.Append(direct, reflect.Indirect(vp)))
                }
        }

        return nil
}

func updateListItem(c Connection, list string, item interface{}, endpoint string, sharepointid int, fields ...string) error {

        var byteValue json.RawMessage
        var jsonBody []byte
        var err error

        structs.DefaultTagName = "sharepoint"

        sheetMap := structs.Map(item)
        validMap := make(map[string]json.RawMessage)

        // If fields were passed in, apply the filter
        if len(fields) > 0 {

                for _, field := range fields {
                        byteValue, _ = json.Marshal(sheetMap[field])
                        validMap[field] = byteValue
                }
                jsonBody, err = json.Marshal(validMap)
                if err != nil {
                        return errMarshal
                }
        } else {

                // Produce a cleaned up version of the object that was passed in
                for key, value := range sheetMap {

                        byteValue, _ = json.Marshal(value)
                        stringValue := string(byteValue)

                        // Ignore values that SharePoint will automatically create
                        if key == "Created" || key == "Modified" || key == "GUID" || key == "Attachments" {
                                continue
                        }

                        // Ignore IDs when they are set to 0
                        if strings.Contains(strings.ToLower(key), "id") && stringValue == "0" {
                                continue
                        }

                        validMap[key] = byteValue
                }

                jsonBody, err = json.Marshal(validMap)
                if err != nil {
                        return errMarshal
                }
        }

        response, err := patch(c, endpoint, jsonBody)
        if err != nil {
                return err
        }

        // Only status code 204 indicates success
        if response.StatusCode != 204 {
                return errStatusCode
        }

        return nil
}
