package sharepoint

// RawSharePointResponse represents the json data that is returned from a HTTP request
type RawSharePointResponse struct {
        Value []interface{} `json:"value" sharepoint:"value"`
}
