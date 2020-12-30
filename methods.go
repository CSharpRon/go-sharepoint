package sharepoint

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"net/http"
	"strings"
)

// get returns the body from a GET request response on the SharePoint api
func get(c Connection, endpoint string) ([]byte, int, error) {

	accessToken, err := c.accessToken()

	// Execute the request
	request, err := http.NewRequest("GET", endpoint, strings.NewReader(""))
	if err != nil {
		return nil, 0, err // An error indicates that the endpoint is incorect
	}

	// Needed headers
	request.Header.Add("Authorization", fmt.Sprintf("Bearer %s", accessToken))
	request.Header.Add("Accept", "application/json;odata=nometadata")
	request.Header.Add("Host", c.DomainHost)

	client := &http.Client{}
	response, err := client.Do(request)
	if err != nil {
		return nil, 0, err
	}

	body, err := ioutil.ReadAll(response.Body)
	defer response.Body.Close()

	return body, response.StatusCode, nil
}

// patch returns the body from a PATCH response on the SharePoint API
func patch(c Connection, endpoint string, form []byte) (*http.Response, error) {

	if c.DisableMutations {
		return &http.Response{StatusCode: http.StatusNoContent}, nil
	}

	accessToken, err := c.accessToken()

	// Execute the request
	request, err := http.NewRequest("PATCH", endpoint, bytes.NewBuffer(form))
	if err != nil {
		return nil, err // An error indicates that the endpoint is incorect
	}

	// Needed headers
	request.Header.Add("Authorization", fmt.Sprintf("Bearer %s", accessToken))
	request.Header.Add("Accept", "application/json")
	request.Header.Add("Host", c.DomainHost)
	request.Header.Add("If-Match", "*")
	request.Header.Add("X-HTTP-Method", "MERGE")
	request.Header.Add("Content-Type", "application/json;odata=nometadata")
	request.Header.Add("Content-Length", fmt.Sprintf("%d", (request.ContentLength+1)))

	client := &http.Client{}
	response, err := client.Do(request)
	if err != nil {
		return nil, err
	}

	return response, nil

}

// post returns the body from a POST request response on the SharePoint API
func post(c Connection, endpoint string, form []byte) (*http.Response, error) {

	if c.DisableMutations {
		return &http.Response{StatusCode: http.StatusCreated}, nil
	}

	accessToken, err := c.accessToken()

	// Build the request
	request, err := http.NewRequest("POST", endpoint, bytes.NewBuffer(form))
	if err != nil {
		return nil, err // An error indicates that the endpoint is incorect
	}

	// Needed headers
	request.Header.Add("Authorization", fmt.Sprintf("Bearer %s", accessToken))
	request.Header.Add("Content-Type", "application/json;odata=nometadata")
	request.Header.Add("Accept", "application/json")
	request.Header.Add("Host", c.DomainHost)
	request.Header.Add("Content-Length", fmt.Sprintf("%d", (request.ContentLength+1)))

	client := &http.Client{}
	response, err := client.Do(request)
	if err != nil {
		return nil, err
	}

	// A post should only ever return one of these two status codes
	if response.StatusCode == http.StatusOK || response.StatusCode == http.StatusCreated {
		return response, nil
	}

	return response, errStatusCode

}
