package main

type Operation struct {
	Type     string            `json:"type"`
	Sheet    string            `json:"sheet"`
	Column   string            `json:"column,omitempty"`
	Count    int               `json:"count,omitempty"`
	Mappings map[string]string `json:"mappings,omitempty"`
}
