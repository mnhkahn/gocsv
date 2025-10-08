package gocsv

import (
	"context"
	"testing"
)

func TestLarkSheet(t *testing.T) {
	res := make([]*a, 0, 1)
	err := LarkSheet(context.Background(), "your_app_id", "your_app_token", "https://fntxar02de.sg.larksuite.com/sheets/El7rsO6a3h15WCtiP8SlcxiogYd", &res)
	if err != nil {
		t.Error(err.Error())
	}
	for _, a := range res {
		t.Log(a.Word)
	}
}

type a struct {
	Word string `csv:"字"`
}
