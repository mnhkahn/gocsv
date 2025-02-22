package gocsv

import (
	"context"
	"testing"
)

func TestLarkSheet(t *testing.T) {
	res := make([]*a, 0, 1)
	err := LarkSheet(context.Background(), "cli_a733fba967381003", "amMVjbLCLoFvReAoCEtucfVKbxdlwqdk", "https://fntxar02de.larksuite.com/sheets/HATpsm58GhyuxPt3NVCuEHygs2g?sheet=KKOdcC", &res)
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
