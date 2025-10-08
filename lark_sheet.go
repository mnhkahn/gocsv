package gocsv

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"reflect"
	"strings"

	lark "github.com/larksuite/oapi-sdk-go/v3"
	larkcore "github.com/larksuite/oapi-sdk-go/v3/core"
	larksheets "github.com/larksuite/oapi-sdk-go/v3/service/sheets/v3"
)

// https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet/query?appId=cli_9f806e1fcb60d00e
func LarkSheet(ctx context.Context, appID, appToken, url string, out interface{}) error {
	token, sheet, err := GetSheetTokenByUrl(ctx, appID, appToken, url)
	if err != nil {
		return err
	}

	// 创建 Client
	res, err := readSheetsContent(ctx, appID, appToken, token, sheet)
	// 处理错误
	if err != nil {
		return err
	}

	outValue, outType := getConcreteReflectValueAndType(out)                   // Get the concrete type (not pointer) (Slice<?> or Array<?>)
	outInnerWasPointer, outInnerType := getConcreteContainerInnerType(outType) // Get the concrete inner type (not pointer) (Container<"?">)

	headers := make([]string, 0, len(res.ValueRange.Values[0]))
	tagMap := make(map[string]int, len(res.ValueRange.Values[0]))
	for i := 0; i < outInnerType.NumField(); i++ {
		tag := outInnerType.Field(i).Tag.Get("csv")
		if tag == "" || tag == "-" {
			continue
		}
		tagMap[tag] = i
	}
	for i, row := range res.ValueRange.Values {
		if i == 0 {
			for _, v := range row {
				headers = append(headers, v.(string))
			}
		} else {
			outInner := createNewOutInner(outInnerWasPointer, outInnerType)
			for j, columnContent := range row {
				dstV := reflect.ValueOf(columnContent)
				header := headers[j]
				tagIndex, ok := tagMap[header]
				if !ok {
					continue
				}

				dst := outInnerType.Field(tagIndex).Type
				field := outInner.Elem().Field(tagIndex)
				defer func() {
					if r := recover(); r != nil {
						fmt.Println("Recovered from panic:", r, header, columnContent, tagIndex)
					}
				}()

				if dst.Kind() == reflect.String && dstV.Kind() != reflect.String {
					// println(tagIndex, header, columnContent, dst.Name())
					a := getStringValue(columnContent)
					field.Set(reflect.ValueOf(a).Convert(dst))
				} else {
					// println(tagIndex, header, columnContent, dst.Name(), "not string")
					field.Set(dstV.Convert(dst))
				}
			}
			outValue.Set(reflect.Append(outValue, outInner))
		}
	}

	return nil
}

func readSheetsContent(ctx context.Context, appID, appToken, spreadsheetToken string, valueRange string) (*ReadSheetsContent, error) {
	client := lark.NewClient(appID, appToken)
	url := fmt.Sprintf("https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/%s/values/%s", spreadsheetToken, valueRange)
	resp, err := client.Get(ctx, url, nil, larkcore.AccessTokenTypeTenant)
	if err != nil {
		return nil, err
	} else if resp.StatusCode != http.StatusOK {
		return nil, errors.New(string(resp.RawBody))
	}
	res := new(ReadSheetsContentResponse)
	err = json.Unmarshal(resp.RawBody, res)
	if err != nil {
		return nil, err
	}
	return res.Data, nil
}

func GetSheetTokenByUrl(ctx context.Context, appID, appToken, larkUrl string) (string, string, error) {
	if larkUrl == "" {
		return "", "", errors.New("文档链接为空")
	}
	u, err := url.Parse(larkUrl)
	if err != nil {
		return "", "", err
	}
	path := u.Path
	if path == "" {
		return "", "", errors.New("文档链接异常")
	}
	split := strings.Split(path, "/")
	if len(split) < 3 || split[1] != "sheets" || split[2] == "" {
		return "", "", errors.New("文档token获取失败")
	}
	sheet := u.Query().Get("sheet")
	if sheet == "" {
		res, err := GetSheet(ctx, appID, appToken, split[2])
		if err != nil {
			return split[2], "", err
		} else {
			return split[2], res[0], nil
		}
	}
	return split[2], sheet, nil
}

type ReadSheetsContentResponse struct {
	Code int                `json:"code"`
	Msg  string             `json:"msg"`
	Data *ReadSheetsContent `json:"data"`
}

type ReadSheetsContent struct {
	Revision         int    `json:"revision"`
	SpreadsheetToken string `json:"spreadsheetToken"`
	ValueRange       *struct {
		MajorDimension string          `json:"majorDimension"`
		Range          string          `json:"range"`
		Revision       int             `json:"revision"`
		Values         [][]interface{} `json:"values"`
	} `json:"valueRange"`
}

// https://open.larkoffice.com/document/server-docs/docs/sheets-v3/spreadsheet-sheet/query?appId=cli_a626110429dd100e
func GetSheet(ctx context.Context, appID, appToken, token string) ([]string, error) {
	// 创建 Client
	client := lark.NewClient(appID, appToken)
	// 创建请求对象
	req := larksheets.NewQuerySpreadsheetSheetReqBuilder().
		SpreadsheetToken(token).
		Build()

	// 发起请求
	resp, err := client.Sheets.SpreadsheetSheet.Query(context.Background(), req)
	if err != nil {
		return nil, err
	} else if resp.CodeError.Code != 0 {
		return nil, errors.New(resp.CodeError.Msg)
	}
	res := make([]string, 0)
	for _, sheet := range resp.Data.Sheets {
		res = append(res, *sheet.SheetId)
	}
	return res, nil
}
