package main

import (
	"context"
	"errors"
	"flag"
	"fmt"
	"log"
	"os"

	"github.com/xuri/excelize/v2"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

func main() {
	upload := flag.String("upload", "", "Usage: go run main.go --upload <file.xlsx>")

	if *upload != "" {
		err := UploadToSpreadSheet(*upload)
		if err != nil {
			fmt.Printf("failed to upload file: %v", err)
		}
		os.Exit(1)
	}
	if *upload == "" {
		fmt.Println("Error: Missing required flags.")
		flag.PrintDefaults()
		os.Exit(1)
	}
}

func UploadToSpreadSheet(excel string) error {
	// Dont forget to add the SpreadSheet ID
	spreadsheetID := os.Getenv("SPS_ID")
	if spreadsheetID == "" {
		return errors.New("missing spreadsheet id")
	}

	ctx := context.Background()

	srv, err := sheets.NewService(ctx, option.WithCredentialsFile("credentials.json"))
	if err != nil {
		return err
	}

	// Open the Excel file
	file, err := excelize.OpenFile(excel)
	if err != nil {
		return err
	}

	// Sheets to process
	sheetNames := file.GetSheetList()

	for _, sheetName := range sheetNames {
		rows, err := file.GetRows(sheetName)
		if err != nil {
			log.Printf("Sheet '%s' not found in the Excel file. Skipping.\n", sheetName)
			continue
		}

		// Prepare data for upload
		var data [][]interface{}
		for _, row := range rows {
			var rowData []interface{}
			for _, cell := range row {
				rowData = append(rowData, cell)
			}
			data = append(data, rowData)
		}

		// Check if the Google Sheet contains this sheet, or create it
		spreadsheet, err := srv.Spreadsheets.Get(spreadsheetID).Do()
		if err != nil {
			return err
		}

		sheetExists := false
		for _, s := range spreadsheet.Sheets {
			if s.Properties.Title == sheetName {
				sheetExists = true
				break
			}
		}

		if !sheetExists {
			_, err := srv.Spreadsheets.BatchUpdate(spreadsheetID, &sheets.BatchUpdateSpreadsheetRequest{
				Requests: []*sheets.Request{
					{
						AddSheet: &sheets.AddSheetRequest{
							Properties: &sheets.SheetProperties{
								Title: sheetName,
							},
						},
					},
				},
			}).Do()
			if err != nil {
				return err
			}
		}

		// Clear existing content (this is not replacement but rewrite it)
		_, err = srv.Spreadsheets.Values.Clear(spreadsheetID, sheetName, &sheets.ClearValuesRequest{}).Do()
		if err != nil {
			return err
		}

		// Upload excel to the Google Sheet
		_, err = srv.Spreadsheets.Values.Update(spreadsheetID, sheetName+"!A1", &sheets.ValueRange{
			Values: data,
		}).ValueInputOption("RAW").Do()
		if err != nil {
			return err
		}

		fmt.Printf("Data from '%s' successfully uploaded to Google Sheets.\n", sheetName)
	}
	return nil
}
