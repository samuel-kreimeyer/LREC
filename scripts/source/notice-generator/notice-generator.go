package main

import (
	"encoding/csv"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"text/template"
	"time"

	"github.com/xuri/excelize/v2"
)

const noticeTemplate = `Dear Friends and Engineers,

We're pleased to invite you to the next meeting of the Little Rock Engineers Club for 2025-2026, to be held at {{.Location}} at {{.Time}}. {{.LunchMessage}} Members are welcome to arrive 15 minutes early to enjoy lunch and informal networking with fellow professionals before we begin. We're excited to host guest speaker {{.Speaker}}. {{if .Bio}}{{.Bio}} {{end}}Our topic will be {{.Topic}}.
Meeting Details:

    Location: {{.Location}}
    Time: {{.Time}} (Arrive 15 minutes prior for lunch and networking)
    Speakers: {{.Speaker}}

We look forward to seeing you there and taking part in a great season of learning and collaboration.

Best regards,`

type Event struct {
	Date     time.Time
	Topic    string
	Speaker  string
	Location string
	Time     string
}

type TemplateData struct {
	Date         string
	Topic        string
	Speaker      string
	Location     string
	Time         string
	Bio          string
	LunchMessage string
}

func main() {
	var bio string
	var lunchProvided bool
	var output string
	var templatePath string

	flag.StringVar(&bio, "bio", "", "Speaker bio (optional)")
	flag.BoolVar(&lunchProvided, "lunch-provided", false, "Use 'Lunch will be provided.' instead of default message")
	flag.StringVar(&output, "output", "notices.txt", "Output file path")
	flag.StringVar(&output, "o", "notices.txt", "Output file path (short form)")
	flag.StringVar(&templatePath, "template", "notice_template", "Template file path (ignored - using embedded template)")

	flag.Parse()

	if flag.NArg() < 1 {
		fmt.Fprintf(os.Stderr, "Usage: %s [OPTIONS] SPREADSHEET\n", os.Args[0])
		flag.PrintDefaults()
		os.Exit(1)
	}

	spreadsheet := flag.Arg(0)

	lunchMessage := "Feel free to bring your own lunch."
	if lunchProvided {
		lunchMessage = "Lunch will be provided."
	}

	events, err := readSpreadsheet(spreadsheet)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error reading spreadsheet: %v\n", err)
		os.Exit(1)
	}

	now := time.Now()
	var closestEvent *Event
	var minDiff time.Duration

	for _, event := range events {
		if event.Date.After(now) {
			diff := event.Date.Sub(now)
			if closestEvent == nil || diff < minDiff {
				closestEvent = &event
				minDiff = diff
			}
		}
	}

	if closestEvent == nil {
		fmt.Println("No future events found in the spreadsheet.")
		return
	}

	tmpl, err := template.New("notice").Parse(noticeTemplate)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error parsing template: %v\n", err)
		os.Exit(1)
	}

	data := TemplateData{
		Date:         closestEvent.Date.Format("2006-01-02"),
		Topic:        closestEvent.Topic,
		Speaker:      closestEvent.Speaker,
		Location:     closestEvent.Location,
		Time:         closestEvent.Time,
		Bio:          bio,
		LunchMessage: lunchMessage,
	}

	file, err := os.Create(output)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error creating output file: %v\n", err)
		os.Exit(1)
	}
	defer file.Close()

	err = tmpl.Execute(file, data)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error executing template: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("Generated notice for %s event and saved to %s\n", closestEvent.Date.Format("2006-01-02"), output)
}

func readSpreadsheet(filename string) ([]Event, error) {
	ext := strings.ToLower(filepath.Ext(filename))

	if ext == ".xlsx" || ext == ".xls" {
		return readExcel(filename)
	}
	return readCSV(filename)
}

func readCSV(filename string) ([]Event, error) {
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		return nil, err
	}

	if len(records) < 2 {
		return nil, fmt.Errorf("spreadsheet must have header and at least one data row")
	}

	header := records[0]
	dateIdx, topicIdx, speakerIdx, locationIdx, timeIdx := -1, -1, -1, -1, -1

	for i, col := range header {
		switch strings.ToLower(strings.TrimSpace(col)) {
		case "date":
			dateIdx = i
		case "topic":
			topicIdx = i
		case "speaker":
			speakerIdx = i
		case "location":
			locationIdx = i
		case "time":
			timeIdx = i
		}
	}

	if dateIdx == -1 || topicIdx == -1 || speakerIdx == -1 || locationIdx == -1 || timeIdx == -1 {
		return nil, fmt.Errorf("spreadsheet must have columns: date, topic, speaker, location, time")
	}

	var events []Event
	for _, row := range records[1:] {
		if len(row) <= dateIdx || len(row) <= topicIdx || len(row) <= speakerIdx ||
		   len(row) <= locationIdx || len(row) <= timeIdx {
			continue
		}

		date, err := parseDate(row[dateIdx])
		if err != nil {
			continue
		}

		events = append(events, Event{
			Date:     date,
			Topic:    row[topicIdx],
			Speaker:  row[speakerIdx],
			Location: row[locationIdx],
			Time:     row[timeIdx],
		})
	}

	return events, nil
}

func readExcel(filename string) ([]Event, error) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	if len(rows) < 2 {
		return nil, fmt.Errorf("spreadsheet must have header and at least one data row")
	}

	header := rows[0]
	dateIdx, topicIdx, speakerIdx, locationIdx, timeIdx := -1, -1, -1, -1, -1

	for i, col := range header {
		switch strings.ToLower(strings.TrimSpace(col)) {
		case "date":
			dateIdx = i
		case "topic":
			topicIdx = i
		case "speaker":
			speakerIdx = i
		case "location":
			locationIdx = i
		case "time":
			timeIdx = i
		}
	}

	if dateIdx == -1 || topicIdx == -1 || speakerIdx == -1 || locationIdx == -1 || timeIdx == -1 {
		return nil, fmt.Errorf("spreadsheet must have columns: date, topic, speaker, location, time")
	}

	var events []Event
	for _, row := range rows[1:] {
		if len(row) <= dateIdx || len(row) <= topicIdx || len(row) <= speakerIdx ||
		   len(row) <= locationIdx || len(row) <= timeIdx {
			continue
		}

		date, err := parseDate(row[dateIdx])
		if err != nil {
			continue
		}

		events = append(events, Event{
			Date:     date,
			Topic:    row[topicIdx],
			Speaker:  row[speakerIdx],
			Location: row[locationIdx],
			Time:     row[timeIdx],
		})
	}

	return events, nil
}

func parseDate(dateStr string) (time.Time, error) {
	formats := []string{
		"2006-01-02",
		"01/02/2006",
		"1/2/2006",
		"2006/01/02",
		"02-Jan-2006",
		"2-Jan-2006",
		"Jan 2, 2006",
		"January 2, 2006",
		time.RFC3339,
	}

	for _, format := range formats {
		if t, err := time.Parse(format, dateStr); err == nil {
			return t, nil
		}
	}

	excelEpoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	var days float64
	if _, err := fmt.Sscanf(dateStr, "%f", &days); err == nil && days > 0 {
		return excelEpoch.AddDate(0, 0, int(days)), nil
	}

	return time.Time{}, fmt.Errorf("unable to parse date: %s", dateStr)
}