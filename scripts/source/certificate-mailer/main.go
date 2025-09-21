package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/joho/godotenv"
	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
	"gopkg.in/gomail.v2"
)

type EventInfo struct {
	Date     string
	Topic    string
	Speaker  string
	Location string
	Time     string
}

type Attendee struct {
	Name  string
	Email string
}

type EmailConfig struct {
	SMTPHost    string
	SMTPPort    int
	Email       string
	AppPassword string
}

func main() {
	// Load environment variables
	err := godotenv.Load("../../../.env")
	if err != nil {
		log.Fatalf("Error loading .env file: %v", err)
	}

	// Setup email configuration
	emailConfig := EmailConfig{
		SMTPHost:    "smtp.gmail.com",
		SMTPPort:    587,
		Email:       os.Getenv("GMAIL_EMAIL"),
		AppPassword: os.Getenv("GMAIL_APP_PASSWORD"),
	}

	if emailConfig.Email == "" || emailConfig.AppPassword == "" {
		log.Fatalf("Gmail credentials not found in .env file. Please set GMAIL_EMAIL and GMAIL_APP_PASSWORD")
	}

	// Read roster to get email mappings
	roster, err := readRoster("../../../PII/Roster.xlsx")
	if err != nil {
		log.Fatalf("Error reading roster: %v", err)
	}

	// Read attendance data
	attendees, err := readAttendance("../../../PII/Attendance.xlsx")
	if err != nil {
		log.Fatalf("Error reading attendance: %v", err)
	}

	// Match attendees with email addresses from roster
	attendees = matchAttendeesWithEmails(attendees, roster)

	// Read calendar data and get most recent event
	event, err := getMostRecentEvent("../../../PII/Calendar.xlsx")
	if err != nil {
		log.Fatalf("Error reading calendar: %v", err)
	}

	// Create temp directory for PDFs
	tempDir := "temp_certificates"
	os.MkdirAll(tempDir, 0755)

	// Generate certificates and send individual emails
	sentCount := 0
	for _, attendee := range attendees {
		filePath, err := generateCertificate(attendee, event, tempDir)
		if err != nil {
			log.Printf("Error generating certificate for %s: %v", attendee.Name, err)
			continue
		}
		fmt.Printf("Generated certificate for %s\n", attendee.Name)

		err = sendIndividualCertificateEmail(emailConfig, event, attendee, filePath)
		if err != nil {
			log.Printf("Error sending email to %s: %v", attendee.Name, err)
		} else {
			sentCount++
			fmt.Printf("Email sent to %s (%s)\n", attendee.Name, attendee.Email)
		}

	}

	fmt.Printf("\nSuccessfully generated %d certificates and sent %d emails\n", len(attendees), sentCount)
}

func readAttendance(filepath string) ([]Attendee, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in attendance file")
	}

	rows, err := f.GetRows(sheets[0])
	if err != nil {
		return nil, err
	}

	var attendees []Attendee
	nameCol := -1

	// Find Name column
	if len(rows) > 0 {
		for i, cell := range rows[0] {
			if strings.Contains(strings.ToLower(cell), "name") {
				nameCol = i
				break
			}
		}
	}

	if nameCol == -1 {
		return nil, fmt.Errorf("Name column not found")
	}

	// Read attendee names (skip header row)
	for i := 1; i < len(rows); i++ {
		if len(rows[i]) > nameCol && rows[i][nameCol] != "" {
			name := convertNameFormat(rows[i][nameCol])
			attendees = append(attendees, Attendee{Name: name, Email: ""})
		}
	}

	return attendees, nil
}

func convertNameFormat(name string) string {
	// Convert from "Last, First" to "First Last"
	parts := strings.Split(name, ",")
	if len(parts) == 2 {
		first := strings.TrimSpace(parts[1])
		last := strings.TrimSpace(parts[0])
		return first + " " + last
	}
	return strings.TrimSpace(name)
}

func getMostRecentEvent(filepath string) (EventInfo, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return EventInfo{}, err
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return EventInfo{}, fmt.Errorf("no sheets found in calendar file")
	}

	rows, err := f.GetRows(sheets[0])
	if err != nil {
		return EventInfo{}, err
	}

	// Find column indices - check first two rows for headers
	dateCol, topicCol, speakerCol, locationCol, timeCol := -1, -1, -1, -1, -1
	headerRow := 0

	for rowIdx := 0; rowIdx < 2 && rowIdx < len(rows); rowIdx++ {
		for i, cell := range rows[rowIdx] {
			cellLower := strings.ToLower(cell)
			if strings.Contains(cellLower, "date") {
				dateCol = i
				headerRow = rowIdx
			} else if strings.Contains(cellLower, "topic") {
				topicCol = i
			} else if strings.Contains(cellLower, "speaker") {
				speakerCol = i
			} else if strings.Contains(cellLower, "location") {
				locationCol = i
			} else if strings.Contains(cellLower, "time") {
				timeCol = i
			}
		}
		if dateCol != -1 && topicCol != -1 && speakerCol != -1 {
			break
		}
	}

	if dateCol == -1 || topicCol == -1 || speakerCol == -1 {
		return EventInfo{}, fmt.Errorf("required columns not found")
	}

	// Find the most recent non-empty event
	var events []EventInfo
	for i := headerRow + 1; i < len(rows); i++ {
		if len(rows[i]) > dateCol && rows[i][dateCol] != "" {
			event := EventInfo{}
			event.Date = rows[i][dateCol]

			if len(rows[i]) > topicCol {
				event.Topic = rows[i][topicCol]
			}
			if len(rows[i]) > speakerCol {
				event.Speaker = rows[i][speakerCol]
			}
			if locationCol != -1 && len(rows[i]) > locationCol {
				event.Location = rows[i][locationCol]
			}
			if timeCol != -1 && len(rows[i]) > timeCol {
				event.Time = rows[i][timeCol]
			}

			if event.Topic != "" && event.Speaker != "" {
				events = append(events, event)
			}
		}
	}

	if len(events) == 0 {
		return EventInfo{}, fmt.Errorf("no valid events found")
	}

	// Filter events to only include past events and sort by date to get most recent past event
	now := time.Now()
	var pastEvents []EventInfo

	for _, event := range events {
		eventDate, err := parseFlexibleDate(event.Date)
		if err == nil && eventDate.Before(now) {
			pastEvents = append(pastEvents, event)
		}
	}

	// If no past events, use all events (fallback)
	if len(pastEvents) == 0 {
		pastEvents = events
	}

	// Sort past events by date to get most recent
	sort.Slice(pastEvents, func(i, j int) bool {
		// Try to parse dates
		date1, err1 := parseFlexibleDate(pastEvents[i].Date)
		date2, err2 := parseFlexibleDate(pastEvents[j].Date)

		if err1 == nil && err2 == nil {
			return date1.After(date2)
		}
		// If parsing fails, do string comparison
		return pastEvents[i].Date > pastEvents[j].Date
	})

	return pastEvents[0], nil
}

func parseFlexibleDate(dateStr string) (time.Time, error) {
	formats := []string{
		"01/02/2006",
		"1/2/2006",
		"1/02/2006",
		"01/2/2006",
		"2006-01-02",
		"January 2, 2006",
		"Jan 2, 2006",
		"2 January 2006",
		"2 Jan 2006",
	}

	for _, format := range formats {
		if t, err := time.Parse(format, dateStr); err == nil {
			return t, nil
		}
	}

	return time.Time{}, fmt.Errorf("unable to parse date: %s", dateStr)
}

func readRoster(filepath string) (map[string]string, error) {
	f, err := excelize.OpenFile(filepath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in roster file")
	}

	rows, err := f.GetRows(sheets[0])
	if err != nil {
		return nil, err
	}

	nameToEmail := make(map[string]string)
	nameCol, emailCol := -1, -1

	// Find Name and Email columns
	if len(rows) > 0 {
		for i, cell := range rows[0] {
			cellLower := strings.ToLower(cell)
			if strings.Contains(cellLower, "name") {
				nameCol = i
			} else if strings.Contains(cellLower, "email") {
				emailCol = i
			}
		}
	}

	if nameCol == -1 || emailCol == -1 {
		return nil, fmt.Errorf("Name or Email column not found in roster")
	}

	// Read name-email mappings (skip header row)
	for i := 1; i < len(rows); i++ {
		if len(rows[i]) > nameCol && len(rows[i]) > emailCol {
			name := strings.TrimSpace(rows[i][nameCol])
			email := strings.TrimSpace(rows[i][emailCol])
			if name != "" && email != "" {
				// Convert name to match attendance format
				name = convertNameFormat(name)
				nameToEmail[name] = email
			}
		}
	}

	return nameToEmail, nil
}

func matchAttendeesWithEmails(attendees []Attendee, roster map[string]string) []Attendee {
	for i, attendee := range attendees {
		if email, found := roster[attendee.Name]; found {
			attendees[i].Email = email
		}
	}
	return attendees
}

func generateCertificate(attendee Attendee, event EventInfo, outputDir string) (string, error) {
	// Create PDF in landscape orientation - US Letter
	pdf := gofpdf.New("L", "mm", "Letter", "")
	pdf.AddPage()

	// Set up the certificate layout
	pageWidth, _ := pdf.GetPageSize()

	// Add skyline image at the top left
	skylinePath := "../../../scripts/skyline.png"
	imageInfo := pdf.RegisterImage(skylinePath, "PNG")
	if imageInfo != nil {
		// Place skyline image at top left
		pdf.ImageOptions(skylinePath, 25, 15, 50, 0, false, gofpdf.ImageOptions{ImageType: "PNG", ReadDpi: true}, 0, "")

		// Add "LITTLE ROCK ENGINEERS CLUB" text next to skyline at top - same font size as name (24pt)
		pdf.SetFont("Times", "B", 24)
		pdf.SetXY(80, 25)
		pdf.Cell(0, 10, "LITTLE ROCK ENGINEERS CLUB")
	}

	// Add main title - large and centered (moved closer to header)
	pdf.SetFont("Times", "B", 36)
	pdf.SetXY(0, 55)
	pdf.SetTextColor(0, 0, 0)
	titleWidth := pdf.GetStringWidth("CERTIFICATE OF ATTENDANCE")
	titleX := (pageWidth - titleWidth) / 2
	pdf.SetX(titleX)
	pdf.Cell(titleWidth, 15, "CERTIFICATE OF ATTENDANCE")

	// Add certification text - centered (moved up 25mm = 1 inch)
	pdf.SetFont("Times", "", 18)
	pdf.SetXY(0, 70)
	certTextWidth := pdf.GetStringWidth("This is to certify that")
	certTextX := (pageWidth - certTextWidth) / 2
	pdf.SetX(certTextX)
	pdf.Cell(certTextWidth, 10, "This is to certify that")

	// Add attendee name with underline - properly centered with center alignment (moved up 25mm)
	pdf.SetFont("Times", "B", 24)
	pdf.SetXY(0, 95)
	// Use CellFormat with center alignment for proper centering
	pdf.CellFormat(pageWidth, 10, attendee.Name, "", 0, "C", false, 0, "")
	// Draw underline centered under the name
	nameWidth := pdf.GetStringWidth(attendee.Name)
	nameX := (pageWidth - nameWidth) / 2
	pdf.Line(nameX, 107, nameX+nameWidth, 107)

	// Add earned PDH text - centered (moved up 25mm)
	pdf.SetFont("Times", "", 16)
	pdf.SetXY(0, 120)
	pdhText := "Earned one (1) Professional Development Hour (PDH) by attending"
	pdhWidth := pdf.GetStringWidth(pdhText)
	pdhX := (pageWidth - pdhWidth) / 2
	pdf.SetX(pdhX)
	pdf.Cell(pdhWidth, 10, pdhText)

	pdf.SetXY(0, 135)
	presentationText := "the presentation by:"
	presentationWidth := pdf.GetStringWidth(presentationText)
	presentationX := (pageWidth - presentationWidth) / 2
	pdf.SetX(presentationX)
	pdf.Cell(presentationWidth, 10, presentationText)

	// Add speaker and title - centered (moved up 25mm)
	pdf.SetFont("Times", "I", 18)
	pdf.SetXY(0, 150)
	speakerWidth := pdf.GetStringWidth(event.Speaker)
	speakerX := (pageWidth - speakerWidth) / 2
	pdf.SetX(speakerX)
	pdf.Cell(speakerWidth, 10, event.Speaker)

	pdf.SetXY(0, 165)
	topicWidth := pdf.GetStringWidth(event.Topic)
	topicX := (pageWidth - topicWidth) / 2
	pdf.SetX(topicX)
	pdf.Cell(topicWidth, 10, event.Topic)

	// Add location and date - centered (moved up 25mm)
	pdf.SetFont("Times", "", 16)
	pdf.SetXY(0, 185)
	locationText := fmt.Sprintf("Conducted in Little Rock, Arkansas on %s", event.Date)
	locationWidth := pdf.GetStringWidth(locationText)
	locationX := (pageWidth - locationWidth) / 2
	pdf.SetX(locationX)
	pdf.Cell(locationWidth, 10, locationText)

	// Generate filename
	cleanName := strings.ReplaceAll(attendee.Name, " ", "_")
	cleanDate := strings.ReplaceAll(event.Date, "/", "-")
	filename := fmt.Sprintf("COA_%s_%s.pdf", cleanName, cleanDate)
	filepath := filepath.Join(outputDir, filename)

	err := pdf.OutputFileAndClose(filepath)
	if err != nil {
		return "", err
	}
	return filepath, nil
}

func sendIndividualCertificateEmail(config EmailConfig, event EventInfo, attendee Attendee, certificatePath string) error {
	// Create email message
	m := gomail.NewMessage()

	recipient := attendee.Email
	if recipient == "" {
		return fmt.Errorf("no email address for attendee %s", attendee.Name)
	}

	// Set email headers
	m.SetHeader("From", config.Email)
	m.SetHeader("To", recipient)
	m.SetHeader("Subject", fmt.Sprintf("LREC Certificate of Attendance - %s - %s", attendee.Name, event.Date))

	// Create email body
	body := fmt.Sprintf(`Dear %s,

Please find attached your Certificate of Attendance for the Little Rock Engineers Club presentation:

Speaker: %s
Topic: %s
Date: %s

Thank you for attending this presentation.

Best regards,
Little Rock Engineers Club`, attendee.Name, event.Speaker, event.Topic, event.Date)

	m.SetBody("text/plain", body)

	// Attach the individual certificate
	m.Attach(certificatePath)

	// Create SMTP dialer
	d := gomail.NewDialer(config.SMTPHost, config.SMTPPort, config.Email, config.AppPassword)

	// Send email
	if err := d.DialAndSend(m); err != nil {
		return fmt.Errorf("failed to send email: %v", err)
	}

	return nil
}
