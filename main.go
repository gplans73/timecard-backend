package main

import (
    "bytes"
    "encoding/base64"
    "encoding/json"
    "fmt"
    "log"
    "net/http"
    "net/smtp"
    "os"
    "os/exec"
    "path/filepath"
    "strings"
    "time"

    "github.com/xuri/excelize/v2"
)

/* =========================
   Models (match your Swift)
   ========================= */

type TimecardRequest struct {
    EmployeeName    string     `json:"employee_name"`
    PayPeriodNum    int        `json:"pay_period_num"`
    Year            int        `json:"year"`
    WeekStartDate   string     `json:"week_start_date"`
    WeekNumberLabel string     `json:"week_number_label"`
    Jobs            []Job      `json:"jobs"`
    Entries         []Entry    `json:"entries"`
    Weeks           []WeekData `json:"weeks,omitempty"`
}

type Job struct {
    // JobCode is the JOB NUMBER (e.g., "29699", "12215")
    JobCode string `json:"job_code"`
    // JobName is the LABOUR CODE (e.g., "201", "223", "H")
    JobName string `json:"job_name"`
}

type Entry struct {
    Date         string  `json:"date"`
    JobCode      string  `json:"job_code"` // JOB NUMBER
    Hours        float64 `json:"hours"`
    Overtime     bool    `json:"overtime"`
    IsNightShift bool    `json:"is_night_shift"`
}

// accept both snake_case and camelCase keys
func (e *Entry) UnmarshalJSON(data []byte) error {
    type rawEntry struct {
        Date              string  `json:"date"`
        JobCode           string  `json:"job_code"`
        Code              string  `json:"code"`
        Hours             float64 `json:"hours"`
        Overtime          *bool   `json:"overtime"`
        IsOvertimeCamel   *bool   `json:"isOvertime"`
        NightShift        *bool   `json:"night_shift"`
        IsNightShiftSnake *bool   `json:"is_night_shift"`
        IsNightShiftCamel *bool   `json:"isNightShift"`
    }
    var aux rawEntry
    if err := json.Unmarshal(data, &aux); err != nil {
        return err
    }

    e.Date = aux.Date
    if aux.JobCode != "" {
        e.JobCode = aux.JobCode
    } else {
        e.JobCode = aux.Code
    }
    e.Hours = aux.Hours

    if aux.Overtime != nil {
        e.Overtime = *aux.Overtime
    } else if aux.IsOvertimeCamel != nil {
        e.Overtime = *aux.IsOvertimeCamel
    }

    if aux.NightShift != nil {
        e.IsNightShift = *aux.NightShift
    } else if aux.IsNightShiftSnake != nil {
        e.IsNightShift = *aux.IsNightShiftSnake
    } else if aux.IsNightShiftCamel != nil {
        e.IsNightShift = *aux.IsNightShiftCamel
    }

    log.Printf("  Unmarshaled entry: JobCode=%s, Hours=%.2f, OT=%v, Night=%v",
        e.JobCode, e.Hours, e.Overtime, e.IsNightShift)
    return nil
}

type WeekData struct {
    WeekNumber    int     `json:"week_number"`
    WeekStartDate string  `json:"week_start_date"`
    WeekLabel     string  `json:"week_label"`
    Entries       []Entry `json:"entries"`
}

type EmailTimecardRequest struct {
    TimecardRequest
    To      string  `json:"to"`
    CC      *string `json:"cc"`
    Subject string  `json:"subject"`
    Body    string  `json:"body"`
}

/* ===============
   Server bootstrap
   =============== */

func main() {
    port := os.Getenv("PORT")
    if port == "" {
        port = "8080"
    }

    http.HandleFunc("/health", healthHandler)
    http.HandleFunc("/api/generate-timecard", corsMiddleware(generateTimecardHandler))
    http.HandleFunc("/api/generate-pdf", corsMiddleware(generatePDFHandler))
    http.HandleFunc("/api/email-timecard", corsMiddleware(emailTimecardHandler))

    log.Printf("Server starting on :%s ...", port)
    if err := http.ListenAndServe(":"+port, nil); err != nil {
        log.Fatal(err)
    }
}

func healthHandler(w http.ResponseWriter, r *http.Request) {
    w.WriteHeader(http.StatusOK)
    _, _ = w.Write([]byte("OK"))
}

func corsMiddleware(next http.HandlerFunc) http.HandlerFunc {
    return func(w http.ResponseWriter, r *http.Request) {
        w.Header().Set("Access-Control-Allow-Origin", "*")
        w.Header().Set("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        w.Header().Set("Access-Control-Allow-Headers", "Content-Type, Authorization")
        if r.Method == http.MethodOptions {
            w.WriteHeader(http.StatusOK)
            return
        }
        next(w, r)
    }
}

/* ===================
   API: Generate / Mail
   =================== */

func generateTimecardHandler(w http.ResponseWriter, r *http.Request) {
    if r.Method != http.MethodPost {
        http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
        return
    }

    var req TimecardRequest
    if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
        log.Printf("decode error: %v", err)
        http.Error(w, fmt.Sprintf("invalid request: %v", err), http.StatusBadRequest)
        return
    }

    log.Printf("Generating timecard for %s", req.EmployeeName)

    excelData, err := generateExcelFile(req)
    if err != nil {
        log.Printf("excel error: %v", err)
        http.Error(w, fmt.Sprintf("error generating timecard: %v", err), http.StatusInternalServerError)
        return
    }

    w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=\"timecard_%s.xlsx\"", req.EmployeeName))
    w.WriteHeader(http.StatusOK)
    _, _ = w.Write(excelData)

    log.Printf("OK: timecard bytes=%d", len(excelData))
}

func generatePDFHandler(w http.ResponseWriter, r *http.Request) {
    if r.Method != http.MethodPost {
        http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
        return
    }

    var req TimecardRequest
    if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
        log.Printf("decode error: %v", err)
        http.Error(w, fmt.Sprintf("invalid request: %v", err), http.StatusBadRequest)
        return
    }

    log.Printf("Generating PDF timecard for %s", req.EmployeeName)

    // First generate Excel
    excelData, err := generateExcelFile(req)
    if err != nil {
        log.Printf("excel error: %v", err)
        http.Error(w, fmt.Sprintf("error generating Excel: %v", err), http.StatusInternalServerError)
        return
    }

    // Convert to PDF
    pdfData, err := generatePDFFromExcel(excelData, fmt.Sprintf("timecard_%s.xlsx", req.EmployeeName))
    if err != nil {
        log.Printf("pdf conversion error: %v", err)
        http.Error(w, fmt.Sprintf("error converting to PDF: %v", err), http.StatusInternalServerError)
        return
    }

    w.Header().Set("Content-Type", "application/pdf")
    w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=\"timecard_%s.pdf\"", req.EmployeeName))
    w.WriteHeader(http.StatusOK)
    _, _ = w.Write(pdfData)

    log.Printf("OK: PDF bytes=%d", len(pdfData))
}

func emailTimecardHandler(w http.ResponseWriter, r *http.Request) {
    if r.Method != http.MethodPost {
        http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
        return
    }

    var req EmailTimecardRequest
    if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
        log.Printf("decode error: %v", err)
        http.Error(w, fmt.Sprintf("invalid request: %v", err), http.StatusBadRequest)
        return
    }

    log.Printf("Emailing timecard for %s → %s", req.EmployeeName, req.To)

    excelData, err := generateExcelFile(req.TimecardRequest)
    if err != nil {
        log.Printf("excel error: %v", err)
        http.Error(w, fmt.Sprintf("error generating timecard: %v", err), http.StatusInternalServerError)
        return
    }

    if err := sendEmail(req.To, req.CC, req.Subject, req.Body, excelData, req.EmployeeName); err != nil {
        log.Printf("send email error: %v", err)
        http.Error(w, fmt.Sprintf("error sending email: %v", err), http.StatusInternalServerError)
        return
    }

    w.Header().Set("Content-Type", "application/json")
    _ = json.NewEncoder(w).Encode(map[string]string{
        "status":  "success",
        "message": fmt.Sprintf("Email sent to %s", req.To),
    })
}

/* ===========================
   Excel generation (Excelize)
   =========================== */

func generateExcelFile(req TimecardRequest) ([]byte, error) {
    const templatePath = "template.xlsx"

    f, err := excelize.OpenFile(templatePath)
    if err != nil {
        log.Printf("Template not found, using basic file: %v", err)
        return generateBasicExcelFile(req)
    }
    defer func() { _ = f.Close() }()

    sheets := f.GetSheetList()
    if len(sheets) == 0 {
        return nil, fmt.Errorf("no sheets in template")
    }

    if len(req.Weeks) > 0 {
        if err := fillWeekSheet(f, sheets[0], req, req.Weeks[0], 1); err != nil {
            log.Printf("Week 1 fill error: %v", err)
        }
    }
    if len(sheets) > 1 && len(req.Weeks) > 1 {
        if err := fillWeekSheet(f, sheets[1], req, req.Weeks[1], 2); err != nil {
            log.Printf("Week 2 fill error: %v", err)
        }
    }

    // Clear cached values so Excel recalculates on open
    if err := f.UpdateLinkedValue(); err != nil {
        log.Printf("UpdateLinkedValue warning: %v", err)
    }

    buf, err := f.WriteToBuffer()
    if err != nil {
        return nil, err
    }
    return buf.Bytes(), nil
}

// Generate PDF from Excel using LibreOffice (pixel-perfect conversion)
func generatePDFFromExcel(excelData []byte, filename string) ([]byte, error) {
    // Save Excel data to temp file
    tmpExcel, err := os.CreateTemp("", "timecard-*.xlsx")
    if err != nil {
        return nil, fmt.Errorf("create temp excel: %w", err)
    }
    tmpExcelPath := tmpExcel.Name()
    defer os.Remove(tmpExcelPath)

    if _, err := tmpExcel.Write(excelData); err != nil {
        tmpExcel.Close()
        return nil, fmt.Errorf("write excel: %w", err)
    }
    tmpExcel.Close()

    // Create temp output directory for PDF
    tmpDir, err := os.MkdirTemp("", "pdf-")
    if err != nil {
        return nil, fmt.Errorf("create temp dir: %w", err)
    }
    defer os.RemoveAll(tmpDir)

    log.Printf("🔄 Converting Excel to PDF using LibreOffice...")

    // Convert using LibreOffice headless mode
    cmd := exec.Command(
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", tmpDir,
        tmpExcelPath,
    )

    // Capture output for debugging
    output, err := cmd.CombinedOutput()
    if err != nil {
        log.Printf("❌ LibreOffice conversion failed: %s", string(output))
        return nil, fmt.Errorf("libreoffice conversion failed: %w\nOutput: %s", err, string(output))
    }

    log.Printf("LibreOffice output: %s", string(output))

    // Find the generated PDF file
    files, err := os.ReadDir(tmpDir)
    if err != nil {
        return nil, fmt.Errorf("read output dir: %w", err)
    }

    if len(files) == 0 {
        return nil, fmt.Errorf("no PDF generated by LibreOffice")
    }

    // Read the PDF file
    pdfPath := filepath.Join(tmpDir, files[0].Name())
    pdfData, err := os.ReadFile(pdfPath)
    if err != nil {
        return nil, fmt.Errorf("read pdf: %w", err)
    }

    log.Printf("✅ Generated LibreOffice PDF: %d bytes (perfect Excel conversion)", len(pdfData))
    return pdfData, nil
}

// Apply borders to a range of cells
func applyBordersToRange(f *excelize.File, sheet string, startCell string, endCell string) error {
    style, err := f.NewStyle(&excelize.Style{
        Border: []excelize.Border{
            {Type: "left", Color: "000000", Style: 1},
            {Type: "top", Color: "000000", Style: 1},
            {Type: "bottom", Color: "000000", Style: 1},
            {Type: "right", Color: "000000", Style: 1},
        },
    })
    if err != nil {
        return err
    }
    return f.SetCellStyle(sheet, startCell, endCell, style)
}

// Fill a single week sheet with headers and daily hours
func fillWeekSheet(f *excelize.File, sheet string, req TimecardRequest, week WeekData, weekNum int) error {
    weekStart, err := time.Parse(time.RFC3339, week.WeekStartDate)
    if err != nil {
        return fmt.Errorf("parse week start: %w", err)
    }
    log.Printf("=== Filling %s (week %d) start=%s entries=%d ===",
        sheet, weekNum, weekStart.Format("2006-01-02"), len(week.Entries))

    // Header info - just set values
    _ = f.SetCellValue(sheet, "M2", req.EmployeeName)
    _ = f.SetCellValue(sheet, "AJ2", req.PayPeriodNum)
    _ = f.SetCellValue(sheet, "AJ3", req.Year)
    _ = f.SetCellValue(sheet, "B4", timeToExcelDate(weekStart))
    _ = f.SetCellValue(sheet, "AJ4", week.WeekLabel)

    // Columns: labour codes in C,E,G,... and job numbers in D,F,H,...
    codeCols := []string{"C", "E", "G", "I", "K", "M", "O", "Q", "S", "U", "W", "Y", "AA", "AC", "AE", "AG"}
    jobCols := []string{"D", "F", "H", "J", "L", "N", "P", "R", "T", "V", "X", "Z", "AB", "AD", "AF", "AH"}

    // Job lookup by job number
    jobMap := make(map[string]*Job, len(req.Jobs))
    for i := range req.Jobs {
        jobMap[req.Jobs[i].JobCode] = &req.Jobs[i]
    }

    regularKeys := getUniqueJobNumbersForType(week.Entries, false)
    overtimeKeys := getUniqueJobNumbersForType(week.Entries, true)

    // Fill row 4 (regular headers)
    if len(regularKeys) > 0 {
        for i, key := range regularKeys {
            if i >= len(codeCols) {
                break
            }
            actual := key
            night := strings.HasPrefix(key, "N")
            if night {
                actual = key[1:]
            }
            if job := jobMap[actual]; job != nil {
                code := job.JobName
                if night {
                    code = "N" + code
                }
                _ = f.SetCellValue(sheet, codeCols[i]+"4", code)
                _ = f.SetCellValue(sheet, jobCols[i]+"4", actual)
                log.Printf("  regular header %s4=%s (code), %s4=%s (job)",
                    codeCols[i], code, jobCols[i], actual)
            }
        }
    }

    // Fill row 15 (overtime headers)
    if len(overtimeKeys) > 0 {
        for i, key := range overtimeKeys {
            if i >= len(codeCols) {
                break
            }
            actual := key
            night := strings.HasPrefix(key, "N")
            if night {
                actual = key[1:]
            }
            if job := jobMap[actual]; job != nil {
                code := job.JobName
                if night {
                    code = "N" + code
                }
                _ = f.SetCellValue(sheet, codeCols[i]+"15", code)
                _ = f.SetCellValue(sheet, jobCols[i]+"15", actual)
                log.Printf("  overtime header %s15=%s (code), %s15=%s (job)",
                    codeCols[i], code, jobCols[i], actual)
            }
        }
    }

    // Aggregate hours by date+job key
    regMap := make(map[string]map[string]float64)
    otMap := make(map[string]map[string]float64)

    for _, e := range week.Entries {
        t, err := time.Parse(time.RFC3339, e.Date)
        if err != nil {
            log.Printf("  bad entry date %q: %v", e.Date, err)
            continue
        }
        date := t.Format("2006-01-02")
        key := e.JobCode
        if e.IsNightShift {
            key = "N" + key
        }

        if e.Overtime {
            if otMap[date] == nil {
                otMap[date] = map[string]float64{}
            }
            otMap[date][key] += e.Hours
        } else {
            if regMap[date] == nil {
                regMap[date] = map[string]float64{}
            }
            regMap[date][key] += e.Hours
        }
    }

    // Write dates + hours
    for d := 0; d < 7; d++ {
        day := weekStart.AddDate(0, 0, d)
        dateKey := day.Format("2006-01-02")
        dateSerial := timeToExcelDate(day)

        rowReg := 5 + d
        rowOT := 16 + d

        _ = f.SetCellValue(sheet, fmt.Sprintf("B%d", rowReg), dateSerial)
        _ = f.SetCellValue(sheet, fmt.Sprintf("B%d", rowOT), dateSerial)

        if hours := regMap[dateKey]; hours != nil {
            for i, key := range regularKeys {
                if i >= len(codeCols) {
                    break
                }
                if v, ok := hours[key]; ok && v != 0 {
                    cell := fmt.Sprintf("%s%d", codeCols[i], rowReg)
                    _ = f.SetCellValue(sheet, cell, v)
                    log.Printf("    REG %s = %.2f (%s)", cell, v, key)
                }
            }
        }
        if hours := otMap[dateKey]; hours != nil {
            for i, key := range overtimeKeys {
                if i >= len(codeCols) {
                    break
                }
                if v, ok := hours[key]; ok && v != 0 {
                    cell := fmt.Sprintf("%s%d", codeCols[i], rowOT)
                    _ = f.SetCellValue(sheet, cell, v)
                    log.Printf("    OT  %s = %.2f (%s)", cell, v, key)
                }
            }
        }
    }

    // Apply borders to the entire Regular Time table (rows 4-11, columns A-AJ)
    log.Printf("Applying borders to Regular Time table...")
    if err := applyBordersToRange(f, sheet, "A4", "AJ12"); err != nil {
        log.Printf("Warning: Failed to apply borders to regular table: %v", err)
    }

    // Apply borders to the entire Overtime table (rows 15-23, columns A-AJ)  
    log.Printf("Applying borders to Overtime table...")
    if err := applyBordersToRange(f, sheet, "A15", "AJ24"); err != nil {
        log.Printf("Warning: Failed to apply borders to overtime table: %v", err)
    }

    log.Printf("=== %s week %d done ===", sheet, weekNum)
    return nil
}

func getUniqueJobNumbersForType(entries []Entry, isOvertime bool) []string {
    seen := make(map[string]bool)
    var out []string
    for _, e := range entries {
        if e.Overtime != isOvertime {
            continue
        }
        key := e.JobCode
        if e.IsNightShift {
            key = "N" + key
        }
        if !seen[key] {
            seen[key] = true
            out = append(out, key)
        }
    }
    return out
}

func timeToExcelDate(t time.Time) float64 {
    excelEpoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
    return t.Sub(excelEpoch).Hours() / 24.0
}

func generateBasicExcelFile(req TimecardRequest) ([]byte, error) {
    f := excelize.NewFile()
    defer func() { _ = f.Close() }()
    const sheet = "Sheet1"
    _ = f.SetCellValue(sheet, "A1", "Employee:")
    _ = f.SetCellValue(sheet, "B1", req.EmployeeName)
    buf, err := f.WriteToBuffer()
    if err != nil {
        return nil, err
    }
    return buf.Bytes(), nil
}

/* ==========
   Email utils
   ========== */

func sendEmail(to string, cc *string, subject string, body string, attachment []byte, employeeName string) error {
    smtpHost := os.Getenv("SMTP_HOST")
    smtpPort := os.Getenv("SMTP_PORT")
    smtpUser := os.Getenv("SMTP_USER")
    smtpPass := os.Getenv("SMTP_PASS")
    fromEmail := os.Getenv("SMTP_FROM")

    if smtpHost == "" || smtpPort == "" || smtpUser == "" || smtpPass == "" {
        return fmt.Errorf("SMTP not configured")
    }
    if fromEmail == "" {
        fromEmail = smtpUser
    }

    recipients := strings.Split(to, ",")
    for i := range recipients {
        recipients[i] = strings.TrimSpace(recipients[i])
    }

    var ccRecipients []string
    if cc != nil && *cc != "" {
        ccRecipients = strings.Split(*cc, ",")
        for i := range ccRecipients {
            ccRecipients[i] = strings.TrimSpace(ccRecipients[i])
        }
    }

    all := append([]string{}, recipients...)
    all = append(all, ccRecipients...)

    fileName := fmt.Sprintf("timecard_%s_%s.xlsx",
        strings.ReplaceAll(employeeName, " ", "_"),
        time.Now().Format("2006-01-02"))

    msg := buildEmailMessage(fromEmail, recipients, ccRecipients, subject, body, attachment, fileName)
    auth := smtp.PlainAuth("", smtpUser, smtpPass, smtpHost)
    addr := fmt.Sprintf("%s:%s", smtpHost, smtpPort)
    return smtp.SendMail(addr, auth, fromEmail, all, []byte(msg))
}

func buildEmailMessage(from string, to []string, cc []string, subject string, body string, attachment []byte, fileName string) string {
    boundary := "==BOUNDARY=="
    var buf bytes.Buffer

    buf.WriteString(fmt.Sprintf("From: %s\r\n", from))
    buf.WriteString(fmt.Sprintf("To: %s\r\n", strings.Join(to, ", ")))
    if len(cc) > 0 {
        buf.WriteString(fmt.Sprintf("Cc: %s\r\n", strings.Join(cc, ", ")))
    }
    buf.WriteString(fmt.Sprintf("Subject: %s\r\n", subject))
    buf.WriteString("MIME-Version: 1.0\r\n")
    buf.WriteString(fmt.Sprintf("Content-Type: multipart/mixed; boundary=\"%s\"\r\n\r\n", boundary))

    // body
    buf.WriteString(fmt.Sprintf("--%s\r\n", boundary))
    buf.WriteString("Content-Type: text/plain; charset=\"utf-8\"\r\n\r\n")
    buf.WriteString(body + "\r\n\r\n")

    // attachment
    if len(attachment) > 0 {
        buf.WriteString(fmt.Sprintf("--%s\r\n", boundary))
        buf.WriteString("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n")
        buf.WriteString(fmt.Sprintf("Content-Disposition: attachment; filename=\"%s\"\r\n", fileName))
        buf.WriteString("Content-Transfer-Encoding: base64\r\n\r\n")
        enc := base64.StdEncoding.EncodeToString(attachment)
        for i := 0; i < len(enc); i += 76 {
            end := i + 76
            if end > len(enc) {
                end = len(enc)
            }
            buf.WriteString(enc[i:end] + "\r\n")
        }
        buf.WriteString("\r\n")
    }

    buf.WriteString(fmt.Sprintf("--%s--\r\n", boundary))
    return buf.String()
}
