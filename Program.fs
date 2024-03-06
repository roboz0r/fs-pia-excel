open Microsoft.Office.Interop.Excel

let app = new ApplicationClass()
let wb = app.Workbooks.Add()

(wb.ActiveSheet :?> Worksheet).Cells.[1,1] <- "Hello from F#"

wb.SaveAs(@"C:\Users\Robert\Desktop\test.xlsx")
wb.Close()
app.Quit()