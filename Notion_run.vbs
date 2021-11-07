dim objExcel
set objExcel = createobject("Excel.Application")

objExcel.Workbooks.Open "C:\Users\smoka\OneDrive\0-MAIN\Notion.xlsm"
objExcel.visible = True

objExcel.Run "notionMobile"

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close(0)

objExcel.Echo "Finished." 'Tutaj możemy wpisać komunikat, który wyświetli się po uruchomieniu makra
objExcel.Quit
