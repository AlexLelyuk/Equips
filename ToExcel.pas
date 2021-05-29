unit ToExcel;

interface

uses Grids;

 function SaveAsExcelFile(AGrid: TStringGrid; ASheetName, AFileName: string): Boolean;
 function SaveAsCSV(AGrid: TStringGrid; ASheetName, AFileName: string): Boolean;

implementation

uses ComObj, SysUtils, Variants;

 function RefToCell(ARow, ACol: Integer): string;
 begin
   Result := Chr(Ord('A') + ACol - 1) + IntToStr(ARow);
 end;

 function SaveAsCSV(AGrid: TStringGrid; ASheetName, AFileName: string): Boolean;
 var i, j : integer;
     f : TextFile;
 begin
    AssignFile(f,'ObrData.csv');
    Rewrite(f);
    for i := 0 to AGrid.RowCount - 1 do begin
       for j := 0 to AGrid.ColCount - 1 do begin
          if copy(AGrid.Cells[j, i],1,1)<>'+' then Write(f,AGrid.Cells[j, i]+';')
          else Write(f,copy(AGrid.Cells[j, i],2)+';');
       end;
       Writeln(f);
    end;
    CloseFile(f);
    Result:=True;
 end;

 function SaveAsExcelFile(AGrid: TStringGrid; ASheetName, AFileName: string): Boolean;
 const
   xlWBATWorksheet = -4167;
 var
   XLApp, Sheet{, Data}: OLEVariant;
   i, j: Integer;
 begin
  Result := False;
   XLApp := CreateOleObject('Excel.Application');
   try
     // Hide Excel
    XLApp.Visible := False;
     // Add new Workbook
    XLApp.Workbooks.Add(xlWBatWorkSheet);
     Sheet := XLApp.Workbooks[1].WorkSheets[1];
     Sheet.Name := ASheetName;

   for i := 0 to AGrid.ColCount - 1 do
     for j := 0 to AGrid.RowCount - 1 do
        Sheet.Cells[j+1, i+1].Value := AGrid.Cells[i, j];

     // Save Excel Worksheet
    try
       XLApp.Workbooks[1].SaveAs(AFileName);
       Result := True;
     except
       // Error ?
    end;
   finally
     // Quit Excel
    if not VarIsEmpty(XLApp) then
     begin
       XLApp.DisplayAlerts := False;
       XLApp.Quit;
       XLAPP := Unassigned;
       Sheet := Unassigned;
     end;
   end;
 end;

end.
