unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ToolWin, ComCtrls, ImgList, Menus, ExtCtrls, StdCtrls, Buttons,
  DateUtils, DB, System.Win.ComObj, System.Types, System.UITypes,
  Grids, math, IniFiles, common, AppEvnts, Synpdf, ShellAPI, IdFTP, IdAllFTPListParsers, IdException,
  System.ImageList, IdMessage, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack,
  IdSSL, IdSSLOpenSSL, IdMessageClient, IdSMTPBase, IdSMTP, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase, EquipClass;

type


  TForm1 = class(TForm)
    PageControl1: TPageControl;
    StatusBar1: TStatusBar;
    TabSheet2: TTabSheet;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    Excel1: TMenuItem;
    N8: TMenuItem;
    ImageList1: TImageList;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolBar2: TToolBar;
    ToolBar3: TToolBar;
    DT1: TDateTimePicker;
    ToolButton1: TToolButton;
    ToolButton2: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ScrollBox1: TScrollBox;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    TabSheet1: TTabSheet;
    StringGrid1: TStringGrid;
    Button1: TButton;
    N9: TMenuItem;
    TabSheet3: TTabSheet;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    ComboBox1: TComboBox;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ApplicationEvents1: TApplicationEvents;
    N10: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    StringGrid2: TStringGrid;
    Button5: TButton;
    IdFTP1: TIdFTP;
    ProgressBar1: TProgressBar;
    ToolButton13: TToolButton;
    Sendemail1: TMenuItem;
    IdSMTP1: TIdSMTP;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    IdMessage1: TIdMessage;
    N14: TMenuItem;
    N15: TMenuItem;
    CheckBox1: TCheckBox;
    procedure N7Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure DrawHoldingGraph;
    procedure GetSpeedAllAndDrawGraph;
    procedure ToolButton8Click(Sender: TObject);
    procedure DT1Change(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
    procedure ToolButton6Click(Sender: TObject);
    procedure FormMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure ToolButton10Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure N9Click(Sender: TObject);
    procedure EqpClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure Button4Click(Sender: TObject);
    procedure ScrollBox1MouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure TabSheet1Exit(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure TabSheet1Enter(Sender: TObject);
    procedure ToolButton12Click(Sender: TObject);
    procedure ApplicationEvents1Message(var Msg: tagMSG; var Handled: Boolean);
    procedure N10Click(Sender: TObject);
    procedure N11Click(Sender: TObject);
    procedure N12Click(Sender: TObject);
    procedure N13Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Sendemail1Click(Sender: TObject);
    procedure N15Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure StringGrid2DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure ToolButton13Click(Sender: TObject);
  private
    LogFile : TextFile;
    procedure AppendToLog(t : string);
  public
    { Public declarations }
  end;

const
  needMaxLine = True;

var
  Form1: TForm1;
  Holding : THolding;

  ChangeDate : TDate; // используется для предотвращения двойного срабатывания DT1.OnChange
  PathSi8 : string;
  RepPath : string;
  smena3 : record
    sm1beg : integer;
    sm1end : integer;
    sm2beg : integer;
    sm2end : integer;
    sm3beg : integer;
    sm3end : integer;
  end;
  smena2 : record
    sm1beg : integer;
    sm1end : integer;
    sm2beg : integer;
    sm2end : integer;
  end;
  ListExt : TStringList;
  TabSheetScrollPosition : integer;
  OldView : boolean;
  VisibleCex : integer;

implementation

uses reports, ToExcel, Unit2, Unit3;
{$SETPEFlAGS IMAGE_FILE_RELOCS_STRIPPED or IMAGE_FILE_DEBUG_STRIPPED or IMAGE_FILE_LINE_NUMS_STRIPPED or IMAGE_FILE_LOCAL_SYMS_STRIPPED}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$R *.dfm}



procedure TForm1.ApplicationEvents1Message(var Msg: tagMSG;
  var Handled: Boolean);
begin
  if Msg.Message=WM_MOUSEWHEEL then begin
    Form1.Perform(CM_MOUSEWHEEL, Msg.wParam, Msg.lParam);
    Handled := True;
  end;
end;

procedure ColumnWidthAlign(aSg: TStringGrid; aColNum : Longword; aDefaultColWidth : Integer = -1);
var
  RowNum      : Integer;
  ColWidth    : Integer;
  MaxColWidth : Integer;
begin
  if (aColNum < 0) or (aColNum > Pred(aSg.ColCount)) then Exit;
  if aDefaultColWidth < 0 then
    MaxColWidth := aSg.DefaultColWidth
  else
    MaxColWidth := aDefaultColWidth;
  for RowNum := 0 to Pred(aSg.RowCount) do begin
    ColWidth := aSg.Canvas.TextWidth(aSg.Cells[aColNum, RowNum]);
    if MaxColWidth < ColWidth then MaxColWidth := ColWidth;
  end;
  //+5 - потому что иногда текст всё же немного не умещается по ширине. :-)
  aSg.ColWidths[aColNum] := MaxColWidth + 5;
end;

procedure TForm1.Button1Click(Sender: TObject);   //  расчитать статистику за период
var
  i, j, k, m : Integer;
  month_rab_minute : integer;
  rab_sm_minute, sm_min : array[1..2] of integer;
  d : integer;
  kmv_all, kmv_rab : real;
  kmv_cex_all, kmv_cex_rab : real;
  kmv_zavod_all, kmv_zavod_rab : real;
  adr, l, x, z, y : integer;
  dir_name, fname, s : string;
  T : TDateTime;
begin
  // расчет за месяц
  for i := 1 to StringGrid1.ColCount - 1 do for j:= 1 to StringGrid1.RowCount - 1 do StringGrid1.Cells[i,j]:='';
  d:=DayOf(DT1.Date);
  m:=0;
  T:=Time;
  for i:=0 to Holding.Count-1 do begin
     inc(m);
     FillChar(Holding.ListFactory[i].kmv_rab_day, SizeOf(Holding.ListFactory[i].kmv_rab_day), 0);
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        FillChar(Holding.ListFactory[i].ListCex[j].kmv_rab_day, SizeOf(Holding.ListFactory[i].ListCex[j].kmv_rab_day), 0);
        inc(m);
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
          inc(m);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].log_month, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].log_month), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_day, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_day), 0);
        end;
     end;
  end;
  AppendToLog('Очистка массивов: '+FormatDateTime('h:n:s.z', Time - T));
  StringGrid1.RowCount := m + 1;
/// тут надо закачать файлы за период
  if ftp.enable=1 then begin
     idFTP1.Connect;
     if idFTP1.Connected then begin
       Formatsettings.ShortDateFormat:='ddmmyy';
       dir_name:=copy(DateToStr(DT1.Date),3,4);
       idFTP1.ChangeDir(dir_name);
       if not DirectoryExists('temp/'+dir_name) then CreateDir('temp/'+dir_name);
       for j:=1 to d do begin
         if j<10 then s:='0'+IntToStr(j) else s:=IntToStr(j);
         fname:=s+dir_name;
         for i:=0 to ListExt.Count-1 do begin
            s:='temp/'+dir_name+'/'+fname+'.'+ListExt[i];
            if not FileExists(s) or (idFTP1.Size(fname+'.'+ListExt[i])<>GetFileSize(s)) then begin
               try
                 idFTP1.Get(fname+'.'+ListExt[i],'temp/'+dir_name+'/'+fname+'.'+ListExt[i],True);
               except
                  on e: EIdException do begin
                   AppendToLog(' Не могу скачать '+fname+'.'+ListExt[i]);
                  end;
               end;
            end;
            s:='temp/'+dir_name+'/'+fname+'.lo'+copy(ListExt[i],3,1);
            if not FileExists(s)  or (idFTP1.Size(fname+'.'+ListExt[i])<>GetFileSize(s)) then begin
               try
                 idFTP1.Get(fname+'.lo'+copy(ListExt[i],3,1),s,True);
               except
                 on e: EIdException do begin
                   AppendToLog(' Не могу скачать '+fname+'.lo'+copy(ListExt[i],3,1));
                 end;
               end;
            end;
         end;
       end;
     end;
     idFTP1.Disconnect;
  end;
  T:=Time;
  for i:=0 to ListExt.Count - 1 do Holding.GetStatMonth(PathSi8, ListExt[i], Form1.DT1.Date);
  AppendToLog('GetStatMonth: '+FormatDateTime('h:n:s.z', Time - T));
  T:=Time;
  for i:=0 to Holding.Count-1 do begin
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
            Holding.ListFactory[i].ListCex[j].ListEquips[k].InterpolateMonth(DayOf(d));
        end;
     end;
  end;
  AppendToLog('Интерполирование: '+FormatDateTime('h:n:s.z', Time - T));
  T:=Time;
  for i:=0 to Holding.Count-1 do begin
    Holding.ListFactory[i].kmv:=0;
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        Holding.ListFactory[i].ListCex[j].kmv:=0;
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
            Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month:=0;
            Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_month:=0;
            Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_month:=0;
            month_rab_minute:=0;
            for l:=1 to d do begin
               rab_sm_minute[1]:=0;
               rab_sm_minute[2]:=0;
               sm_min[1]:=0;
               sm_min[2]:=0;
               for m:=0 to 1439 do begin
                  if m<=smena2.sm1end then begin
                     if Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month[l,m]>0 then inc(rab_sm_minute[1]);
                     inc(sm_min[1]);
                  end
                  else begin
                     if Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month[l,m]>0 then inc(rab_sm_minute[2]);
                     inc(sm_min[2]);
                  end;
               end;
               Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day[l]:=sum(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month[l]);
               Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_month:=Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_month+Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day[l];
               month_rab_minute:=month_rab_minute+rab_sm_minute[1]+rab_sm_minute[2];
               Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l]:=(rab_sm_minute[1]+rab_sm_minute[2])/(sm_min[1]+sm_min[2]);
               if rab_sm_minute[1]=0 then sm_min[1]:=0;
               if rab_sm_minute[2]=0 then sm_min[2]:=0;
               if ((sm_min[1]+sm_min[2])>0) then Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_day[l]:=(rab_sm_minute[1]+rab_sm_minute[2])/(sm_min[1]+sm_min[2]);
               Holding.ListFactory[i].ListCex[j].kmv_rab_day[l]:=Holding.ListFactory[i].ListCex[j].kmv_rab_day[l]+Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l];
               Holding.ListFactory[i].kmv_rab_day[l]:=Holding.ListFactory[i].kmv_rab_day[l]+Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l];
            end;
            Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month:=sum(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day);
            Holding.ListFactory[i].ListCex[j].kmv:=Holding.ListFactory[i].ListCex[j].kmv+Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month;
        end;
        Holding.ListFactory[i].kmv:=Holding.ListFactory[i].kmv+Holding.ListFactory[i].ListCex[j].kmv;
     end;
  end;
  AppendToLog('Расчет кмв: '+FormatDateTime('h:n:s.z', Time - T));
  y:=0;
  FormatSettings.DecimalSeparator:='.';
  T:=Time;
  for i:=0 to Holding.Count-1 do begin
     StringGrid1.Cells[0,y+1]:='ЗАВОД '+Holding.ListFactory[i].name;
     if Holding.ListFactory[i].KolObr<>0 then
        StringGrid1.Cells[2,y+1]:=Format('%4.2f', [Holding.ListFactory[i].kmv/Holding.ListFactory[i].KolObr/d]);
     if Holding.ListFactory[i].KolObr<>0 then
        for l:=1 to d do StringGrid1.Cells[3+l,y+1]:=Format('%4.2f', [Holding.ListFactory[i].kmv_rab_day[l]/Holding.ListFactory[i].KolObr]);
     inc(y);
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        StringGrid1.Cells[0,y+1]:='  '+Holding.ListFactory[i].ListCex[j].name;
        if Holding.ListFactory[i].ListCex[j].Count<>0 then StringGrid1.Cells[2,y+1]:=Format('%4.2f', [Holding.ListFactory[i].ListCex[j].kmv/Holding.ListFactory[i].ListCex[j].Count/d]);
        if Holding.ListFactory[i].ListCex[j].Count<>0 then
           for l:=1 to d do StringGrid1.Cells[3+l,y+1]:=Format('%4.2f', [Holding.ListFactory[i].ListCex[j].kmv_rab_day[l]/Holding.ListFactory[i].ListCex[j].Count]);
        inc(y);
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
            StringGrid1.Cells[0,y+1]:='    '+Holding.ListFactory[i].ListCex[j].ListEquips[k].info;
            StringGrid1.Cells[3,y+1]:=Format('%10.3f', [Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_month/1000]);
            StringGrid1.Cells[2,y+1]:=Format('%4.2f', [Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month/d]);
            for l:=1 to d do
               if Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day[l]>0 then
                  StringGrid1.Cells[3+l,y+1]:=Format('%10.3f', [Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day[l] / 1000])+' (кмв '+
                                              Format('%4.2f', [Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l]])+')';
            inc(y);
        end;
     end;
  end;
  AppendToLog('Вывод в grid: '+FormatDateTime('h:n:s.z', Time - T));
  for i := 0 to StringGrid1.ColCount - 1 do ColumnWidthAlign(StringGrid1, i);
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  XLApp, Sheet: OLEVariant;
  i, j, m : Integer;
  month_rab_minute : integer;
  month_sm_minute : integer;
  rab_sm_minute, sm_min : array[1..2] of integer;
  d : integer;
  Zavod : integer;
  cex, ind_cex, count_obr_cex : integer;
  _zavod, ind_zavod, count_obr_zavod : integer;
  kmv_cex_all, kmv_cex_rab : real;
  kmv_zavod_all, kmv_zavod_rab : real;
  l, adr, x : integer;
  id_zavod, id_cex :integer;
  kol_obr_zavod : integer;
  kol_obr_cex : integer;
  Range, Cell1, Cell2 : Variant;
  s, dir_name, fname : string;
  SearchRec: TSearchRec;
  y, k, row_itog : integer;
  Itog : array[1..31] of integer;
begin
  // расчет за месяц кмв и люди
  for i:=1 to StringGrid2.RowCount-1 do StringGrid2.Rows[i].Clear;
  d:=DayOf(DT1.Date);
  m:=0;
  for i:=0 to Holding.Count-1 do begin
     inc(m);
     FillChar(Holding.ListFactory[i].kmv_rab_day, SizeOf(Holding.ListFactory[i].kmv_rab_day), 0);
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        FillChar(Holding.ListFactory[i].ListCex[j].kmv_rab_day, SizeOf(Holding.ListFactory[i].ListCex[j].kmv_rab_day), 0);
        inc(m);
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
          inc(m);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].log_month, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].log_month), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_day), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_day, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_day), 0);
        end;
     end;
  end;
//  StringGrid2.RowCount := m - 1;

/// тут надо закачать файлы за период
  if ftp.enable=1 then begin
     StatusBar1.Panels[0].Text:='Скачиваем данные с FTP-сервера';
     Progressbar1.Min:=0;
     ProgressBar1.Max:=d;
     idFTP1.Connect;
     if idFTP1.Connected then begin
       Formatsettings.ShortDateFormat:='ddmmyy';
       dir_name:=copy(DateToStr(DT1.Date),3,4);
       idFTP1.ChangeDir(dir_name);
       if not DirectoryExists('temp/'+dir_name) then CreateDir('temp/'+dir_name);
       for j:=1 to d do begin
         ProgressBar1.Position:=ProgressBar1.Position+1;
         if j<10 then s:='0'+IntToStr(j) else s:=IntToStr(j);
         fname:=s+dir_name;
         for i:=0 to ListExt.Count-1 do begin
            s:='temp/'+dir_name+'/'+fname+'.'+ListExt[i];
            if not FileExists(s) or (idFTP1.Size(fname+'.'+ListExt[i])<>GetFileSize(s)) then begin
               try
                 idFTP1.Get(fname+'.'+ListExt[i],'temp/'+dir_name+'/'+fname+'.'+ListExt[i],True);
               except
                  on e: EIdException do begin
                   AppendToLog(' Не могу скачать '+fname+'.'+ListExt[i]);
                  end;
               end;
            end;
            s:='temp/'+dir_name+'/'+fname+'.lo'+copy(ListExt[i],3,1);
            if not FileExists(s)  or (idFTP1.Size(fname+'.'+ListExt[i])<>GetFileSize(s)) then begin
               try
                 idFTP1.Get(fname+'.lo'+copy(ListExt[i],3,1),s,True);
               except
                 on e: EIdException do begin
                   AppendToLog(' Не могу скачать '+fname+'.lo'+copy(ListExt[i],3,1));
                 end;
               end;
            end;
         end;
       end;
     end;
     idFTP1.Disconnect;
  end;
///
  Form3.Progressbar1.Min:=0;
  Form3.ProgressBar1.Max:=(ListExt.Count-1)*4+4;
  Form3.ProgressBar1.Position:=0;
  Form3.Memo1.Lines.Clear;
  Form3.Show;

  StatusBar1.Panels[0].Text:='Рассчитываем статистику за период';
  Form3.Caption:='Подождите. Расчет КМВ за период';
  Form3.Memo1.Lines.Add('Загружаем данные по работе оборудования...');
  for i:=0 to ListExt.Count - 1 do begin
//     ProgressBar1.Position:=ProgressBar1.Position+1;
     Form3.ProgressBar1.Position:=Form3.ProgressBar1.Position+1;
     Holding.GetStatMonth(PathSi8, ListExt[i], Form1.DT1.Date);
  end;
  Form3.Memo1.Lines.Add('Интерполяция данных по работе оборудования...');
  for i:=0 to Holding.Count-1 do begin
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
            Holding.ListFactory[i].ListCex[j].ListEquips[k].InterpolateMonth(DayOf(d));
        end;
     end;
     Form3.ProgressBar1.Position:=Form3.ProgressBar1.Position+1;
  end;
  y:=0;
  Form3.Memo1.Lines.Add('Рассчитываем КМВ...');
  for i:=0 to Holding.Count-1 do begin
    Holding.ListFactory[i].kmv:=0;
    inc(y);
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        Holding.ListFactory[i].ListCex[j].kmv:=0;
        inc(y);
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
            inc(y);
            Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month:=0;
            Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_month:=0;
            Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_month:=0;
            month_rab_minute:=0;
//            month_sm_minute:=0;
            for l:=1 to d do begin
               rab_sm_minute[1]:=0;
               rab_sm_minute[2]:=0;
               sm_min[1]:=0;
               sm_min[2]:=0;
               for m:=0 to smena2.sm1end do begin
                   if Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month[l,m]>0 then inc(rab_sm_minute[1]);
                   inc(sm_min[1]);
               end;
               for m:=smena2.sm2beg to 1439 do begin
                   if Holding.ListFactory[i].ListCex[j].ListEquips[k].speed_month[l,m]>0 then inc(rab_sm_minute[2]);
                   inc(sm_min[2]);
               end;
               month_rab_minute:=month_rab_minute+rab_sm_minute[1]+rab_sm_minute[2];
//               if rab_sm_minute[1]>0 then month_sm_minute:=month_sm_minute+sm_min[1];
//               if rab_sm_minute[2]>0 then month_sm_minute:=month_sm_minute+sm_min[2];
               Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l]:=(rab_sm_minute[1]+rab_sm_minute[2])/(sm_min[1]+sm_min[2]);
               if rab_sm_minute[1]=0 then sm_min[1]:=0;
               if rab_sm_minute[2]=0 then sm_min[2]:=0;
               if ((sm_min[1]+sm_min[2])>0) then Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_day[l]:=(rab_sm_minute[1]+rab_sm_minute[2])/(sm_min[1]+sm_min[2]);
               Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month:=Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month+Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l];
               Holding.ListFactory[i].ListCex[j].kmv_rab_day[l]:=Holding.ListFactory[i].ListCex[j].kmv_rab_day[l]+Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l];
               Holding.ListFactory[i].kmv_rab_day[l]:=Holding.ListFactory[i].kmv_rab_day[l]+Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l];
            end;
            Holding.ListFactory[i].ListCex[j].kmv:=Holding.ListFactory[i].ListCex[j].kmv+Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month;
        end;
        Holding.ListFactory[i].kmv:=Holding.ListFactory[i].kmv+Holding.ListFactory[i].ListCex[j].kmv;
     end;
     Form3.ProgressBar1.Position:=Form3.ProgressBar1.Position+1;
  end;
  Holding.GetNumEmp(PathSi8, DT1.Date);
  StringGrid2.RowCount:=y+Holding.Count+5;
  y:=1;
  FormatSettings.DecimalSeparator:='.';
  Form3.Memo1.Lines.Add('Выводим в таблицу...');
  for i:=0 to Holding.Count-1 do begin
     for j:=1 to 31 do Itog[j]:=0;
//     FillChar(itog, Length(Itog), 0);
     row_itog:=y+2;
     StringGrid2.Cells[0,y+1]:='ЗАВОД '+Holding.ListFactory[i].name+' ('+IntToStr(Holding.ListFactory[i].KolObr)+' ед.обор.)';
     StringGrid2.Cells[0,y+2]:='        в том числе рабочих';
     if Holding.ListFactory[i].KolObr<>0 then begin
        StringGrid2.Cells[1,y+1]:=Format('%4.2f', [Holding.ListFactory[i].kmv/Holding.ListFactory[i].KolObr/d]);
        for l:=1 to d do begin
          StringGrid2.Cells[l*2,y+1]:=Format('%4.2f', [Holding.ListFactory[i].kmv_rab_day[l]/Holding.ListFactory[i].KolObr]);
          StringGrid2.Cells[l*2+1,y+1]:=IntToStr(Holding.ListFactory[i].data_json[l]);
        end;
     end;
     y:=y+2;
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        StringGrid2.Cells[0,y+1]:='  '+Holding.ListFactory[i].ListCex[j].name+' ('+IntToStr(Holding.ListFactory[i].ListCex[j].Count)+' ед.обор.)';
        if Holding.ListFactory[i].ListCex[j].Count<>0 then begin
           StringGrid2.Cells[1,y+1]:=Format('%4.2f', [Holding.ListFactory[i].ListCex[j].kmv/Holding.ListFactory[i].ListCex[j].Count/d]);
           for l:=1 to d do begin
              StringGrid2.Cells[l*2,y+1]:=Format('%4.2f', [Holding.ListFactory[i].ListCex[j].kmv_rab_day[l]/Holding.ListFactory[i].ListCex[j].Count]);
              StringGrid2.Cells[l*2+1,y+1]:=IntToStr(Holding.ListFactory[i].ListCex[j].data_json[l]);
              Itog[l]:=Itog[l]+Holding.ListFactory[i].ListCex[j].data_json[l];
           end;
        end;
        inc(y);
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
            StringGrid2.Cells[0,y+1]:='    '+Holding.ListFactory[i].ListCex[j].ListEquips[k].info;
            StringGrid2.Cells[1,y+1]:=Format('%4.2f', [Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_month/d]);
            for l:=1 to d do StringGrid2.Cells[l*2,y+1]:=Format('%4.2f', [Holding.ListFactory[i].ListCex[j].ListEquips[k].kmv_rab_day[l]]);
            inc(y);
        end;
     end;
     for l:=1 to d do StringGrid2.Cells[l*2+1,row_itog]:=IntToStr(Itog[l]);
     Form3.ProgressBar1.Position:=Form3.ProgressBar1.Position+1;
  end;
  for i:=0 to StringGrid2.ColCount - 1 do ColumnWidthAlign(StringGrid2, i);
  StatusBar1.Panels[0].Text:='';
  Form3.Hide;
end;


procedure TForm1.Button3Click(Sender: TObject);
begin
   if SaveAsExcelFile(StringGrid1, 'My Stringgrid Data', 'c:\MyExcelFile.xls') then ShowMessage('StringGrid saved!');
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
   if SaveAsCSV(StringGrid1, 'My Stringgrid Data', 'c:\MyExcelFile.xls') then ShowMessage('StringGrid saved!');
end;

procedure TForm1.Button5Click(Sender: TObject);
var
  XLApp, Sheet: OLEVariant;
  Range{, Cell1, Cell2} : Variant;
  d, i, m, j : integer;
//  i: TObject;
  Vals : Variant;
  day_start : integer;
begin
  Form3.Progressbar1.Min:=0;
  Form3.ProgressBar1.Max:=StringGrid2.RowCount;
  Form3.ProgressBar1.Position:=0;
  Form3.Memo1.Lines.Clear;
  Form3.Show;
  Form3.Caption:='Выгружаем в Excel...';
  d:=DayOf(DT1.Date);
//  экспорт в excel
   XLApp := CreateOleObject('Excel.Application');

   // Add new Workbook
   XLApp.Workbooks.Add(xlWBatWorkSheet);
   Sheet := XLApp.Workbooks[1].WorkSheets[1];
   Sheet.Name := 'Отчет по кмв';
//   тут нужно по другому выгружать
  Form3.Memo1.Lines.Add('Формируем заголовок документа...');
// рисуем сетку по все книге
   XLApp.Range[Sheet.Cells[2, 1], Sheet.Cells[StringGrid2.RowCount-1, d*2+2]].Select;
   XLApp.Selection.Borders.LineStyle:=xlContinuous;
   XLApp.Selection.Borders.Weight:=xlThin;
   XLApp.Selection.HorizontalAlignment:=xlCenter;
   XLApp.Selection.Borders[xlEdgeBottom].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeBottom].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeLeft].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeLeft].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeRight].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeRight].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeTop].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeTop].Weight:=xlMedium;

   Sheet.Cells[1,1].Value:='Отчет по работе оборудования';


   Sheet.Cells[2,3].Value:=GetNameMonth(MonthOf(DT1.Date))+' '+IntToStr(YearOf(DT1.Date))+' г.';
// объединяем ячейки с названием отчета
   XLApp.Range[Sheet.Cells[1, 1], Sheet.Cells[1, d*2+2]].Select;
   XLApp.Selection.mergeCells:=True;
   XLApp.Selection.HorizontalAlignment:=xlCenter;
//
   XLApp.Range[Sheet.Cells[2, 1], Sheet.Cells[4, 1]].mergeCells:=True;
// объединяем ячейки с надписью КМВ
   XLApp.Range[Sheet.Cells[2, 2], Sheet.Cells[4, 2]].mergeCells:=True;
// объединяем ячейки с названием месяца
   XLApp.Range[Sheet.Cells[2, 3], Sheet.Cells[2, d*2+2]].mergeCells:=True;

   for i:=1 to d do begin
     Sheet.Cells[3,i*2+1].Value:=IntToStr(i);
     XLApp.Range[Sheet.Cells[3, i*2+1], Sheet.Cells[3, i*2+2]].mergeCells:=True;
     Sheet.Cells[4,i*2+1].Value:='кмв';
     Sheet.Cells[4,i*2+2].Value:='числ';
   end;
   m:=5;
// жирная линия вокруг ячйеки месяц год
   XLApp.Range[Sheet.Cells[2, 3], Sheet.Cells[2, d*2+2]].Select;
   XLApp.Selection.Borders[xlEdgeBottom].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeBottom].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeLeft].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeLeft].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeRight].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeRight].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeTop].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeTop].Weight:=xlMedium;
//  жирная линия
   Sheet.Cells[2,2].Value:='КМВ нараст';
   XLApp.Range[Sheet.Cells[2, 2], Sheet.Cells[3, 2]].Select;
   XLApp.Selection.mergeCells:=True;
   XLApp.Selection.WrapText:=True;
   XLApp.Selection.VerticalAlignment:=3;
   XLApp.Selection.Borders[xlEdgeLeft].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeLeft].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeRight].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeRight].Weight:=xlMedium;
//  жирная линия вокруг заголовка
   XLApp.Range[Sheet.Cells[2, 1], Sheet.Cells[4, d*2+2]].Select;
   XLApp.Selection.Borders[xlEdgeBottom].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeBottom].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeLeft].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeLeft].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeRight].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeRight].Weight:=xlMedium;
   XLApp.Selection.Borders[xlEdgeTop].LineStyle:=xlContinuous;
   XLApp.Selection.Borders[xlEdgeTop].Weight:=xlMedium;
   Vals:=VarArrayCreate([0, d*2+2], varVariant);
//   Progressbar1.Min:=0;
//   ProgressBar1.Max:=(StringGrid2.RowCount-1)*(d*2+3);
   Form3.ProgressBar1.Position:=Form3.ProgressBar1.Position+1;
   Form3.Memo1.Lines.Add('Выводим данные...');
 {  if CheckBox1.Checked then begin
      if d<=5 then day_start:=0
      else day_start:=10
   end
   else day_start:=0;   }

   for i:=2 to StringGrid2.RowCount-1 do begin
     Form3.Memo1.Lines.Add(StringGrid2.Cells[0,i]);
     Form3.ProgressBar1.Position:=Form3.ProgressBar1.Position+1;
     for j:=0 to d*2+2 do begin
          Vals[j]:=StringGrid2.Cells[j,i];
          if pos('ЗАВОД', StringGrid2.Cells[j,i])<>0 then begin
             XLApp.Range[Sheet.Cells[m, 1], Sheet.Cells[m, d*2+2]].Select;
             XLApp.Selection.Interior.Color:=rgb(255,192,0);
             XLApp.Selection.Font.Bold:=True;
             XLApp.Selection.Borders[xlEdgeTop].LineStyle:=xlContinuous;
             XLApp.Selection.Borders[xlEdgeTop].Weight:=xlMedium;
          end;
          if pos('Цех', StringGrid2.Cells[j,i])<>0 then begin
             XLApp.Range[Sheet.Cells[m, 1], Sheet.Cells[m, d*2+2]].Select;
             XLApp.Selection.Interior.Color:=rgb(255,230,153);
             XLApp.Selection.Font.Bold:=True;
          end;
          if (not Odd(j+1)) or (j=0) then begin
             Sheet.Cells[m,j+1].Borders[xlEdgeRight].LineStyle:=xlContinuous;
             Sheet.Cells[m,j+1].Borders[xlEdgeRight].Weight:=xlMedium;
            end;
//            ProgressBar1.Position:=ProgressBar1.Position+1;
     end;
     XLApp.Range[Sheet.Cells[m,1], Sheet.Cells[m, d*2+3]].Value:=Vals;
     inc(m);
   end;
   Range:=Sheet.Columns;
   Range.Columns[1].ColumnWidth := 50;
   Range.Columns[1].HorizontalAlignment:=1;
   Range.Columns[2].ColumnWidth := 9;
   for i:= 3 to d*2+2 do Range.Columns[i].ColumnWidth:=5.14;
   Vals:=Unassigned;
//       Columns("BD:BG").Select
//    Selection.EntireColumn.Hidden = True
//   XLApp.Range[Sheet.Cells.Columns[5],Sheet.Cells.Columns[{d*2-3]}10]].Select;
//   XLApp.Selection.EntireColumn.Hidden:=True;
   Form3.Hide;
   XLApp.Visible := True;
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
var i, j, k, h : integer;
begin
   ScrollBox1.VertScrollBar.Position:=0;
   case ComboBox1.ItemIndex of
      0 : VisibleCex:=0;
      1 : VisibleCex:=11;
      2 : VisibleCex:=12;
      3 : VisibleCex:=13;
      4 : VisibleCex:=14;
      5 : VisibleCex:=21;
      6 : VisibleCex:=22;
      7 : VisibleCex:=23;
      8 : VisibleCex:=24;
      9 : VisibleCex:=31;
      10 : VisibleCex:=41;
   end;
   h:=0;
   for j:=0 to Holding.Count-1 do begin
       if (VisibleCex in Holding.ListFactory[j].IdCexs) or (VisibleCex=0) then begin
          Holding.ListFactory[j].ButCaption.Visible:=True;
          Holding.ListFactory[j].ButCaption.Top:=h;
          h:=h+30;
       end
       else Holding.ListFactory[j].ButCaption.Visible:=False;
       for i:=0 to Holding.ListFactory[j].count-1 do begin
           if (Holding.ListFactory[j].ListCex[i].id=VisibleCex) or (VisibleCex=0) then begin
              Holding.ListFactory[j].ListCex[i].ButCaption.Visible:=True;
              Holding.ListFactory[j].ListCex[i].ButCaption.Top:=h;
              h:=h+30;
           end
           else Holding.ListFactory[j].ListCex[i].ButCaption.Visible:=False;
           for k:=0 to Holding.ListFactory[j].ListCex[i].count-1 do begin
              if (VisibleCex=0) or (Holding.ListFactory[j].ListCex[i].id=VisibleCex) then begin
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].ButCaption.Visible:=True;
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].NewChart.Visible:=True;
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].grd.Visible:=True;
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].ButCaption.Top:=h;
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].NewChart.Top:=h;
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].grd.Top:=h;
                 h:=h+160;
              end
              else begin
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].ButCaption.Visible:=False;
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].NewChart.Visible:=False;
                 Holding.ListFactory[j].ListCex[i].ListEquips[k].grd.Visible:=False;
              end;
           end;
       end;
   end;
   ScrollBox1.SetFocus;
end;

procedure TForm1.GetSpeedAllAndDrawGraph;
var i, j, k : integer;
    fname, dir_name, s : string;
begin
    for i:=0 to Holding.Count-1 do begin
     FillChar(Holding.ListFactory[i].data_json, SizeOf(Holding.ListFactory[i].data_json), 0);
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        FillChar(Holding.ListFactory[i].ListCex[j].data_json, SizeOf(Holding.ListFactory[i].ListCex[j].data_json), 0);
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed), 0);
          FillChar(Holding.ListFactory[i].ListCex[j].ListEquips[k].log, SizeOf(Holding.ListFactory[i].ListCex[j].ListEquips[k].log), 0);
          Holding.ListFactory[i].ListCex[j].ListEquips[k].kol_sd := 0;
        end;
     end;
  end;

//  если ftp, то сначала скачать нужные файлы, иначе сразу считывать данные
  if ftp.enable=1 then begin
     idFTP1.Connect;
     if idFTP1.Connected then begin
         Formatsettings.ShortDateFormat:='ddmmyy';
         fname:=DateToStr(DT1.Date);
         dir_name:=copy(fname,3,4);
         idFTP1.ChangeDir(dir_name);
         if not DirectoryExists('temp/'+dir_name) then CreateDir('temp/'+dir_name);
         for i:=0 to ListExt.Count-1 do begin
            s:='temp/'+dir_name+'/'+fname+'.'+ListExt[i];
            if not FileExists(s) or (idFTP1.Size(fname+'.'+ListExt[i])<>GetFileSize(s)) then begin
//            if (not FileExists('temp/'+dir_name+'/'+fname+'.'+ListExt[i])) or ((Date-DT1.Date)<1) then begin
               try
                 idFTP1.Get(fname+'.'+ListExt[i],'temp/'+dir_name+'/'+fname+'.'+ListExt[i],True);
               except
                   AppendToLog(' Не могу скачать '+fname+'.'+ListExt[i]);
               end;
            end;
            s:='temp/'+dir_name+'/'+fname+'.lo'+copy(ListExt[i],3,1);
            if not FileExists(s)  or (idFTP1.Size(fname+'.'+ListExt[i])<>GetFileSize(s)) then begin
//            if not FileExists(s) or ((Date-DT1.Date)<1) then begin
               try
                 idFTP1.Get(fname+'.lo'+copy(ListExt[i],3,1),s,True);
               except
                   AppendToLog(' Не могу скачать '+fname+'.lo'+copy(ListExt[i],3,1));
               end;
            end;
         end;
     end;
     idFTP1.Disconnect;
  end;
  for i:=0 to ListExt.Count - 1 do Holding.GetSpeed(PathSi8, ListExt[i], Form1.DT1.Date);
//  ListEquips.GetEnergy(PathSi8, Form1.DT1.Date);
 { ListEquips.GetNumEmp(PathSi8, Form1.DT1.Date);}     /// надо под новое
  DrawHoldingGraph
end;

procedure TForm1.DrawHoldingGraph;
var
  fost, fost2: Boolean;
  i, j, k, x, minute, minute2, ost, ost2, curminute: Integer;
  mspd, mspd2, srspd, srspd2, kol, kol2: single;
  d, t: TDate;
begin
  d := Date();
  t := time();
  if ((DayOf(d)=DayOf(DT1.Date)) and (MonthOf(d)=MonthOf(DT1.Date)) and (YearOf(d)=YearOf(DT1.Date))) then
    curminute := HourOf(t) * 60 + MinuteOf(t)
  else
    curminute := 1439;
   for i:=0 to Holding.Count-1 do begin
      for j:=0 to Holding.ListFactory[i].Count-1 do begin
         for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
             Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.DrawGraph(Holding.ListFactory[i].ListCex[j].ListEquips[k].speed,Holding.ListFactory[i].ListCex[j].ListEquips[k].log,curminute,needMin,needInter,CorrectError,Holding.ListFactory[i].ListCex[j].ListEquips[k].minspd,Holding.ListFactory[i].ListCex[j].ListEquips[k].maxspd);
             minute := 0;
             minute2 := 0;
             mspd := 0;
             mspd2 := 0;
             kol := 0;
             kol2 := 0;
             ost := 0;
             ost2 := 0;
             fost := False;
             fost2 := False;
             for x := 0 to curminute do begin
               if Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.speed[x] > 0 then begin
                  if x < Smena2.sm2beg then begin
                    inc(minute);
                    kol := kol + Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.speed[x];
                    if Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.speed[x]>mspd then mspd:=Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.speed[x];
                    fost := True;
                  end
                  else begin
                    inc(minute2);
                    kol2 := kol2 + Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.speed[x];
                    if Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.speed[x]>mspd2 then mspd2:=Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.speed[x];
                    fost2 := True;
                  end;
               end
               else begin
                  if i < Smena2.sm2beg then begin
                     if fost then begin
                         if i<>curminute then inc(ost);
                         fost := False;
                     end;
                  end
                  else begin
                     if fost2 then begin
                        if i<>curminute then inc(ost2);
                        fost2 := False;
                     end;
                  end;
               end;
             end;
             if minute > 0 then srspd := RoundTo(kol / minute, -1) else srspd := 0;
             if minute2 > 0 then srspd2 := RoundTo(kol2 / minute2, -1) else srspd2 := 0;
             FormatSettings.DecimalSeparator:=',';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1, 1] := '      ' + AbsMinuteToStr(minute);
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2, 1] := '      ' + AbsMinuteToStr(minute2);
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3, 1] := '      ' + AbsMinuteToStr(minute+minute2);
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1, 2] := '      ' + FloatToStr(RoundTo(minute / (smena2.sm1end-smena2.sm1beg+1){720}{690}, -2));
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2, 2] := '      ' + FloatToStr(RoundTo(minute2 / (smena2.sm2end-smena2.sm2beg+1){720}{690}, -2));
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3, 2] := '      ' + FloatToStr(RoundTo((minute + minute2) / 1440, -2));
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1, 3] := '' + FloatToStr(RoundTo(mspd, -1)) + ' м/мин';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2, 3] := '' + FloatToStr(RoundTo(mspd2, -1)) + ' м/мин';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3, 3] := '' + FloatToStr(RoundTo(ifthen(mspd > mspd2, mspd, mspd2), -1)) + ' м/мин';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1, 4] := '' + FloatToStr(RoundTo(srspd, -1)) + ' м/мин';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2, 4] := '' + FloatToStr(RoundTo(srspd2, -1)) + ' м/мин';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3, 4] := '' + FloatToStr(RoundTo((srspd + srspd2) / 2, -1)) + ' м/мин';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1, 5] := ' ' + FloatToStr(RoundTo(kol / 1000, -1)) + ' км';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2, 5] := ' ' + FloatToStr(RoundTo(kol2 / 1000, -1)) + ' км';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3, 5] := ' ' + FloatToStr(RoundTo((kol + kol2) / 1000, -1)) + ' км';
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1, 6] := '        ' + IntToStr(ost);
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2, 6] := '        ' + IntToStr(ost2);
             Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3, 6] := '        ' + IntToStr(ost + ost2);
         end;
      end;
   end;
end;

procedure TForm1.DT1Change(Sender: TObject);
begin
  if DT1.Date<>ChangeDate then begin
    ChangeDate:=DT1.Date;
    GetSpeedAllAndDrawGraph;
  end;
end;

procedure TForm1.EqpClick(Sender: TObject);
begin
  Rep.Caption:='Отчеты по '+TButton(Sender).Caption;
  Rep.si8adr:=TEquip(TButton(Sender).owner).si8adr;
  Rep.info:=TEquip(TButton(Sender).owner).info;
  Rep.ext:=TEquip(TButton(Sender).owner).ext;
  Rep.Show;
end;

procedure TForm1.AppendToLog(t : string);
begin
  Append(LogFile);
  writeln(LogFile,DateToStr(date)+' '+TimeToStr(time)+' '+t);
  closefile(LogFile);
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   Holding.ClearItems;
   Holding.Free;
   Holding:=Nil;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  i, j, k: Integer;
  PathToMDB : string;
  fini : TIniFile;
  h : integer;
begin
  VisibleCex:=0;
  AssignFile(LogFile,'equips.log');
  Rewrite(LogFile);
  closefile(LogFile);
  AppendToLog(' Старт.');
  AppendToLog('Читаем конфиг...');
//   читаем параметры с ini файла
  OldView:=False;
  ScrollBox1.VertScrollBar.Position:=0;
  TabSheetScrollPosition:=0;
  fini:=TIniFile.Create(ExtractFilePath(Application.ExeName)+'equips.ini');

//  PathSi8:=fini.ReadString('FileSi30Path','Si30Path','');         // путь к бинарным файлам с данными получаемыми со счетчиков
  RepPath:=fini.ReadString('ReportPath','RepPath','');            // путь, куда сохранять отчеты
  RestAddr:=fini.ReadString('Rest','RestAddr','');
  ftp.enable:=fini.ReadInteger('ftp','enable',0);
  ftp.user:=fini.ReadString('ftp','user','obr');
  ftp.pwd:=fini.ReadString('ftp','pwd','obr1');
  ftp.port:=fini.ReadInteger('ftp','port',0);
  ftp.server:=fini.ReadString('ftp','server','192.168.30.10');
  if ftp.enable=0 then begin
//      PathToMDB:=fini.ReadString('BDAccessPath','AccessPath','');
      PathSi8:=fini.ReadString('FileSi30Path','Si30Path','');
  end
  else begin
//      PathToMDB:=fini.ReadString('BDAccessPath','AccessPathLoc','');
      PathSi8:=fini.ReadString('FileSi30Path','Si30PathLoc','');
  end;
  BeginDay.Hour:=fini.ReadInteger('BeginDay','BeginDayHour',0);      // начало дня (начало первой смены) часы
  BeginDay.Minute:=fini.ReadInteger('BeginDay','BeginDayMinute',0);  //            минуты
  smena3.sm1beg:=fini.ReadInteger('smena1_3','sm1beg',0);         // начало 1 смены  \
  smena3.sm1end:=fini.ReadInteger('smena1_3','sm1end',0);         // конец 1 смены    \
  smena3.sm2beg:=fini.ReadInteger('smena1_3','sm2beg',0);         // начало 2 смены    -   работы в 3 смены
  smena3.sm2end:=fini.ReadInteger('smena1_3','sm2end',0);         // конец 2 смены     -
  smena3.sm3beg:=fini.ReadInteger('smena1_3','sm3beg',0);         // начало 3 смены   /
  smena3.sm3end:=fini.ReadInteger('smena1_3','sm3end',0);         // конец 3 смены   /

  smena2.sm1beg:=TimeToAbsMinute(fini.ReadString('smena1_2','sm1beg','00:00'))-BeginDay.Hour*60-BeginDay.Minute;
  smena2.sm1end:=TimeToAbsMinute(fini.ReadString('smena1_2','sm1end','00:00'))-BeginDay.Hour*60-BeginDay.Minute;
  smena2.sm2beg:=TimeToAbsMinute(fini.ReadString('smena1_2','sm2beg','00:00'))-BeginDay.Hour*60-BeginDay.Minute;
  smena2.sm2end:=TimeToAbsMinute(fini.ReadString('smena1_2','sm2end','00:00'))-BeginDay.Hour*60-BeginDay.Minute;
  if smena2.sm2end<0 then smena2.sm2end:=1440-smena2.sm2end;

  KoefInterMaxSpeed:=fini.ReadInteger('koef','KoefInterMaxSpeed',0);

  ListExt:=TStringList.Create;
  ListExt.Delimiter := ';';
  ListExt.DelimitedText:=fini.ReadString('list_si','list_ext','');  // список расширений файлов данных счетчиков, которые надо читать

  GraphColors.BackGround:=fini.ReadInteger('GraphColor','BackGround',$000000);
  GraphColors.Labels:=fini.ReadInteger('GraphColor','Labels',$FFFFFF);
  GraphColors.ScaleLines:=fini.ReadInteger('GraphColor','ScaleLines',$FFFFFF);
  GraphColors.Smena:=fini.ReadInteger('GraphColor','Smena',$00FFFF);
  GraphColors.GridLines:=fini.ReadInteger('GraphColor','GridLines',$00FFFF);
  GraphColors.LineWidth:=fini.ReadInteger('GraphColor','LineWidth',1);

  HightCaptionCex:=fini.ReadInteger('Hights','HightCaptionCex',30);
  HightGraph:=fini.ReadInteger('Hights','HightGraph',160);
  WidthButCaption:=fini.ReadInteger('Widths','WidthButCaption',200);
  WidthGrd:=fini.ReadInteger('Widths','WidthGrd',291);

  fini.Free;
  AppendToLog('конфиг прочитан.');
  AppendToLog('Загружаем список оборудования.');
  idFTP1.Username:=ftp.user;
  idFTP1.Password:=ftp.pwd;
  idFTP1.Port:=ftp.port;
  idFTP1.Host:=ftp.server;

//  загружаем список оборудования из базы в ListEquips (см.Equips.pas)
  if ftp.enable=1 then begin
//  проверяем наличие и соответствие даты время у локальной и на ftp db1.mdb
//  если одинаковые базы, то использвуем локальную. Если разные, то сначала скачиваем
     AppendToLog('Соединяемся с FTP...');
     idFTP1.Connect;
     if idFTP1.Connected then begin
       AppendToLog('Соединились с FTP...');
       AppendToLog('Меняем каталог');
       idFTP1.ChangeDir('dbdata');
       idFTP1.Get('factory.json','temp\factory.json',True);
       idFTP1.Get('equips.json','temp\equips.json',True);
       idFTP1.Disconnect;
     end;
  end;
  AppendToLog('db1.mdb скопирован.');

  Holding:=THolding.Create;
  if ftp.enable=1 then
     Holding.LoadData(PathSi8)
  else
     Holding.LoadData(PathSi8+'dbdata\');

  AppendToLog('Список оборудования загружен.');

  ChangeDate:=0.0;
  PageControl1.ActivePageIndex := 0;
  StringGrid1.ColWidths[0] := 200;
  StringGrid1.Height:=Form1.Height - 60;
  Button1.Top:=Form1.Height-52;
  Button3.Top:=Form1.Height-52;
  Button4.Top:=Form1.Height-52;
  needMin := True;
  needInter:=False;
  CorrectError:=True;
  SI8CodeError:=False;
  h:=0;
  for i:=0 to Holding.Count-1 do begin
     Holding.ListFactory[i].ButCaption.Caption := Holding.ListFactory[i].name+'(кол-во ед.обор.='+IntToStr(Holding.ListFactory[i].KolObr)+')';
     Holding.ListFactory[i].ButCaption.Font.Style := [fsBold];
     Holding.ListFactory[i].ButCaption.Font.Size := 16;
     Holding.ListFactory[i].ButCaption.Font.Name := 'Times New Roman';
     Holding.ListFactory[i].ButCaption.Top:=h;
     Holding.ListFactory[i].ButCaption.Height:=HightCaptionCex;
     Holding.ListFactory[i].ButCaption.Width:=ScrollBox1.ClientWidth;
     Holding.ListFactory[i].ButCaption.Left:=0;
     Holding.ListFactory[i].ButCaption.Parent := ScrollBox1;
     h:=h+30;
     for j:=0 to Holding.ListFactory[i].Count-1 do begin
        Holding.ListFactory[i].ListCex[j].ButCaption.Caption := Holding.ListFactory[i].ListCex[j].name+'(кол-во ед.обор.='+IntToStr(Holding.ListFactory[i].ListCex[j].Count)+')';
        Holding.ListFactory[i].ListCex[j].ButCaption.Font.Style := [fsBold];
        Holding.ListFactory[i].ListCex[j].ButCaption.Font.Size := 16;
        Holding.ListFactory[i].ListCex[j].ButCaption.Font.Name := 'Times New Roman';
        Holding.ListFactory[i].ListCex[j].ButCaption.Top:=h;
        Holding.ListFactory[i].ListCex[j].ButCaption.Height:=HightCaptionCex;
        Holding.ListFactory[i].ListCex[j].ButCaption.Width:=ScrollBox1.ClientWidth;
        Holding.ListFactory[i].ListCex[j].ButCaption.Left:=0;
        Holding.ListFactory[i].ListCex[j].ButCaption.Parent := ScrollBox1;
        h:=h+30;
        for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Font.Style := [fsBold];
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Font.Size := 12;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Tag:=j;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Caption := Holding.ListFactory[i].ListCex[j].ListEquips[k].info;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.OnClick:=EqpClick;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Top:=h;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Height:=HightGraph;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Width:=WidthButCaption;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Left:=0;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Parent:=ScrollBox1;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Height:=HightGraph;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Top:=h;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Left:=Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Left+Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Width+1;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Width:=ScrollBox1.ClientWidth-WidthButCaption-WidthGrd-2;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Parent := ScrollBox1;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Height:=HightGraph;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Width:=WidthGrd;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Top:=h;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Left:=Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Left+Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Width+1;
           Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Parent:=ScrollBox1;
           h:=h+160;
        end;
     end;
  end;
  DT1.DateTime := Date();
  GetSpeedAllAndDrawGraph;
end;

procedure TForm1.FormMouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
  if PageControl1.ActivePage=TabSheet2 then begin
    if WheelDelta < 0 then
      ScrollBox1.VertScrollBar.Position := ScrollBox1.VertScrollBar.Position + 32
    else
      if ScrollBox1.VertScrollBar.Position > 16 then ScrollBox1.VertScrollBar.Position := ScrollBox1.VertScrollBar.Position-32;
    TabSheetScrollPosition:=ScrollBox1.VertScrollBar.Position;
  end;
end;

procedure TForm1.FormResize(Sender: TObject);
var i, j, k : integer;
begin
  StringGrid1.Height:=Form1.Height - 160;
  StringGrid1.Width:=Form1.Width - 17;
  Button1.Top:=Form1.Height-158;
  Button3.Top:=Form1.Height-158;
  Button4.Top:=Form1.Height-158;
  if Holding<>nil then
    for i:=0 to Holding.Count-1 do begin
       Holding.ListFactory[i].ButCaption.Width:=ScrollBox1.ClientWidth-2;
       for j:=0 to Holding.ListFactory[i].Count-1 do begin
          Holding.ListFactory[i].ListCex[j].ButCaption.Width:=ScrollBox1.ClientWidth-2;
          for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
              Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Width:=
                            ScrollBox1.ClientWidth-
                            Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Width-
                            Holding.ListFactory[i].ListCex[j].ListEquips[k].ButCaption.Width-2;
              Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Left:=Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Left+Holding.ListFactory[i].ListCex[j].ListEquips[k].NewChart.Width+1;
          end;
       end;
    end;
end;


procedure TForm1.N10Click(Sender: TObject);
begin
  if N10.Checked then needMin:=True else needMin:=False;
  DrawHoldingGraph;
end;

procedure TForm1.N11Click(Sender: TObject);
begin
  if N11.Checked then needInter:=True else needInter:=False;
  DrawHoldingGraph;
end;

procedure TForm1.N7Click(Sender: TObject);
begin
  close;
end;

procedure TForm1.N9Click(Sender: TObject);
begin
// тут будут различные отчеты
end;

procedure TForm1.ScrollBox1MouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
begin
   Form1.Caption:=IntToStr(ScrollBox1.VertScrollBar.Position);
end;

procedure TForm1.Sendemail1Click(Sender: TObject);
var msg : TidMessage;
begin
    IdSMTP1.Username:='alexl06@inbox.ru';
    IdSMTP1.Password:='Celeron3Optic3';
    IdSMTP1.Connect;
    msg:=TIdMessage.Create(nil);
    msg.Body.Add('send via Delphi');
    msg.Subject:='from Delphi';
    msg.From.Address:='alexl06@inbox.ru';
    msg.From.Name:='Alexander';
    msg.Recipients.EMailAddresses:='alexl06@yandex.ru';
    msg.IsEncoded:=True;
    if IdSMTP1.Connected=True then begin
      IdSMTP1.Send(msg);
    end;
    msg.Free;
    IdSMTP1.Disconnect;
end;

procedure TForm1.StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var  txt : string;
begin
  txt:=StringGrid1.Cells[ACol,ARow];
  if copy(StringGrid1.Cells[0, ARow],4,3)='Цех' then begin
    StringGrid1.Canvas.Font.Style:=[fsBold];
    StringGrid1.Canvas.Brush.Color:=clSkyBlue;
    StringGrid1.Canvas.FillRect(Rect);
  end
  else
    if copy(StringGrid1.Cells[0, ARow],1,5)='ЗАВОД' then begin
      StringGrid1.Canvas.Font.Style:=[fsBold];
      StringGrid1.Canvas.Brush.Color:=clMoneyGreen;
      StringGrid1.Canvas.FillRect(Rect);
    end
    else begin
      StringGrid1.Canvas.Brush.Color:=clWhite;
      StringGrid1.Canvas.FillRect(Rect);
    end;
  if ACol>=2 then begin
     txt:=StringGrid1.Cells[ACol,ARow];
     StringGrid1.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfCenter,tfSingleLine])
  end
  else
     StringGrid1.Canvas.TextOut(Rect.Left, Rect.Top, StringGrid1.Cells[ACol, ARow]);
end;

procedure TForm1.StringGrid2DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var  txt : string;
     tw : integer;
begin
  txt:=StringGrid2.Cells[ACol,ARow];
  if (ACol=0) or Odd(ACol) then begin
     StringGrid2.Canvas.Pen.Width:=2;
     StringGrid2.Canvas.Pen.Color:=clBlack;
     StringGrid2.Canvas.MoveTo(Rect.Right, Rect.Top);
     StringGrid2.Canvas.LineTo(Rect.Right, Rect.Bottom);
  end;
  if (ARow=1) or ((ARow=0) and (ACol>1)) then begin
     StringGrid2.Canvas.Pen.Width:=2;
     StringGrid2.Canvas.Pen.Color:=clBlack;
     StringGrid2.Canvas.MoveTo(Rect.Left, Rect.Bottom);
     StringGrid2.Canvas.LineTo(Rect.Right, Rect.Bottom);
  end;
  if (ARow=0) and (ACol>1) then begin
     if Odd(ACol) then begin
        txt:=IntToStr(trunc(ACol/2));
        Rect.Left:=Rect.Left-StringGrid2.ColWidths[ACol];
        StringGrid2.Canvas.Brush.Color:=clBtnFace;
        StringGrid2.Canvas.FillRect(Rect);
        StringGrid2.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfCenter,tfSingleLine]);
     end
     else begin
      txt:=IntToStr(trunc(ACol/2));
      Rect.Right:=Rect.Right+StringGrid2.ColWidths[ACol];
      StringGrid2.Canvas.Brush.Color:=clBtnFace;
      StringGrid2.Canvas.FillRect(Rect);
      StringGrid2.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfCenter,tfSingleLine]);
     end;
  end;
  if (ARow=1) then begin     // объединяем ячейкку КМВ
    if ACol=1 then begin
      txt:='КМВ';
      Rect.Top:=Rect.Top-StringGrid2.DefaultRowHeight-1;
      StringGrid2.Canvas.Brush.Color:=clBtnFace;
      StringGrid2.Canvas.FillRect(Rect);
      tw:=StringGrid2.Canvas.TextWidth('КМВ');
      StringGrid2.Canvas.TextOut(Rect.Left+trunc((StringGrid2.ColWidths[ACol]-tw)/2), Rect.Bottom-trunc(StringGrid2.DefaultRowHeight*1.5),'КМВ');
      tw:=StringGrid2.Canvas.TextWidth('нараст');
      StringGrid2.Canvas.TextOut(Rect.Left+trunc((StringGrid2.ColWidths[ACol]-tw)/2), Rect.Bottom-StringGrid2.DefaultRowHeight,'нараст');
    end
    else begin
       if ACol>1 then begin
          if not Odd(ACol) then txt:='кмв' else txt:='числ';
          StringGrid2.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfCenter,tfSingleLine]);
       end;
    end;
  end;
  if ARow>1 then begin
      if copy(StringGrid2.Cells[0, ARow],4,3)='Цех' then begin
        StringGrid2.Canvas.Font.Style:=[fsBold];
        StringGrid2.Canvas.Brush.Color:=clSkyBlue;
      end
      else
        if copy(StringGrid2.Cells[0, ARow],1,5)='ЗАВОД' then begin
          StringGrid2.Canvas.Font.Style:=[fsBold];
          StringGrid2.Canvas.Brush.Color:=clMoneyGreen;
        end
        else StringGrid2.Canvas.Brush.Color:=clWhite;;
       StringGrid2.Canvas.FillRect(Rect);
       if ACol>0 then StringGrid2.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfCenter,tfSingleLine])
       else StringGrid2.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfLeft,tfSingleLine]);
  end;
end;

procedure TForm1.TabSheet1Enter(Sender: TObject);
begin
   ComboBox1.Enabled:=False;
end;

procedure TForm1.TabSheet1Exit(Sender: TObject);
begin
   ComboBox1.Enabled:=True;
end;

procedure TForm1.TabSheet1Show(Sender: TObject);
var i : byte;
begin
  StringGrid1.ColCount := 39;
  StringGrid1.Cells[1, 0] := 'Время';
  StringGrid1.Cells[2, 0] := 'КМВ';
  StringGrid1.Cells[3, 0] := 'Изг.за месяц';
  for i:=4 to 34 do begin
    StringGrid1.ColWidths[i] := 70;
    StringGrid1.Cells[i,0]:='     '+IntToStr(i-3);
  end;
  StringGrid1.ColWidths[3] := 0;
  StringGrid1.ColWidths[4] := 0;
  StringGrid1.ColWidths[5] := 0;
  StringGrid1.ColWidths[6] := 0;
end;

procedure TForm1.ToolButton10Click(Sender: TObject);
begin
  if needMin then needMin := False else needMin := True;
  GetSpeedAllAndDrawGraph;
end;

procedure TForm1.ToolButton12Click(Sender: TObject);
begin
  if needInter then needInter := False else needInter := True;
  GetSpeedAllAndDrawGraph;
end;

procedure TForm1.ToolButton13Click(Sender: TObject);
var Bmp : TBitmap;
    ChartCanvas : TCanvas;
//    Jpg : TJpegImage;
    i, j, k, l : integer;
    nfile : string;
    lPdf   : TPdfDocumentGDI;
    lPage  : TPdfPage;
    x, y, w, h : integer;
    dh : integer;
    sday, eday : integer;
    d : integer;
    w_one : integer;
    w_kmv : integer;
    w_kol : integer;
    width_l : integer;
begin
   d:=DayOf(DT1.Date);
   lPdf := TPdfDocumentGDI.Create;
   lPdf.ScreenLogPixels:=600;
   lPage := lPDF.AddPage;
   lPdf.VCLCanvas.Brush.Style:=bsClear;
   x:=100;
   y:=100;
   lPdf.VCLCanvas.Font.Size:=72;
   lPdf.VCLCanvas.Font.Name:='Courier New';
   lPdf.VCLCanvas.TextOut(x,50,'Данные по работе оборудования за '+DateToStr(DT1.Date));
   dh:=abs(lPdf.VCLCanvas.Font.Height)+20;
//   eday:=DayOf(DT1.Date);
//   sday:=sday-5;
//   if sday<=0 then sday:=1;
   lPdf.VCLCanvas.Font.Size:=58;
   w_one:=0;
   for i:=2 to StringGrid2.RowCount-1 do begin
       if lPdf.VCLCanvas.TextWidth(StringGrid2.Cells[0,i])>w_one then w_one:=lPdf.VCLCanvas.TextWidth(StringGrid2.Cells[0,i]);
   end;
// ------------- временно 5 дней
   d:=5;
// -----------------------------
   y:=y+dh;
   w_kmv:=lPdf.VCLCanvas.TextWidth('00,00');
   w_kol:=lPdf.VCLCanvas.TextWidth('0000');
   width_l:=w_one+(w_kmv+10)+2*(w_kmv+10)*d+50;
//   lPdf.VCLCanvas.MoveTo(x,y);
//   lPdf.VCLCanvas.LineTo(width_l,y);

   for i:=2 to StringGrid2.RowCount-1 do begin
      x:=100;
      lPdf.VCLCanvas.Font.Style:=lPdf.VCLCanvas.Font.Style-[fsBold];
      if (pos('ЗАВОД', StringGrid2.Cells[0,i])<>0) or (pos('Цех', StringGrid2.Cells[0,i])<>0) then begin
         lPdf.VCLCanvas.Brush.Color:=rgb(255,192,0);
         lPdf.VCLCanvas.Font.Style:=lPdf.VCLCanvas.Font.Style+[fsBold];
      end;
      lPdf.VCLCanvas.Rectangle(x-10,y,x+width_l+20,y+dh);
      lPdf.VCLCanvas.TextOut(x,y,StringGrid2.Cells[0,i]);
      lPdf.VCLCanvas.Brush.Style:=bsClear;
      x:=x+w_one+50;
      for j:=1 to d*2+1 do begin
         lPdf.VCLCanvas.TextOut(x,y,PADL(StringGrid2.Cells[j,i],5));
         x:=x+w_kmv+10;
      end;
      y:=y+dh;
      if y>(lPdf.VCLCanvasSize.cy-500) then begin
         lPage := lPDF.AddPage;
         lPdf.VCLCanvas.Font.Name:='Courier New';
         lPdf.VCLCanvas.Font.Size:=58;
         y:=100;
         x:=100;
      end;
   end;
   lPdf.SaveToFile('1.pdf');
   lPdf.Free;
   ShellExecute(Handle, 'open', '1.pdf', nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.ToolButton1Click(Sender: TObject);
var
  f: textfile;
  j, i, k: Integer;
  c : string;
begin
  AssignFile(f,'otchet.txt');
  rewrite(f);
  c:=FormatSettings.ShortDateFormat;
  FormatSettings.ShortDateFormat:='dd.mm.yyyy';
   for i:=0 to Holding.Count-1 do begin
      for j:=0 to Holding.ListFactory[i].Count-1 do begin
         for k:=0 to Holding.ListFactory[i].ListCex[j].Count-1 do begin
            Write(f, StringReplace(Holding.ListFactory[i].ListCex[j].ListEquips[k].info,'"','',[rfReplaceAll, rfIgnoreCase])+';');
            Write(f,DateToStr(DT1.Date)+';');
            Write(f,TimeStr(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1,1])+';');  //время
            Write(f,'Смена 1;');
            Write(f,trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1,2])+';');
            Write(f,copy(trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1,4]),1,pos('м',trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1,4]))-2)+';');
            Writeln(f,copy(trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1,5]),1,pos('к',trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1,5]))-2));

            Write(f, StringReplace(Holding.ListFactory[i].ListCex[j].ListEquips[k].info,'"','',[rfReplaceAll, rfIgnoreCase])+';');
            Write(f,DateToStr(DT1.Date)+';');
            Write(f,TimeStr(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2,1])+';');  //время
            Write(f,'Смена 2;');
            Write(f,trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[1,2])+';');
            Write(f,copy(trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2,4]),1,pos('м',trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2,4]))-2)+';');
            Writeln(f,copy(trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2,5]),1,pos('к',trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[2,5]))-2));

            Write(f, StringReplace(Holding.ListFactory[i].ListCex[j].ListEquips[k].info,'"','',[rfReplaceAll, rfIgnoreCase])+';');
            Write(f,DateToStr(DT1.Date)+';');
            Write(f,TimeStr(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3,1])+';');  //время
            Write(f,'Сутки;');
            Write(f,trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3,2])+';');
            Write(f,copy(trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3,4]),1,pos('м',trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3,4]))-2)+';');
            Writeln(f,copy(trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3,5]),1,pos('к',trim(Holding.ListFactory[i].ListCex[j].ListEquips[k].grd.Cells[3,5]))-2));
         end;
      end;
   end;
  FormatSettings.ShortDateFormat:=c;
   closefile(f);
end;

procedure TForm1.ToolButton5Click(Sender: TObject);
begin
  DT1.Date := DT1.Date - 1;
  GetSpeedAllAndDrawGraph;
  Form1.SetFocus;
end;

procedure TForm1.ToolButton6Click(Sender: TObject);
begin
  DT1.Date := DT1.Date + 1;
  if DT1.Date > Date() then DT1.Date := DT1.Date - 1;
  GetSpeedAllAndDrawGraph;
  Form1.SetFocus;
end;

procedure TForm1.ToolButton8Click(Sender: TObject);
begin
  GetSpeedAllAndDrawGraph;
  Form1.SetFocus;
end;

procedure TForm1.N12Click(Sender: TObject);
begin
   if N12.Checked then CorrectError:=True else CorrectError:=False;
   DrawHoldingGraph;
end;

procedure TForm1.N13Click(Sender: TObject);
begin
   if N13.Checked then SI8CodeError:=True else SI8CodeError:=False;
   DrawHoldingGraph;
end;

procedure TForm1.N15Click(Sender: TObject);
begin
   Form2.Show;
end;

end.
