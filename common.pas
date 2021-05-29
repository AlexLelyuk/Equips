unit common;

interface
uses SysUtils, Graphics, Classes, System.JSON, idHTTP, DateUtils;

const Month: Array [1..12] of String = ('Январь', 'Февраль', 'Март', 'Апрель',
                                'Май', 'Июнь', 'Июль', 'Август',
                                'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь');
// индекс границы
  xlDiagonalDown=5;
  xlDiagonalUp=6;
  xlEdgeBottom=9;
  xlEdgeLeft=7;
  xlEdgeRight=10;
  xlEdgeTop=8;
  xlInsideHorizontal=12;
  xlInsideVertical=11;
// Стили линий (LineStyle) могут быть следующие:
  xlContinuous=1;
  xlDash=-4115;
  xlDashDot=4;
  xlDashDotDot=5;
  xlDot=-4118;
  xlDouble=-4119;
  xlLineStyleNone=-4142;
  xlSlantDashDot=13;
// Для толщины линии определены константы:
  xlHairline=1;
  xlMedium=-4138;
  xlThick=4;
  xlThin=2;
//
  xlCenter=-4108;
//
  xlWBATWorksheet = -4167;


var
  BeginDay : record
     Hour : integer;
     Minute : integer;
  end;
  KoefInterMaxSpeed : integer;
  needMin: Boolean;
  needInter : Boolean;
  CorrectError : Boolean;
  SI8CodeError : Boolean;
  GraphColors : record
     BackGround : TColor;
     Labels : TColor;
     ScaleLines : TColor;
     Smena : TColor;
     GridLines : TColor;
     LineWidth : integer;
  end;
  ftp  : record
     enable : integer;
     user : string;
     pwd : string;
     port : integer;
     server : string;
  end;
  RestAddr : string;
  HightCaptionCex : integer;
  HightGraph : integer;
  WidthButCaption : integer;
  WidthGrd : integer;


type TSpeed=array[0..1439] of single;
     TLog=array[0..1439] of byte;


function TimeStr(t : string): string;
function TimeToAbsMinute(t : string): integer;
function AbsMinuteToStr(minute : integer) : string;
function PADL(Src: string; Lg: Integer): string;
function PADR(Src: string; Lg: Integer): string;
function GetNameMonth(d : integer) : string;
function GetFileSize(f : string) : longint;
function GetJSON(path: string; jfile : string; d: TDateTime) : string;

implementation

function GetJSON(path: string; jfile : string; d: TDateTime) : string;
var //i, j, k : integer;
//    dir_name : string;
    jsonToSend : TStringStream;
//    Json, {obj,} item: TJSONObject;
    //jArr:TJSONArray;
    //v:TJSONValue;
    mon, god : integer;
    idHTTP1 : TidHTTP;
    res : string;
    s : string;
begin
//  if FileExists(path+jfile) then begin
    idHTTP1:=TidHTTP.Create;
    Formatsettings.ShortDateFormat:='ddmmyy';
    mon:=MonthOf(d);
    god:=YearOf(d);
//    dir_name:=copy(DateToStr(d),3,4);
    jsonToSend := TStringStream.Create;
    s:='{"mon":'+IntToStr(mon)+', "god":'+IntToStr(god-2000)+', "f_name":'+jfile+'}';
    jsonToSend.WriteString(s);
    idHTTP1.Request.ContentType := 'application/json';
    res:=idHTTP1.post(RestAddr+'get', jsonToSend);
    jsonToSend.Free;
    idHTTP1.Free;
    GetJSON:=res;
//  end
//  else GetJSON:='{}';
end;

function GetFileSize(f : string) : longint;
var SearchRec: TSearchRec;
begin
  if FindFirst(f, faAnyFile, SearchRec)=0 then begin
    GetFileSize := SearchRec.Size;
    FindClose(SearchRec);
  end
  else GetFileSize:=-1;
end;

function GetNameMonth(d : integer) : string;
begin
    if (d < 1) or (d > 12) then result:='unknown' else result:=Month[d];
end;

function PADR(Src: string; Lg: Integer): string;
begin
  Result := Src;
  while Length(Result) < Lg do Result := Result + ' ';
end;

function PADL(Src: string; Lg: Integer): string;
begin
  Result := Src;
  while Length(Result) < Lg do Result := ' ' + Result;
end;

function TimeStr(t : string):string;
var i:integer;
    s, hour, minute : string;
begin
   s:=trim(t);
   i:=pos(':',s);
   hour:=copy(s,1,i-1);
   minute:=copy(s,i+1,MaxInt);
   TimeStr:=IntToStr(StrToInt(hour)*60+StrToInt(minute));
end;

function TimeToAbsMinute(t : string): integer;
var i:integer;
    s, hour, minute : string;
begin
   s:=trim(t);
   i:=pos(':',s);
   hour:=copy(s,1,i-1);
   minute:=copy(s,i+1,MaxInt);
   TimeToAbsMinute:=StrToInt(hour)*60+StrToInt(minute);
end;

function AbsMinuteToStr(minute : integer) : string;
var hh,mm :string;
begin
  hh:=IntToStr(minute div 60);
  if Length(hh)=1 then hh:='0'+hh;
  mm:=IntToStr(minute mod 60);
  if Length(mm)=1 then mm:='0'+mm;
  AbsMinuteToStr:=hh+':'+mm;
end;


end.
