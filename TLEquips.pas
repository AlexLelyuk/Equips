unit TLEquips;

interface

uses SysUtils, Classes, Windows, Controls, Graphics, Types, StdCtrls, ExtCtrls, Math, common;

type

  TTimeLine = class(TGraphicControl)
  private
  protected
  public
    Log : TLog;
    speed : TSpeed;
    MaxSpeed : single;
    GraphRect : TRect;
    GraphTop : integer;
    GraphLeft : integer;
    GraphBottom : integer;
    GraphRight : integer;
    NeedRestore : boolean;
    old_x : integer;
    old_y : integer;
    DrawBmp : Graphics.TBitmap;
    CopyBmp : Graphics.TBitmap;
    constructor Create(Aowner: TComponent); override;
    procedure Paint; override;
    function GetCanvas: TCanvas;
    procedure PaintScale;
    procedure Resize; override;
    procedure MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure DrawGraph(spd : TSpeed; lg : TLog; CurMin : integer; needMin,needInter,CorrectError : Boolean; minspd : single; maxspd : single);
    function GetImage : TBitmap;
  published
    property Height default 50;
    property Width default 1505;
    property Top default 1;
    property Left default 1;
  end;

implementation

procedure TTimeLine.DrawGraph(spd : TSpeed; lg : TLog; CurMin : integer; needMin,needInter,CorrectError : Boolean; minspd : single; maxspd : single);
var i : integer;
begin
  for i:=0 to CurMin do begin
    speed[i]:=spd[i];
    log[i]:=lg[i];
  end;
  for i:=CurMin+1 to 1439 do begin
     speed[i]:=0;
     log[i]:=0;
  end;
  if CorrectError then begin
    for i:=1 to CurMin do begin
      if log[i]>=3 then speed[i]:=speed[i-1];
      if speed[i]>(maxspd*KoefInterMaxSpeed) then speed[i]:=speed[i-1];
      if (log[i]=2) and (speed[i]<0.001) then speed[i]:=speed[i-1];
      if i<CurMin then begin
        if (speed[i-1]=0) and (speed[i+1]=0) then speed[i]:=0;
        if (log[i]=0) and (log[i-1]<>0) and (log[i+1]<>0) then speed[i]:=speed[i-1];
      end;
    end;
//    if speed[CurMin]>(maxspd*KoefInterMaxSpeed) then speed[CurMin]:=speed[CurMin-1];
  end;
  if needMin then for i:=0 to CurMin do begin
    if speed[i]<minspd then speed[i]:=0;
  end;
  if needInter then begin
    if speed[0]>0 then speed[0]:=(5*speed[0]+2*speed[1]-speed[2])/6;
      if speed[1]>0 then speed[1]:=(speed[0]+speed[1]+speed[2])/3;
        for i:=2 to 1438 do begin
          if speed[i]>0 then speed[i]:=(speed[i-1]+speed[i]+speed[i+1])/3;
        end;
      if speed[1439]>0 then speed[1439]:=(-1*speed[1437]+2*speed[1438]+5*speed[1439])/6;
  end;
  MaxSpeed:=0;
  for i:=0 to CurMin do if MaxSpeed<speed[i] then MaxSpeed:=speed[i];
  Refresh;
end;

function TTimeLine.GetCanvas;
begin
  GetCanvas := Canvas;
end;

procedure TTimeLine.MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var HintText : string;
    tmpX : integer;
    k : real;
    h : integer;
    i : integer;
    nwidth : integer;
    Rect : TRect;
begin
  tmpX:=X-GraphRect.Left;
  k:=(GraphRect.Right-GraphRect.Left)/1440;
  h:=trunc((tmpX/k+480));
  i:=h-480;
  if h>1440 then h:=h-1440;
  HintText:='';
// восстанавливаем изображение окна
  if NeedRestore then begin
    BitBlt(Canvas.Handle, old_x, old_y, CopyBmp.Width, CopyBmp.Height,CopyBmp.Canvas.Handle, 0, 0, SRCCOPY);
  end;
// отрисовка хинта
  if (X>GraphRect.Left) and (X<GraphRect.Right) and (Y>15) and (Y<GraphRect.Bottom) then begin
    HintText:=FloatToStrF(speed[i],ffFixed,7,1)+' м/мин';
    nwidth := DrawBmp.Canvas.TextWidth(HintText) + 5;
    DrawBmp.Width := nwidth;
    if SI8CodeError then DrawBmp.Height := 45 else DrawBmp.Height := 30;
    DrawBmp.Canvas.Brush.Color := clWhite;
    DrawBmp.Canvas.Font.Color := clBlack;
    Rect.Left := 0;
    Rect.Top := 0;
    Rect.Right := nwidth;
    Rect.Bottom := DrawBmp.Height;
    DrawBmp.Canvas.FillRect(Rect);
    DrawBmp.Canvas.Pen.Color := clBlack;
    DrawBmp.Canvas.Rectangle(Rect);
    DrawBmp.Canvas.Pen.Color := clWhite;
    InflateRect(Rect,-1,-1);
    DrawBmp.Canvas.Rectangle(Rect);
    HintText:=AbsMinuteToStr(h);
    DrawBmp.Canvas.TextOut(3, 2, HintText);
    HintText:=FloatToStrF(speed[i],ffFixed,7,1)+' м/мин';
    DrawBmp.Canvas.TextOut(3, 15, HintText);
    if SI8CodeError then begin
      HintText:='';
      case log[i] of
        0 : HintText:='—четчик не ответил';      //счетчик не ответил
        1 : HintText:='Ok. V=0';      // счетчик ответил, скорость = 0
        2 : HintText:='Ok. V>0';  // счетчик ответил, скорость > 0
        3 : HintText:='CRC Err';        // CRC Error
        5 : HintText:='Packet Err';     // не пришел символ конца пакета
        6 : HintText:='2Query. V=0';      // два опроса, скорость = 0
        7 : HintText:='2Query.V>0';       // два опроса, скорость > 0
  //      8 : Canvas.Pen.Color:=clInfoBk;
      end;
      HintText:=HintText+' ('+IntToStr(log[i])+')';
      DrawBmp.Canvas.TextOut(3, 30, HintText);
    end;
    CopyBmp.Width := nwidth;
    CopyBmp.Height := DrawBmp.Height;
    BitBlt(CopyBmp.Canvas.Handle, 0, 0, nwidth, DrawBmp.Height, Canvas.Handle, X+10, Y-15, SRCCOPY);
    BitBlt(Canvas.Handle, X+10, Y-15, nwidth, DrawBmp.Height, DrawBmp.Canvas.Handle, 0, 0, SRCCOPY);
    NeedRestore := True;
    old_x := X+10;
    old_y := Y-15;
  end;
end;

procedure TTimeLine.Resize;
begin
  inherited Resize;
  GraphRect.Top:=35;
  GraphRect.Bottom:=Height-25;
  GraphRect.Right:=Width-15;
  GraphRect.Left:=50;
end;

constructor TTimeLine.Create(Aowner: TComponent);
//var i : integer;
begin
  inherited Create(Aowner);
  Top := 0;
  Left := 0;
  Width := 1505;
  Height := 200;
  MaxSpeed:=0;
  GraphTop:=35;
  GraphLeft:=50;
  GraphBottom:=25;
  GraphRight:=15;
  GraphRect.Top:=35;
  GraphRect.Bottom:=Height-25;
  GraphRect.Right:=Width-15;
  GraphRect.Left:=50;
{  for i:=0 to 1439 do begin
    speed[i]:=0;
    log[i]:=0;
  end;  }
  FillChar(speed, SizeOf(speed),0);
  FillChar(log, SizeOf(speed),0);
  NeedRestore:=False;
  DrawBmp := Graphics.TBitmap.Create;
  CopyBmp := Graphics.TBitmap.Create;
  CopyBmp.Width := 1;
  CopyBmp.Height := 1;
  old_x:=0;
  old_y:=0;
  onMouseMove:=MouseMove;
end;

procedure TTimeLine.PaintScale;
var
  i: integer;
  h : integer;
  hx : integer;
  hy : integer;
  hspeed : real;
  Vkoef : real;
  MaxYScale : integer;
  PointArray : array[0..1439] of TPoint;
begin
  Canvas.Font.Size := 8;
  Canvas.Pen.Color := clGreen;
  Canvas.Font.Color:=clGreen;
// ось X
  Canvas.MoveTo(GraphRect.Left, GraphRect.Bottom); Canvas.LineTo(GraphRect.Right, GraphRect.Bottom);
// ось Y
  Canvas.MoveTo(GraphRect.Left, GraphRect.Bottom); Canvas.LineTo(GraphRect.Left, GraphRect.Top);
// рисуем подписи оси X
  hspeed:=(GraphRect.Right-GraphRect.Left)/24;
  for i:=0 to 24 do begin
    h:=GraphRect.Left+trunc(i*hspeed);
    Canvas.MoveTo(h, GraphRect.Bottom); Canvas.LineTo(h, GraphRect.Bottom+5);
    if (i+8)<25 then h:=i+8 else h:=(i+8)-24;
    Canvas.TextOut(GraphRect.Left+trunc(i*hspeed)-10, GraphRect.Bottom+5, IntToStr(h)+':00');
  end;
// по оси Y
  h:=trunc((GraphRect.Bottom-GraphRect.Top)/10);
  hspeed:=(GraphRect.Right-GraphRect.Left)/1440;
  MaxYScale:=0;
  if MaxSpeed>0 then begin
    if MaxSpeed<=10 then begin
      MaxYScale:=10;
    end;
    if (MaxSpeed>10) and (MaxSpeed<=100) then begin
      MaxYScale:=trunc(SimpleRoundTo(MaxSpeed,1));
       if (MaxYScale+5)>MaxSpeed then MaxYScale:=MaxYScale+10 else MaxYScale:=MaxYScale+20;
    end;
    if (MaxSpeed>100) and (MaxSpeed<=1000) then begin
       MaxYScale:=trunc(SimpleRoundTo(MaxSpeed,2))+50;
    end;
    if MaxSpeed>1000 then begin
       MaxYScale:=trunc(SimpleRoundTo(MaxSpeed,2))+100;
    end;
    hy:=trunc(MaxYScale/10);
    for i:=0 to 10 do begin
      Canvas.Pen.Style:=psSolid;
      hx:=GraphRect.Bottom-i*h;
      Canvas.MoveTo(GraphRect.Left,hx); Canvas.LineTo(GraphRect.Left-5,hx);
      Canvas.TextOut(GraphRect.Left-30, GraphRect.Bottom-i*h-8, IntToStr(hy*i));
      Canvas.Pen.Style:=psDot;
      Canvas.MoveTo(GraphRect.Left,hx); Canvas.LineTo(GraphRect.Right,hx);
    end;
    Canvas.Pen.Color:=clBlue;
    Canvas.MoveTo(GraphRect.Left, GraphRect.Bottom);
    Vkoef:=(GraphRect.Bottom-GraphRect.Top)/(hy*10);
    Canvas.Pen.Style:=psSolid;
    for i:=0 to 1439 do begin
       PointArray[i]:=Point(GraphRect.Left+trunc(i*hspeed), GraphRect.Bottom-trunc(speed[i]*Vkoef));
    end;
    Canvas.Polyline(PointArray);
  end;
  for i:=0 to 1439 do begin
    case log[i] of
      0 : Canvas.Pen.Color:=clWhite;      //счетчик не ответил
      1 : Canvas.Pen.Color:=clBlack;      // счетчик ответил, скорость = 0
      2 : Canvas.Pen.Color:=clHighLight;  // счетчик ответил, скорость > 0
      3 : Canvas.Pen.Color:=clRed;        // CRC Error
      4 : Canvas.Pen.Color:=clLime;       //  нет такого лога
      5 : Canvas.Pen.Color:=clYellow;     // не пришел символ конца пакета
      6 : Canvas.Pen.Color:=clGreen;      // два опроса, скорость = 0
      7 : Canvas.Pen.Color:=clAqua;       // два опроса, скорость > 0
//      8 : Canvas.Pen.Color:=clInfoBk;
    end;
    hx:=GraphRect.Left+trunc(i*hspeed);
    h:=GraphRect.Top-20;
    hy:=GraphRect.Top-5;
    Canvas.MoveTo(hx, h); Canvas.LineTo(hx, hy);
  end;
end;

procedure TTimeLine.Paint;
var
  Rect : TRect;
begin
  if HasParent then begin
    Rect := GetClientRect;
    Canvas.Brush.Color:=clSilver;
    Canvas.Pen.Color:=clBlack;
    Canvas.FillRect(Rect);
    Canvas.Font.Size:=8;
    Frame3D(Canvas, Rect, clBtnShadow, clBtnHighlight, 2);
    PaintScale;
  end;
end;

function TTimeLine.GetImage : TBitmap;
var bmp : TBitmap;
    MyGraphRect : TRect;
  i: integer;
  h : integer;
  hx : integer;
  hy : integer;
  hspeed : real;
  Vkoef : real;
  MaxYScale : integer;
  PointArray : array[0..1439] of TPoint;
begin
  MyGraphRect.Top:=35;
  MyGraphRect.Bottom:=Height-25;
  MyGraphRect.Right:=Width-15;
  MyGraphRect.Left:=50;

  //  возвращает bitmap  с графиком
  bmp:=TBitmap.Create;
  bmp.SetSize(1505, 160);

  MyGraphRect.Top:=35;
  MyGraphRect.Bottom:=bmp.Height-25;
  MyGraphRect.Right:=bmp.Width-15;
  MyGraphRect.Left:=50;

  bmp.Canvas.Font.Size := 8;
  bmp.Canvas.Pen.Color := clGreen;
  bmp.Canvas.Font.Color:=clGreen;
// ось X
  bmp.Canvas.MoveTo(MyGraphRect.Left, MyGraphRect.Bottom); bmp.Canvas.LineTo(MyGraphRect.Right, MyGraphRect.Bottom);
// ось Y
  bmp.Canvas.MoveTo(MyGraphRect.Left, MyGraphRect.Bottom); bmp.Canvas.LineTo(MyGraphRect.Left, MyGraphRect.Top);
// рисуем подписи оси X
  hspeed:=(MyGraphRect.Right-MyGraphRect.Left)/24;
  for i:=0 to 24 do begin
    h:=MyGraphRect.Left+trunc(i*hspeed);
    bmp.Canvas.MoveTo(h, MyGraphRect.Bottom); bmp.Canvas.LineTo(h, MyGraphRect.Bottom+5);
    if (i+8)<25 then h:=i+8 else h:=(i+8)-24;
    bmp.Canvas.TextOut(MyGraphRect.Left+trunc(i*hspeed)-10, MyGraphRect.Bottom+5, IntToStr(h)+':00');
  end;
// по оси Y
  h:=trunc((MyGraphRect.Bottom-MyGraphRect.Top)/10);
  hspeed:=(MyGraphRect.Right-MyGraphRect.Left)/1440;
  MaxYScale:=0;
  if MaxSpeed>0 then begin
    if MaxSpeed<=10 then begin
      MaxYScale:=10;
    end;
    if (MaxSpeed>10) and (MaxSpeed<=100) then begin
      MaxYScale:=trunc(SimpleRoundTo(MaxSpeed,1));
       if (MaxYScale+5)>MaxSpeed then MaxYScale:=MaxYScale+10 else MaxYScale:=MaxYScale+20;
    end;
    if (MaxSpeed>100) and (MaxSpeed<=1000) then begin
       MaxYScale:=trunc(SimpleRoundTo(MaxSpeed,2))+50;
    end;
    if MaxSpeed>1000 then begin
       MaxYScale:=trunc(SimpleRoundTo(MaxSpeed,2))+100;
    end;
    hy:=trunc(MaxYScale/10);
    for i:=0 to 10 do begin
      bmp.Canvas.Pen.Style:=psSolid;
      hx:=MyGraphRect.Bottom-i*h;
      bmp.Canvas.MoveTo(MyGraphRect.Left,hx); bmp.Canvas.LineTo(MyGraphRect.Left-5,hx);
      bmp.Canvas.TextOut(MyGraphRect.Left-30, MyGraphRect.Bottom-i*h-8, IntToStr(hy*i));
      bmp.Canvas.Pen.Style:=psDot;
      bmp.Canvas.MoveTo(MyGraphRect.Left,hx); bmp.Canvas.LineTo(MyGraphRect.Right,hx);
    end;
    bmp.Canvas.Pen.Color:=clBlue;
    bmp.Canvas.MoveTo(MyGraphRect.Left, MyGraphRect.Bottom);
    Vkoef:=(MyGraphRect.Bottom-MyGraphRect.Top)/(hy*10);
    bmp.Canvas.Pen.Style:=psSolid;
    for i:=0 to 1439 do begin
       PointArray[i]:=Point(MyGraphRect.Left+trunc(i*hspeed), MyGraphRect.Bottom-trunc(speed[i]*Vkoef));
    end;
    bmp.Canvas.Polyline(PointArray);
  end;
  for i:=0 to 1439 do begin
    case log[i] of
      0 : bmp.Canvas.Pen.Color:=clWhite;      //счетчик не ответил
      1 : bmp.Canvas.Pen.Color:=clBlack;      // счетчик ответил, скорость = 0
      2 : bmp.Canvas.Pen.Color:=clHighLight;  // счетчик ответил, скорость > 0
      3 : bmp.Canvas.Pen.Color:=clRed;        // CRC Error
      4 : bmp.Canvas.Pen.Color:=clLime;       //  нет такого лога
      5 : bmp.Canvas.Pen.Color:=clYellow;     // не пришел символ конца пакета
      6 : bmp.Canvas.Pen.Color:=clGreen;      // два опроса, скорость = 0
      7 : bmp.Canvas.Pen.Color:=clAqua;       // два опроса, скорость > 0
//      8 : Canvas.Pen.Color:=clInfoBk;
    end;
    hx:=MyGraphRect.Left+trunc(i*hspeed);
    h:=MyGraphRect.Top-20;
    hy:=MyGraphRect.Top-5;
    bmp.Canvas.MoveTo(hx, h); Canvas.LineTo(hx, hy);
  end;



  result:=bmp;
end;

end.
