unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, IBX.IBCustomDataSet,
  IBX.IBQuery, IBX.IBDatabase, Vcl.StdCtrls, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdHTTP, System.JSON, Vcl.Samples.Spin, DateUtils,
  Vcl.Grids, WinProcs, common, System.Generics.Collections;

const
  sko : array[1..5] of String = ('sko.json','sko1.json', 'sko2.json', 'sko3.json', 'sko5.json');
  kab : array[1..5] of String = ('kab.json','kab1.json', 'kab2.json', 'kab3.json', 'kab7.json');
  kat : array[1..2] of String = ('kat.json', 'kat1.json');
  plast : array[1..2] of String = ('plast.json', 'plast2.json');

type
  TForm2 = class(TForm)
    IdHTTP1: TIdHTTP;
    Button1: TButton;
    ComboBox1: TComboBox;
    SpinEdit1: TSpinEdit;
    StringGrid1: TStringGrid;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    Button2: TButton;
    procedure Button1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.Button1Click(Sender: TObject);
var res : string;
    Json, {obj,} item: TJSONObject;
    jArr:TJSONArray;
    v:TJSONValue;
    i, j : integer;
    jsonToSend : TStringStream;
    fname : string;
    sss : string;
begin
//  for i:=0 to 30 do StringGrid1.Cells[i,1]:='';
  for i:=1 to StringGrid1.ColCount-1 do
    for j:=1 to StringGrid1.RowCount-1 do StringGrid1.Cells[i,j]:='';
      
  case ComboBox2.ItemIndex of
     0 : begin
           StringGrid1.RowCount:=6;
           StringGrid1.Cells[0,1]:='СКО всего';
           StringGrid1.Cells[0,2]:='Цех №1';
           StringGrid1.Cells[0,3]:='Цех №2';
           StringGrid1.Cells[0,4]:='Цех №3';
           StringGrid1.Cells[0,5]:='Цех №5';
//           fname:=sko[ComboBox3.ItemIndex];
         end;
     1 : begin
           StringGrid1.RowCount:=6;
           StringGrid1.Cells[0,1]:='ЭМ-Кабель всего';
           StringGrid1.Cells[0,2]:='Цех №1';
           StringGrid1.Cells[0,3]:='Цех №2';
           StringGrid1.Cells[0,4]:='Цех №3';
           StringGrid1.Cells[0,5]:='Цех №7';
 //          fname:=kab[ComboBox3.ItemIndex];
         end;
     2 : begin
           StringGrid1.RowCount:=3;
           StringGrid1.Cells[0,1]:='ЭМ-Кат всего';
           StringGrid1.Cells[0,2]:='Цех';
//           fname:='kat.num';
         end;
     3 : begin
           StringGrid1.RowCount:=3;
           StringGrid1.Cells[0,1]:='ЭМ-Пласт всего';
           StringGrid1.Cells[0,2]:='Цех №2';
//           fname:='plast.num';
         end;
  end;
  for j:=1 to StringGrid1.RowCount-1 do begin
    case ComboBox2.ItemIndex of
       0 : fname:=sko[j];
       1 : fname:=kab[j];
       2 : fname:=kat[j];
       3 : fname:=plast[j];
    end;
    jsonToSend := TStringStream.Create;              // ', "fname":'+sko[ComboBox3.ItemIndex]+
    sss:='{"mon":'+IntToStr(ComboBox1.ItemIndex+1)+', "god":'+IntToStr(SpinEdit1.Value-2000)+'}';
    jsonToSend.WriteString('{"mon":'+IntToStr(ComboBox1.ItemIndex+1)+', "god":'+IntToStr(SpinEdit1.Value-2000)+', "f_name":'+'"'+fname+'"'+'}');
    idHTTP1.Request.ContentType := 'application/json';
    res:=idHTTP1.post(RestAddr+'get', jsonToSend);
    jsonToSend.Free;
    if res<>'{}' then begin
      Json := TJSONObject.ParseJSONValue(res) as TJSONObject;
      jArr := Json.getValue('num') as TJSONArray;
      for i:=0 to jArr.Count-1 do begin
        item:=jarr.Items[i] as TJSONObject;
        v:=item.GetValue(IntToStr(i+1));
        StringGrid1.Cells[i+1,j]:=v.ToString;
      end;
    end;
  end;
end;

procedure TForm2.Button2Click(Sender: TObject);
var i, j : integer;
    JsonToSend : TStringStream;
    s : string;
    res : string;
begin
   case ComboBox2.ItemIndex of
      0 : begin
         JsonToSend:=TStringStream.Create;
         for j:=1 to 5 do begin
           s:='{"fname":"'+sko[j]+'", "mon":'+IntToStr(ComboBox1.ItemIndex+1)+', "god":'+IntToStr(SpinEdit1.Value-2000)+', "num":[';
           for i:=1 to 31 do begin
              if StringGrid1.Cells[i,j]='' then
                s:=s+'{"'+IntToStr(i)+'":0}'
              else
                s:=s+'{"'+IntToStr(i)+'":'+StringGrid1.Cells[i,j]+'}';
              if i<>31 then s:=s+', ';
           end;
           s:=s+']}';
           JsonToSend.WriteString(s);
           idHTTP1.Request.ContentType := 'application/json';
           res:=idHTTP1.post(RestAddr+'put', jsonToSend);
           JsonToSend.Clear;
         end;
         jsonToSend.Free;
      end;
      1 : begin
         JsonToSend:=TStringStream.Create;
         for j:=1 to 5 do begin
           s:='{"fname":"'+kab[j]+'", "mon":'+IntToStr(ComboBox1.ItemIndex+1)+', "god":'+IntToStr(SpinEdit1.Value-2000)+', "num":[';
           for i:=1 to 31 do begin
              if StringGrid1.Cells[i,j]='' then
                s:=s+'{"'+IntToStr(i)+'":0}'
              else
                s:=s+'{"'+IntToStr(i)+'":'+StringGrid1.Cells[i,j]+'}';
              if i<>31 then s:=s+', ';
           end;
           s:=s+']}';
           JsonToSend.WriteString(s);
           idHTTP1.Request.ContentType := 'application/json';
           res:=idHTTP1.post(RestAddr+'put', jsonToSend);
           JsonToSend.Clear;
         end;
         jsonToSend.Free;
      end;
      2 : begin
         JsonToSend:=TStringStream.Create;
         for j:=1 to 2 do begin
           s:='{"fname":"'+kat[j]+'", "mon":'+IntToStr(ComboBox1.ItemIndex+1)+', "god":'+IntToStr(SpinEdit1.Value-2000)+', "num":[';
           for i:=1 to 31 do begin
              if StringGrid1.Cells[i,j]='' then
                s:=s+'{"'+IntToStr(i)+'":0}'
              else
                s:=s+'{"'+IntToStr(i)+'":'+StringGrid1.Cells[i,j]+'}';
              if i<>31 then s:=s+', ';
           end;
           s:=s+']}';
           JsonToSend.WriteString(s);
           idHTTP1.Request.ContentType := 'application/json';
           res:=idHTTP1.post(RestAddr+'put', jsonToSend);
           JsonToSend.Clear;
         end;
         jsonToSend.Free;
      end;
      3 : begin
         JsonToSend:=TStringStream.Create;
         for j:=1 to 2 do begin
           s:='{"fname":"'+plast[j]+'", "mon":'+IntToStr(ComboBox1.ItemIndex+1)+', "god":'+IntToStr(SpinEdit1.Value-2000)+', "num":[';
           for i:=1 to 31 do begin
              if StringGrid1.Cells[i,j]='' then
                s:=s+'{"'+IntToStr(i)+'":0}'
              else
                s:=s+'{"'+IntToStr(i)+'":'+StringGrid1.Cells[i,j]+'}';
              if i<>31 then s:=s+', ';
           end;
           s:=s+']}';
           JsonToSend.WriteString(s);
           idHTTP1.Request.ContentType := 'application/json';
           res:=idHTTP1.post(RestAddr+'put', jsonToSend);
           JsonToSend.Clear;
         end;
         jsonToSend.Free;
      end;
   end;
   Application.MessageBox(PChar('Данные сохранены в базу'),PChar('Сообщение.'),MB_OK);
end;

procedure TForm2.ComboBox2Change(Sender: TObject);
begin
  if ComboBox2.ItemIndex=0 then begin
    ComboBox3.Items.Clear;
    ComboBox3.Items.Add('Цех №1');
    ComboBox3.Items.Add('Цех №2');
    ComboBox3.Items.Add('Цех №3');
    ComboBox3.Items.Add('Цех №5');
    ComboBox3.ItemIndex:=0;
    StringGrid1.RowCount:=6;
    StringGrid1.Cells[0,1]:='СКО всего';
    StringGrid1.Cells[0,2]:='Цех №1';
    StringGrid1.Cells[0,3]:='Цех №2';
    StringGrid1.Cells[0,4]:='Цех №3';
    StringGrid1.Cells[0,5]:='Цех №5';
  end;
  if ComboBox2.ItemIndex=1 then begin
    ComboBox3.Items.Clear;
    ComboBox3.Items.Add('Цех №1');
    ComboBox3.Items.Add('Цех №2');
    ComboBox3.Items.Add('Цех №3');
    ComboBox3.Items.Add('Цех №7');
    ComboBox3.ItemIndex:=0;
    StringGrid1.RowCount:=6;
    StringGrid1.Cells[0,1]:='ЭМ-Кабель всего';
    StringGrid1.Cells[0,2]:='Цех №1';
    StringGrid1.Cells[0,3]:='Цех №2';
    StringGrid1.Cells[0,4]:='Цех №3';
    StringGrid1.Cells[0,5]:='Цех №7';
  end;
  if ComboBox2.ItemIndex=2 then begin
    ComboBox3.Items.Clear;
    ComboBox3.Items.Add('ЭМ-КАТ');
    ComboBox3.ItemIndex:=0;
    StringGrid1.RowCount:=3;
    StringGrid1.Cells[0,1]:='ЭМ-Кат всего';
    StringGrid1.Cells[0,2]:='Цех';
  end;
  if ComboBox2.ItemIndex=3 then begin
    ComboBox3.Items.Clear;
    ComboBox3.Items.Add('Цех №2');
    ComboBox3.ItemIndex:=0;
    StringGrid1.RowCount:=3;
    StringGrid1.Cells[0,1]:='ЭМ-Пласт всего';
    StringGrid1.Cells[0,2]:='Цех №2';
  end;
end;

procedure TForm2.FormActivate(Sender: TObject);
var i : integer;
begin
   StringGrid1.ColWidths[0]:=100;
   ComboBox1.ItemIndex:=MonthOf(Date)-1;
   SpinEdit1.Value:=YearOf(Date);
//  тут надо выставить индекс в зависимости от завода, запустившего программу
   ComboBox2.ItemIndex:=0;
// --------
   ComboBox3.Items.Clear;
   ComboBox3.Items.Add('Цех №1');
   ComboBox3.Items.Add('Цех №2');
   ComboBox3.Items.Add('Цех №3');
   ComboBox3.Items.Add('Цех №5');
   ComboBox3.ItemIndex:=0;
   case ComboBox2.ItemIndex of
     0 : begin
           StringGrid1.RowCount:=6;
           StringGrid1.Cells[0,1]:='СКО всего';
           StringGrid1.Cells[0,2]:='Цех №1';
           StringGrid1.Cells[0,3]:='Цех №2';
           StringGrid1.Cells[0,4]:='Цех №3';
           StringGrid1.Cells[0,5]:='Цех №5';
         end;
     1 : begin
           StringGrid1.RowCount:=6;
           StringGrid1.Cells[0,1]:='ЭМ-Кабель всего';
           StringGrid1.Cells[0,2]:='Цех №1';
           StringGrid1.Cells[0,3]:='Цех №2';
           StringGrid1.Cells[0,4]:='Цех №3';
           StringGrid1.Cells[0,5]:='Цех №7';
         end;
     2 : begin
           StringGrid1.RowCount:=3;
           StringGrid1.Cells[0,1]:='ЭМ-Кат всего';
           StringGrid1.Cells[0,2]:='Цех';
         end;
     3 : begin
           StringGrid1.RowCount:=3;
           StringGrid1.Cells[0,1]:='ЭМ-Пласт всего';
           StringGrid1.Cells[0,2]:='Цех №2';
         end;
   end;
   for i:=1 to 31 do StringGrid1.Cells[i,0]:=IntToStr(i);  
end;

procedure TForm2.StringGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var// Format : word;
    //C: array[0..255] of Char;
    txt : string;
begin
  txt:= StringGrid1.Cells[ACol,ARow];
//  StringGrid1.Canvas.Brush.Color:= clSkyBlue;
//  StringGrid1.Canvas.FillRect(Rect);
  if (ARow>0) and (ACol>0) then StringGrid1.Canvas.Brush.Color:= clWhite
  else StringGrid1.Canvas.Brush.Color:= clBtnFace;
  if gdFocused in State then StringGrid1.Canvas.Brush.Color:=clGradientActiveCaption;
  if ACol=0 then txt:='  '+txt;
  InflateRect(Rect, -1, -1);
  StringGrid1.Canvas.FillRect(Rect);
  if ACol=0 then StringGrid1.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfLeft,tfSingleLine])
  else StringGrid1.Canvas.TextRect(Rect,txt,[tfVerticalCenter,tfCenter,tfSingleLine]);
end;

end.
