unit EquipClass;

interface

uses Windows, TLEquips, StdCtrls, Grids, common, System.SysUtils, System.Classes, System.JSON, System.Generics.Collections, Vcl.Forms,
     DateUtils, VCL.Dialogs, idHTTP;

type

  TSpeed_month = array[1..31,0..1439] of single;

//  TEquip = class
  TEquip = class(TComponent)
    { - любое оборудование обязательно должно иметь уникальный KOD, т.к. он явля-
      ется индексом в массиве данных. Даже в том случае, если датчик реально не
      подключен. Один KOD = один датчик, даже если под одним кодом прописано нес-
      колько единиц оборудования - все они будут иметь одни и те же показания.
      - id присваивается оборудованию один раз и навсегда. На него не влияют ни из-
      менения параметров, ни перенос в другую бригаду, ни создание записи-копии с
      другим диапазоном дат.
      - Изменение в справочнике названия оборудования, порядка сортировки, видимос-
      ти, единиц измерения, цеха, координат, размеров и скоростей не влечет ника-
      ких последствий.
      - При изменении кода, тарифов, коэ-та, адреса СИ8, коэф-та длины СИ8 добавля-
      ется новая запись с другим диапазоном дат, а id остается прежним.
      - id+sdate или id+edate однозначно определяют запись в мправочнике.
      - Изменение кода (номера датчика) требует обновления данных с даты изменения.
    }
    public
    id: integer; // уникальный идентификатор в пределах времени действия
    si8adr: integer; // адрес счетчика "СИ8"
    ord: integer; // порядок сортировки
    info: string; // название
    minspd: single; // минимальная учитываемая скорость
    maxspd: single; // максимальная учитываемая скорость
    MaxSpdPerDay: single; // маскимальная скорость за сутки
    sdate: integer; // дата начала действия записи (включительно)
    edate: integer; // дата окончания действия записи (включительно)
    ButCaption: TButton;
    NewChart: TTimeLine;
    grd: TStringGrid;
    indZavod: byte;
    kol_sd: single; // изготовлено за сутки
    kol_month: single;
    kmv_month: single;
    sm_b : integer;
    sm_e : integer;
    ext : string; // расширение файла данных счетчика, откуда прочитались данные
    Zavod: string[30];
    speed: TSpeed{array[0 .. 1439] of single}; // таблица скорости по минутно
    energy: array[0..23] of single; // таблица потребляемой энергии в разрезе каждого часа
    log: TLog{array [0 .. 1439] of byte};
//    kmv_month : single;
    kmv_rab_month : single;
    kol_day: array[1..31] of single;
    kmv_day: array[1..31] of single;
    kmv_rab_day: array[1..31] of single;
    speed_month : TSpeed_month;
    log_month : array[1..31, 0..1439] of byte;
    visible : boolean;
//    json : string;
//    procedure Interpolate(CurMin : integer);
    procedure InterpolateMonth(ToDay : integer);
    constructor Create;
  protected
  public
//    constructor Create(_id: integer);
  end;

  TCex = class
      id : integer;  // код цеха
      ord : integer;
      name : string; //  наименование цеха
      ListEquips : array of TEquip;
      Count : integer;
      json : string;
      kmv : single;
      ButCaption: TButton;
//      kmv_day: array[1..31] of single;
      kmv_rab_day: array[1..31] of single;
      data_json : array[1..31] of integer;
      constructor Create;
    protected
    public
  end;

   TFactory = class
       id : integer;  // код предприятия
       ord : integer;
       name : string; // наименование предприятия
       ListCex : array of TCex;
       IdCexs : set of byte;
       Count : integer;
       KolObr : integer;
       json : string;
       kmv : single;
       ButCaption: TButton;
//       kmv_day: array[1..31] of single;
       kmv_rab_day: array[1..31] of single;
       data_json : array[1..31] of integer;
       constructor Create;
     protected
     public
   end;

   THolding = class
       ListFactory : array of TFactory;
       Count : integer; // количество предприятий
       KolObr : integer;
       constructor Create;
       procedure LoadData(path : string);
       procedure ClearItems;
       function GetEquipByAdr(si8adr : integer; ext : string) : TEquip;
       procedure GetSpeed(path: string; ext: string; d: TDateTime);
       procedure GetStatMonth(path: string; ext: string; d: TDateTime);
       procedure GetNumEmp(path: string; d: TDateTime);
     protected
     public
   end;

implementation

constructor TCex.Create;
begin
  ButCaption:=TButton.Create(nil);
  ButCaption.WordWrap:=True;
//  ButCaption.Parent:=self;
  Count:=0;
end;

constructor TFactory.Create;
begin
  ButCaption:=TButton.Create(nil);
  ButCaption.WordWrap:=True;
//  ButCaption.Parent:=self;
  Count:=0;
  KolObr:=0;
  IdCexs:=[];
end;

constructor TEquip.Create;
begin
  sm_b := {450}BeginDay.Minute+BeginDay.Hour*60;
  sm_e := {990}1440-sm_b;
  grd := TStringGrid.Create(nil);
  grd.DefaultRowHeight:=20;
  grd.DefaultColWidth:=60;
  grd.ColCount:=4;
  grd.RowCount:=7;
  grd.ColWidths[0] := 90;
  grd.Cells[1, 0] := '      I см';
  grd.Cells[2, 0] := '     II см';
  grd.Cells[3, 0] := '     сутки';
  grd.Cells[0, 1] := ' Время';
  grd.Cells[0, 2] := ' КМВ';
  grd.Cells[0, 3] := ' Макс.скорость';
  grd.Cells[0, 4] := ' Ср.скорость';
  grd.Cells[0, 5] := ' Изготовлено';
  grd.Cells[0, 6] := ' Кол.остановок';
  grd.Cells[0, 7] := ' Изготовлено за мес.';
  NewChart:=TTimeLine.Create(nil);
  ButCaption:=TButton.Create(self);
//  ButCaption.owner:=self;
  ButCaption.WordWrap:=True;
end;

constructor THolding.Create;
begin
  Count:=0;
  KolObr:=0;
end;

procedure THolding.ClearItems;
var i, j, k : integer;
begin
// не забыть добавить освобождение памяти от списка оборудования
   for j:=0 to Count-1 do begin
       for i:=0 to ListFactory[j].count-1 do begin
           for k:=0 to ListFactory[j].ListCex[i].count-1 do ListFactory[j].ListCex[i].ListEquips[k].Free;
           ListFactory[j].ListCex[i].Free;
       end;
       ListFactory[j].Free;
   end;
end;

procedure THolding.LoadData(path : string);
var fs : TFileStream;
    FactoryJsonStr : AnsiString;
    EquipsJsonStr : AnsiString;
    Json, Factory, Cex, Equip, Eq: TJSONObject;
    ArrFactory, ArrCex, ArrEquips : TJSONArray;
//    ArrEq : TJsonArray;
 //   id, ord, name, json_s : TJSONValue;
    i, j, k : integer;
    id : string;
    s : string;
begin
// загружаем данные и json файлов
  FormatSettings.DecimalSeparator:='.';
  if FileExists(path+'factory.json') then begin
     fs:=TFileStream.Create(path+'factory.json',fmOpenRead);
     fs.Position:=0;
     SetLength(FactoryJsonStr,fs.Size);
     fs.Read(FactoryJsonStr[1],fs.Size);
     fs.Free;
     fs:=TFileStream.Create(path+'equips.json',fmOpenRead);
     fs.Position:=0;
     SetLength(EquipsJsonStr,fs.Size);
     fs.Read(EquipsJsonStr[1],fs.Size);
     fs.Free;
     if FactoryJsonStr<>'' then begin
        Json := TJSONObject.ParseJSONValue(FactoryJsonStr) as TJSONObject;
        ArrFactory := Json.getValue('Factory') as TJSONArray;
        Equip := TJSONObject.ParseJSONValue(EquipsJsonStr) as TJSONObject;
        Count:=ArrFactory.Count;
        SetLength(ListFactory, Count);
        for j:=0 to Count-1 do begin
           Factory:=ArrFactory.Items[j] as TJSONObject;
           ListFactory[j]:=TFactory.Create;
           ListFactory[j].id:=StrToInt(Factory.GetValue('id').ToString);
           ListFactory[j].ord:=StrToInt(Factory.GetValue('ord').ToString);
           ListFactory[j].name:=StringReplace(Factory.GetValue('name').ToString,'"','',[rfReplaceAll, rfIgnoreCase]);
           ListFactory[j].json:=Factory.GetValue('json').ToString;
           ArrCex := Factory.GetValue('workshops') as TJSONArray;
           ListFactory[j].count:=ArrCex.Count;
           SetLength(ListFactory[j].ListCex, ListFactory[j].count);
           for i:=0 to ListFactory[j].count-1 do begin
              Cex:=ArrCex.Items[i] as TJSONObject;
              ListFactory[j].ListCex[i]:=TCex.Create;
              ListFactory[j].ListCex[i].id:=StrToInt(Cex.GetValue('id').ToString);
              ListFactory[j].ListCex[i].ord:=StrToInt(Cex.GetValue('ord').ToString);
              ListFactory[j].ListCex[i].name:=StringReplace(Cex.GetValue('name').ToString,'"','',[rfReplaceAll, rfIgnoreCase]);
              ListFactory[j].ListCex[i].json:=Cex.GetValue('json').ToString;
              Include(ListFactory[j].IdCexs,ListFactory[j].ListCex[i].id);
//   тут еще будем загружать список оборудования
              id:=IntToStr(ListFactory[j].ListCex[i].id);
              ArrEquips:=Equip.GetValue(id) as TJSONArray;
              if ArrEquips<>nil then begin
                ListFactory[j].ListCex[i].Count:=ArrEquips.Count;
                ListFactory[j].KolObr:=ListFactory[j].KolObr+ArrEquips.Count;
                KolObr:=KolObr+ArrEquips.Count;
                SetLength(ListFactory[j].ListCex[i].ListEquips,ListFactory[j].ListCex[i].Count);
                for k:=0 to ListFactory[j].ListCex[i].Count-1 do begin
                  Eq:=ArrEquips.Items[k] as TJSONObject;
                  ListFactory[j].ListCex[i].ListEquips[k]:=TEquip.Create;
                  ListFactory[j].ListCex[i].ListEquips[k].id:=StrToInt(Eq.GetValue('id').ToString);
                  ListFactory[j].ListCex[i].ListEquips[k].si8adr:=StrToInt(Eq.GetValue('si8adr').ToString);
                  ListFactory[j].ListCex[i].ListEquips[k].ord:=StrToInt(Eq.GetValue('ord').ToString);
//                  StringReplace(Holding.ListFactory[i].ListCex[j].ListEquips[k].info,'"','',[rfReplaceAll, rfIgnoreCase])
                  s:=StringReplace(Eq.GetValue('name').ToString,'"','',[rfReplaceAll, rfIgnoreCase]);
                  s:=StringReplace(s,'\','',[rfReplaceAll, rfIgnoreCase]);
                  ListFactory[j].ListCex[i].ListEquips[k].info:=s;
                  ListFactory[j].ListCex[i].ListEquips[k].minspd:=StrToFloat(Eq.GetValue('minspd').ToString);
                  ListFactory[j].ListCex[i].ListEquips[k].maxspd:=StrToFloat(Eq.GetValue('maxspd').ToString);
                  ListFactory[j].ListCex[i].ListEquips[k].sdate:=StrToInt(Eq.GetValue('sdate').ToString);
                  ListFactory[j].ListCex[i].ListEquips[k].edate:=StrToInt(Eq.GetValue('edate').ToString);
                  ListFactory[j].ListCex[i].ListEquips[k].ext:=copy(Eq.GetValue('ext').ToString,2,Length(Eq.GetValue('ext').ToString)-2);
                end;
              end;

          end;
        end;
      end;
    end;
  end;

function THolding.GetEquipByAdr(si8adr : integer; ext : string) : TEquip;
var i : integer;
    j : Integer;
    k : Integer;
    res : TEquip;
begin
   res:=nil;
   for i:=0 to Count-1 do begin
      for j:=0 to ListFactory[i].Count-1 do begin
         for k:=0 to ListFactory[i].ListCex[j].Count-1 do begin
            if (ListFactory[i].ListCex[j].ListEquips[k].si8adr=si8adr) and (ListFactory[i].ListCex[j].ListEquips[k].ext=ext) then begin
              res:=ListFactory[i].ListCex[j].ListEquips[k];
            end;
         end;
      end;
   end;
   result:=res;
end;

procedure THolding.GetSpeed(path: string; ext: string; d: TDateTime);
var
  f_si8, f_lo8: TFileStream;
  kol, h, m, num, i{, j}: byte;
//  adr{, x}: integer;
  eqp : TEquip;
  pok1: single;
  fname, dir_name{, y}: string;
  buf_si8: array of byte;
  ind_buf_si8, dl_buf : Longint;
begin
  Formatsettings.ShortDateFormat:='ddmmyy';
  fname:=DateToStr(d);
  dir_name:=copy(fname,3,4);
  if FileExists(path+dir_name+'\'+fname+'.'+ext) then begin
     f_si8:=TFileStream.Create(path+dir_name+'\'+fname+'.'+ext, fmOpenRead);
     SetLength(buf_si8, f_si8.Size);
     f_si8.Seek(0,0);
     f_si8.ReadBuffer(buf_si8[0], f_si8.Size);
     f_si8.Free;
     ind_buf_si8:=0;
     dl_buf:=Length(buf_si8);
     while (ind_buf_si8 < dl_buf) do begin
        kol:=buf_si8[ind_buf_si8];
        h:=buf_si8[ind_buf_si8+1];
        m:=buf_si8[ind_buf_si8+2];
        ind_buf_si8:=ind_buf_si8+3;
        for i:=1 to kol do begin
           num:=buf_si8[ind_buf_si8];
           CopyMemory(@pok1, @buf_si8[ind_buf_si8+1], 4);
           ind_buf_si8:=ind_buf_si8+5;
           eqp := GetEquipByAdr(num,ext);
           if eqp<>nil then begin
              if h * 60 + m >= eqp.sm_b then
                eqp.speed[h * 60 + m - eqp.sm_b] := pok1
              else
                eqp.speed[h * 60 + m + eqp.sm_e] := pok1;
              eqp.kol_sd := eqp.kol_sd + pok1;
           end;
        end;
     end;
  end
  else begin
    Application.MessageBox(PChar('Нет файла '+path+dir_name+'\'+fname+'.'+ext),PChar('Предупреждение!'),MB_OK);
  end;
  if FileExists(path+dir_name+'\'+fname+'.lo'+copy(ext,3,1)) then begin
     f_lo8:=TFileStream.Create(path+dir_name+'\'+fname+'.lo'+copy(ext,3,1),fmOpenRead);
     SetLength(buf_si8, f_lo8.Size);
     f_lo8.Seek(0,0);
     f_lo8.ReadBuffer(buf_si8[0], f_lo8.Size);
     f_lo8.Free;
     ind_buf_si8:=0;
     dl_buf:=Length(buf_si8);
     while (ind_buf_si8<dl_buf) do begin
        h:=buf_si8[ind_buf_si8];
        m:=buf_si8[ind_buf_si8+1];
        kol:=buf_si8[ind_buf_si8+2];
        ind_buf_si8:=ind_buf_si8+3;
        for i := 0 to kol - 1 do begin
           num:=buf_si8[ind_buf_si8];
           inc(ind_buf_si8);
           if (i > 0) and (num <> 0) then begin
              eqp := GetEquipByAdr(i,ext);
              if eqp<>nil then begin
                 if copy(eqp.ext,3,1)=copy(ext,3,1) then
                    if h * 60 + m >= eqp.sm_b then eqp.log[h * 60 + m - eqp.sm_b] := num
                    else eqp.log[h * 60 + m + eqp.sm_e] := num;
              end;
           end;
        end;
     end;
  end
  else begin
    Application.MessageBox(PChar('Нет файла '+path+dir_name+'\'+fname+'.lo'+copy(ext,3,1)),PChar('Предупреждение!'),MB_OK);
  end;
end;

procedure THolding.GetStatMonth(path: string; ext: string; d: TDateTime);
var
  f : textfile;
  f_si8, f_lo8: TFileStream;
  kol, h, m, num, l, cur_m: byte;
  i : integer;
  pok1: single;
  fname, b_day, day, mon, y: string;
  adr, x: integer;
  buf_si8: array of byte;
  ind_buf_si8, dl_buf : Longint;
  LastDay : integer;
  eqp : TEquip;
begin
  day := IntToStr(DayOf(d));
  mon := IntToStr(MonthOf(d));
  y := IntToStr(YearOf(d) - 2000);
  LastDay:=DayOf(d);
  if length(day) = 1 then day := '0' + day;
  if length(mon) = 1 then mon := '0' + mon;
  if length(y) = 1 then y := '0' + y;
  for l := 1 to LastDay do begin
    b_day := IntToStr(l);
    if length(b_day) = 1 then b_day := '0' + b_day;
    fname := b_day + mon + y;
    if FileExists(path + mon + y + '\' + fname + '.' + ext) then begin
        f_si8 := TFileStream.Create(path + mon + y + '\' + fname + '.' + ext, fmOpenRead);
        SetLength(buf_si8, f_si8.Size);
        f_si8.Seek(0, 0);
        try
          f_si8.ReadBuffer(buf_si8[0], f_si8.Size);
        except
          on E: Exception do
            ShowMessage(E.Message);
        end;
        f_si8.Free;
        ind_buf_si8 := 0;
        dl_buf:=Length(buf_si8);
        while (ind_buf_si8 < dl_buf) do begin
          kol := buf_si8[ind_buf_si8];
          inc(ind_buf_si8);
          h := buf_si8[ind_buf_si8];
          inc(ind_buf_si8);
          m := buf_si8[ind_buf_si8];
          inc(ind_buf_si8);
          for i := 1 to kol do begin
            num := buf_si8[ind_buf_si8];
            inc(ind_buf_si8);
            CopyMemory(@pok1, @buf_si8[ind_buf_si8], 4);
            ind_buf_si8:=ind_buf_si8+4;
            eqp := GetEquipByAdr(num,ext);
            if eqp<>nil then begin
              if h * 60 + m >= eqp.sm_b then
                eqp.speed_month[l,h * 60 + m - eqp.sm_b]:=pok1
              else
                eqp.speed_month[l,h * 60 + m + eqp.sm_e]:=pok1;
            end;

          end;
        end;
      end;
      if FileExists(path + mon + y + '\' + fname + '.lo' + copy(ext,3,1)) then begin
        f_lo8 := TFileStream.Create(path + mon + y + '\' + fname + '.lo' + copy(ext,3,1), fmOpenRead);
        SetLength(buf_si8, f_lo8.Size);
        f_lo8.Seek(0, 0);
        try
          f_lo8.ReadBuffer(buf_si8[0], f_lo8.Size);
        except
          on E: Exception do
            ShowMessage(E.Message);
        end;
        f_lo8.Free;
        ind_buf_si8 := 0;
        dl_buf:=Length(buf_si8);
        while (ind_buf_si8 < dl_buf) do begin
           h := buf_si8[ind_buf_si8];
           m := buf_si8[ind_buf_si8+1];
           kol := buf_si8[ind_buf_si8+2];
           ind_buf_si8:=ind_buf_si8+3;
           for i:=0 to kol-1 do begin
              num:=buf_si8[ind_buf_si8];
              inc(ind_buf_si8);
              if (i > 0) and (num <> 0) then begin
                 eqp := GetEquipByAdr(i,ext);
                 if eqp<>nil then begin
                    if copy(eqp.ext,3,1)=copy(ext,3,1) then begin
                      if h * 60 + m >= eqp.sm_b then eqp.log_month[l,h * 60 + m - eqp.sm_b]:=num
                      else eqp.log_month[l,h * 60 + m + eqp.sm_e]:=num;
                    end;
                 end;
              end;
           end;
        end;
      end;
  end;
end;

{procedure TEquip.Interpolate(CurMin : integer);
begin

end;}

procedure TEquip.InterpolateMonth(ToDay : integer);
var l, x : integer;
begin
   for l := 1 to ToDay do begin
         for x:=1 to 1439 do begin
            if log_month[l, x]>=3 then speed_month[l,x]:=speed_month[l,x-1];
            if speed_month[l,x]>(maxspd*KoefInterMaxSpeed) then speed_month[l,x]:=speed_month[l,x-1];
            if (log_month[l, x]=2) and (speed_month[l,x]<0.001) then speed_month[l,x]:=speed_month[l,x-1];
            if x<1439 then begin
               if (speed_month[l,x-1]=0) and (speed_month[l,x+1]=0) then speed_month[l,x]:=0;
               if (log_month[l, x]=0) and (log_month[l, x-1]<>0) and (log_month[l, x+1]<>0) then speed_month[l,x]:=speed_month[l,x-1];
            end;
            if needMin and (speed_month[l,x]<minspd) then speed_month[l,x]:=0;
          end;
          if needInter then begin
             if speed_month[l,0]>0 then speed_month[l,0]:=(5*speed_month[l,0]+2*speed_month[l,1]-speed_month[l,2])/6;
             if speed_month[l,1]>0 then speed_month[l,1]:=(speed_month[l,0]+speed_month[l,1]+speed_month[l,2])/3;
             for x:=2 to 1438 do begin
                if speed_month[l,x]>0 then speed_month[l,x]:=(speed_month[l,x-1]+speed_month[l,x]+speed_month[l,x+1])/3;
             end;
             if speed_month[l,1439]>0 then speed_month[l,1439]:=(-1*speed_month[l,1437]+2*speed_month[l,1438]+5*speed_month[l,1439])/6;
          end;
   end;
end;

procedure THolding.GetNumEmp(path: string; d: TDateTime);
var i, j, k : integer;
    dir_name : string;
    jfile : string;
//    jsonToSend : TStringStream;
    Json, {obj,} item: TJSONObject;
    jArr:TJSONArray;
    v:TJSONValue;
//    mon, god : integer;
//    idHTTP1 : TidHTTP;
    res : string;
begin
//  idHTTP1:=TidHTTP.Create;
  Formatsettings.ShortDateFormat:='ddmmyy';
{  mon:=MonthOf(d);
  god:=YearOf(d);
  dir_name:=copy(DateToStr(d),3,4);}

   for i:=0 to Count-1 do begin
      FillChar(ListFactory[i].data_json, SizeOf(ListFactory[i].data_json), 0);
      jfile:=ListFactory[i].json;
      res:=GetJSON(path+dir_name+'\',jfile,d);
      if res<>'{}' then begin
         Json := TJSONObject.ParseJSONValue(res) as TJSONObject;
         jArr := Json.getValue('num') as TJSONArray;
         for j:=0 to jArr.Count-1 do begin
           item:=jarr.Items[j] as TJSONObject;
           v:=item.GetValue(IntToStr(j+1));
           ListFactory[i].data_json[j+1]:=StrToInt(v.ToString);
         end;
         Json.Free;
//         jArr.Free;
      end;
      for j:=0 to ListFactory[i].Count-1 do begin
        FillChar(ListFactory[i].ListCex[j].data_json, SizeOf(ListFactory[i].ListCex[j].data_json), 0);
        jfile:=ListFactory[i].ListCex[j].json;
        res:=GetJSON(path+dir_name+'\',jfile,d);
        if res<>'{}' then begin
           Json := TJSONObject.ParseJSONValue(res) as TJSONObject;
           jArr := Json.getValue('num') as TJSONArray;
           for k:=0 to jArr.Count-1 do begin
             item:=jarr.Items[k] as TJSONObject;
             v:=item.GetValue(IntToStr(k+1));
             ListFactory[i].ListCex[j].data_json[k+1]:=StrToInt(v.ToString);
           end;
           Json.Free;
 //          jArr.Free;
        end;
      end;
   end;
end;

end.
