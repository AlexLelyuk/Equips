unit reports;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Unit1, TeEngine, Series, TeeProcs, Chart,
  VclTee.TeeGDIPlus;

type
  TRep = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Chart1: TChart;
    Series1: TFastLineSeries;
    Series2: TBarSeries;
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    ind : integer;  // индекс в списке оборудования, по которому будет будет делаться расчет
    si8adr : integer;
    info : string;
    ext : string;
  end;

var
  Rep: TRep;

implementation

{$R *.dfm}

procedure TRep.Button2Click(Sender: TObject);
begin
   Close;
end;

procedure TRep.FormShow(Sender: TObject);
begin
   Caption:=Caption+ ' адрес '+ IntToStr(si8adr) + ' оборудование '+info;
   Label1.Caption:=info;
   Label2.Caption:='Адрес СИ8='+ IntToStr(si8adr);
   Label3.Caption:='Расширение файла - '+ext;
end;

end.
