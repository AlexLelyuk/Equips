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
    ind : integer;  // ������ � ������ ������������, �� �������� ����� ����� �������� ������
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
   Caption:=Caption+ ' ����� '+ IntToStr(si8adr) + ' ������������ '+info;
   Label1.Caption:=info;
   Label2.Caption:='����� ��8='+ IntToStr(si8adr);
   Label3.Caption:='���������� ����� - '+ext;
end;

end.
