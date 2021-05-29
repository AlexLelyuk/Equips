program Equips;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  ToExcel in 'ToExcel.pas',
  common in 'common.pas',
  TLEquips in 'TLEquips.pas',
  Unit2 in 'Unit2.pas' {Form2},
  EquipClass in 'EquipClass.pas',
  reports in 'reports.pas' {Rep},
  Unit3 in 'Unit3.pas' {Form3};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TForm2, Form2);
  Application.CreateForm(TRep, Rep);
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.
