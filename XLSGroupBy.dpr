program XLSGroupBy;

uses
  Vcl.Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  ProcessDurationUnit in 'ProcessDurationUnit.pas',
  ProcessThreadUnit in 'ProcessThreadUnit.pas',
  ProcessParamsUnit in 'ProcessParamsUnit.pas',
  ProcessParamsValidatorUnit in 'ProcessParamsValidatorUnit.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
