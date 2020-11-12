unit MainUnit;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.Samples.Spin, Vcl.ExtCtrls, ProcessThreadUnit, ProcessDurationUnit,
  ProcessParamsValidatorUnit, ProcessParamsUnit, IniFiles;

const
  OPTIONS_FILENAME = 'options.ini';

type

  TMainForm = class(TForm)
    Label1: TLabel;
    edSrc: TEdit;
    bSrcFilenameBrowse: TButton;
    Label2: TLabel;
    edNew: TEdit;
    bNewFilenameBrowse: TButton;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    bStartStop: TButton;
    pBar: TProgressBar;
    Label3: TLabel;
    seOptionSrcSheetIndex: TSpinEdit;
    OpenDialogXLS: TOpenDialog;
    SaveDialogXLS: TSaveDialog;
    TimerViewInf: TTimer;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    lInfRowNum: TLabel;
    lInfOPS: TLabel;
    lInfPassed: TLabel;
    lInfRemain: TLabel;
    Label15: TLabel;
    cbOptionSrcRange: TComboBox;
    Label16: TLabel;
    edOptionSrcRangeAddress: TEdit;
    GroupBox3: TGroupBox;
    Label4: TLabel;
    seOptionSrcStartRowNum: TSpinEdit;
    seOptionSrcColNumId: TSpinEdit;
    Label5: TLabel;
    Label6: TLabel;
    seOptionSrcColNumBarcode: TSpinEdit;
    procedure bSrcFilenameBrowseClick(Sender: TObject);
    procedure bNewFilenameBrowseClick(Sender: TObject);
    procedure TimerViewInfTimer(Sender: TObject);
    procedure cbOptionSrcRangeSelect(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure bStartStopClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    processStatistic: OperationInfRecord;
    processParams: ProcessParamsRecord;
    processDuration: TProcessDuration;
    process: TProcessThread;

    procedure RefreshStatistics;
    procedure ClearInf;
    procedure ClearStatistic;
    function FormatSecondsToString(const secs: int64): string;
    procedure FillProcessParams(var processParams: ProcessParamsRecord);
    function GetRangeTypeByComboIndex: ProcessParamRangeType;
    procedure OnTerminateProcessThread(Sender: TObject);
    procedure BeforeProcessStartActions;
    procedure AfterProcessStopActions;
    procedure SaveParams;
    procedure LoadParams;
  public
    { Public declarations }
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.AfterProcessStopActions;
begin
  bStartStop.Caption := 'Начать преобразование';
  pBar.Visible := false;
end;

procedure TMainForm.BeforeProcessStartActions;
begin
  bStartStop.Caption := 'Прервать операцию';
  pBar.Position := 0;
  pBar.Visible := true;
end;

procedure TMainForm.bNewFilenameBrowseClick(Sender: TObject);
begin
  if SaveDialogXLS.Execute then
    edNew.Text := SaveDialogXLS.FileName;
end;

procedure TMainForm.bSrcFilenameBrowseClick(Sender: TObject);
begin
  if OpenDialogXLS.Execute then
    edSrc.Text := OpenDialogXLS.FileName;
end;

procedure TMainForm.bStartStopClick(Sender: TObject);
var errors: TStringList;
begin
  if Assigned(process) then
    begin
      process.Terminate;
    end
  else
    begin
      BeforeProcessStartActions;
      FillProcessParams(processParams);

      errors := ValidateProcessParams(processParams);
      if (errors.Count > 0) then
        begin
          ShowMessage('Обнаружены следующие ошибки: ' + StringReplace(#13#10 + trim(errors.Text), #13#10, #13#10'- ', [rfReplaceAll]) + #13#10'Операция отменена!');
          errors.Free;
          exit;
        end;

      process := TProcessThread.Create(False, processParams);
      process.FreeOnTerminate := true;
      process.OnTerminate := OnTerminateProcessThread;

      TimerViewInf.Enabled := true;
      BringToFront;
    end;
end;

procedure TMainForm.cbOptionSrcRangeSelect(Sender: TObject);
begin
  edOptionSrcRangeAddress.Enabled := cbOptionSrcRange.ItemIndex = 0;
end;

procedure TMainForm.ClearInf;
begin
  lInfRowNum.Caption := '';
  lInfOPS.Caption := '';
  lInfPassed.Caption := '';
  lInfRemain.Caption := '';
end;

procedure TMainForm.ClearStatistic;
begin
  processStatistic.ops := 0;
  processStatistic.stepNum := 0;
  processStatistic.timePassedInSeconds := 0;
  processStatistic.timeRemainInSeconds := 0;
end;

procedure TMainForm.FillProcessParams(var processParams: ProcessParamsRecord);
begin
  processParams.srcFilename := edSrc.Text;
  processParams.newFilename := edNew.Text;
  processParams.srcSheetNum := seOptionSrcSheetIndex.Value;
  processParams.srcRangeType := GetRangeTypeByComboIndex();
  processParams.srcRangeAddress := edOptionSrcRangeAddress.Text;
  processParams.srcRangeStartFromRowNum := seOptionSrcStartRowNum.Value;
  processParams.srcRangeColumnId := seOptionSrcColNumId.Value;
  processParams.srcRangeColumnBarcode := seOptionSrcColNumBarcode.Value;
  processParams.durationObj := processDuration;
  processParams.operationInf := @processStatistic;
end;

function TMainForm.FormatSecondsToString(const secs: int64): string;
begin
  Result := Format('%2.2d:%2.2d:%2.2d',[secs div SecsPerHour,(secs div SecsPerMin) mod SecsPerMin,secs mod SecsPerMin]);
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  ClearInf;
  LoadParams;
  processDuration := TProcessDuration.Create;
  pBar.Visible := false;
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if Assigned(process) then
    begin
      process.Terminate;
    end;

  SaveParams;
  processDuration.Free;
end;

function TMainForm.GetRangeTypeByComboIndex: ProcessParamRangeType;
begin
  case cbOptionSrcRange.ItemIndex of
    0: Result := pprtCustom;
    1: Result := pprtAuto;
    else Result := pprtAuto;
  end;
end;

procedure TMainForm.LoadParams;
var ini: TIniFile;
    filename: string;
begin
  filename := ExtractFilePath(Application.ExeName) + OPTIONS_FILENAME;

  if not FileExists(filename) then
    exit;

  ini := TIniFile.Create(filename);
  seOptionSrcSheetIndex.Value := ini.ReadInteger('Options','SourceSheetIndex',1);
  cbOptionSrcRange.ItemIndex := ini.ReadInteger('Options', 'SourceRangeType', 0);
  edOptionSrcRangeAddress.Text := ini.ReadString('Options', 'SourceRangeAddress', 'A1:B100');
  seOptionSrcStartRowNum.Value := ini.ReadInteger('Options', 'SourceRangeStartRowNum', 1);
  seOptionSrcColNumId.Value := ini.ReadInteger('Options', 'SourceRangeColNumId', 1);
  seOptionSrcColNumBarcode.Value := ini.ReadInteger('Options', 'SourceRangeColNumBarcode', 2);
  ini.Free;

  cbOptionSrcRangeSelect(Self);
end;

procedure TMainForm.OnTerminateProcessThread(Sender: TObject);
var msg: string;
begin
  process := nil;
  TimerViewInf.Enabled := false;
  RefreshStatistics;
  AfterProcessStopActions;

  if not processStatistic.terminated then
    ShowMessage('Готово!')
  else
    begin
      msg := 'Операция отменена!';
      if (processStatistic.exceptionMessage <> '') then
        msg := 'Ошибка: ' + processStatistic.exceptionMessage + #13#10 + msg;

      ShowMessage(msg);
    end;
end;

procedure TMainForm.RefreshStatistics;
begin
  pBar.Max := processStatistic.stepsCount;
  pBar.Position := processStatistic.stepNum;
  lInfRowNum.Caption := IntToStr(processStatistic.stepNum) + ' из ' +IntToStr(processStatistic.stepsCount);
  lInfOPS.Caption := FormatFloat('0.00', processStatistic.ops) + ' строк\сек.';
  lInfPassed.Caption := FormatSecondsToString(processStatistic.timePassedInSeconds);
  lInfRemain.Caption := FormatSecondsToString(processStatistic.timeRemainInSeconds);
end;

procedure TMainForm.SaveParams;
var ini: TIniFile;
begin
  ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + OPTIONS_FILENAME);
  ini.WriteInteger('Options', 'SourceSheetIndex', seOptionSrcSheetIndex.Value);
  ini.WriteInteger('Options', 'SourceRangeType', cbOptionSrcRange.ItemIndex);
  ini.WriteString('Options', 'SourceRangeAddress', edOptionSrcRangeAddress.Text);
  ini.WriteInteger('Options', 'SourceRangeStartRowNum', seOptionSrcStartRowNum.Value);
  ini.WriteInteger('Options', 'SourceRangeColNumId', seOptionSrcColNumId.Value);
  ini.WriteInteger('Options', 'SourceRangeColNumBarcode', seOptionSrcColNumBarcode.Value);
  ini.Free;
end;

procedure TMainForm.TimerViewInfTimer(Sender: TObject);
begin
  RefreshStatistics;
end;

end.
