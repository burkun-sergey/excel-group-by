unit ProcessThreadUnit;

interface

uses Classes, ProcessParamsUnit, SysUtils, ComObj, Excel2000,
     Variants;

const
  EXCEL_APPLICATION = 'Excel.Application';

type

  TProcessThread = class(TThread)
  private
    fParams: ProcessParamsRecord;
    fRowsToProcess: integer;
    fRowsProcessed: integer;
    fException: Exception;
    procedure SetNextStep;
    procedure SetStartProcessInf;
    procedure SetStopProcessInf;
    procedure SetExceptionInf;
    procedure CloseWorkbookIfOpened(excelApp: Variant; const filename: string);
  protected
    procedure Execute; override;
    procedure FakeExecute;
    procedure RealExecute;
  public
    constructor Create(const createsuspended: boolean; params: ProcessParamsRecord);
  end;

implementation

{ TProcessThread }

procedure TProcessThread.CloseWorkbookIfOpened(excelApp: Variant; const filename: string);
var i: integer;
begin
  for i := 1 to excelApp.WorkBooks.Count do
    begin
      if AnsiCompareText(excelApp.WorkBooks[i].FullName, filename) = 0 then
        begin
          excelApp.WorkBooks[i].Close;
          break;
        end;
    end;
end;

constructor TProcessThread.Create(const createsuspended: boolean; params: ProcessParamsRecord);
begin
  inherited Create(createsuspended);

  fParams := params;
end;

// Основная операция преобразования данных исходного документа с сохранением результата в другой файл
procedure TProcessThread.Execute;
begin
  inherited;

// раскомментировать FakeExecute только для тестирования (при этом RealExecute должен быть закомментирован)
//  FakeExecute;

  try
    RealExecute;
  except
    on e: Exception do
      begin
        fException := e;
        Synchronize(SetExceptionInf);
      end;
  end;

end;

// Метод-заглушка для отработки взаимодействия главного потока программы с текущим потоком
procedure TProcessThread.FakeExecute;
begin
  fRowsToProcess := 10000;
  fRowsProcessed := 0;

  Synchronize(SetStartProcessInf);

  repeat
    Inc(fRowsProcessed);
    Synchronize(SetNextStep);

    sleep(300);
  until Terminated;

  Synchronize(SetStopProcessInf);
end;

procedure TProcessThread.RealExecute;
var ExcelApp,
    ActiveWorkBook,
    NewWorkBook,
    UsedRange,
    FoundRange: Variant;

    rowFrom,
    rowCount,
    colFrom,
    i,
    newLastInsertedRow: integer;

    stId,
    stBarcode,
    tmp: string;
begin
  newLastInsertedRow := 0;

  try
    ExcelApp := GetActiveOleObject(EXCEL_APPLICATION);
  except
    on Exception do
      Begin
        try
          ExcelApp := CreateOleObject(EXCEL_APPLICATION);
        except
          on Exception do
            Begin
              raise Exception.create('Не удается открыть Microsoft Excel!'#13'Возможно он не установлен.');
            End;
        end;
      End;
  End;

  ExcelApp.visible := False;
  ExcelApp.ScreenUpdating := False;

  ExcelApp.WorkBooks.Open(fParams.srcFilename);
  ActiveWorkBook := ExcelApp.ActiveWorkBook;

  ExcelApp.WorkBooks.Add;
  NewWorkBook := ExcelApp.ActiveWorkBook;

  if fParams.srcRangeType = pprtAuto then
    UsedRange := ActiveWorkBook.WorkSheets[1].UsedRange
  else UsedRange := ActiveWorkBook.WorkSheets[1].Range[fParams.srcRangeAddress];

  rowFrom := UsedRange.Row;
  rowCount := UsedRange.Rows.Count;
  colFrom := UsedRange.Column;

  fRowsToProcess := rowCount;
  fRowsProcessed := 0;

  Synchronize(SetStartProcessInf);

  for i := rowFrom to rowFrom + rowCount - 1 do
    begin
      stId := ActiveWorkBook.WorkSheets[1].Cells[i, colFrom].Text;
      stBarcode := ActiveWorkBook.WorkSheets[1].Cells[i, colFrom+1].Text;

      if trim(stId) = '' then
        break;

      FoundRange := NewWorkBook.WorkSheets[1].Columns[1].Find(What := stId, LookAt := xlWhole);
      if (VarIsClear(FoundRange)) then
        begin
          Inc(newLastInsertedRow);
          NewWorkBook.WorkSheets[1].Cells[newLastInsertedRow, 1].Value := stId;
          NewWorkBook.WorkSheets[1].Cells[newLastInsertedRow, 2].Value := stBarcode;
        end
      else
        begin
          tmp := NewWorkBook.WorkSheets[1].Cells[FoundRange.Row, 2].Text;
          NewWorkBook.WorkSheets[1].Cells[FoundRange.Row, 2].Value := tmp + ';' + stBarcode;
        end;

      Inc(fRowsProcessed);
      Synchronize(SetNextStep);

      if Terminated then
        break;
    end;

  if not Terminated then
    begin
      ExcelApp.DisplayAlerts := false;
      CloseWorkbookIfOpened(ExcelApp, fParams.newFilename);
      NewWorkBook.SaveAs(fParams.newFilename);
      ExcelApp.DisplayAlerts := true;
    end;

  NewWorkBook.Close;
  ActiveWorkBook.Close;
  ExcelApp.ScreenUpdating := True;

  if (ExcelApp.WorkBooks.Count = 0) then
    ExcelApp.Quit;

  Synchronize(SetStopProcessInf);

  FoundRange := unassigned;
  UsedRange := unassigned;
  NewWorkBook := unassigned;
  ActiveWorkBook := unassigned;
  ExcelApp := unassigned;
end;

procedure TProcessThread.SetExceptionInf;
begin
  fParams.operationInf^.exceptionMessage := fException.Message;
  fParams.operationInf^.terminated := true;
end;

procedure TProcessThread.SetNextStep;
begin
  fParams.operationInf^.stepNum := fRowsProcessed;

  if Assigned(fParams.durationObj) then
    begin
      fParams.durationObj.addStep;
      fParams.operationInf^.ops := fParams.durationObj.getOPS;
      fParams.operationInf^.timePassedInSeconds := fParams.durationObj.getPassedTime;
      fParams.operationInf^.timeRemainInSeconds := fParams.durationObj.getRemainingTime(fRowsToProcess - fRowsProcessed);
    end;
end;

procedure TProcessThread.SetStartProcessInf;
begin
  if Assigned(fParams.durationObj) then
    fParams.durationObj.start;

  fParams.operationInf^.stepsCount := fRowsToProcess;
  fParams.operationInf^.stepNum := 1;

  if Assigned(fParams.durationObj) then
    begin
      fParams.operationInf^.ops := fParams.durationObj.getOPS;
      fParams.operationInf^.timePassedInSeconds := fParams.durationObj.getPassedTime;
      fParams.operationInf^.timeRemainInSeconds := fParams.durationObj.getRemainingTime(fRowsToProcess - fRowsProcessed);
    end;
end;

procedure TProcessThread.SetStopProcessInf;
begin
  if Assigned(fParams.durationObj) then
    fParams.durationObj.stop;

  fParams.operationInf^.terminated := Terminated;
end;

end.
