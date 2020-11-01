unit ProcessDurationUnit;

interface

uses
  Classes, SysUtils, DateUtils;

type

  { TProcessDuration }

  //  ласс дл€ подсчета статистики по операции
  TProcessDuration = class(TObject)
  private
    startTimestamp: TDateTime;
    stepsCount: integer;
  public
    procedure start();
    procedure stop();
    procedure addStep();
    function getOPS(): real;
    function getRemainingTime(const stepCountRemain: integer): Int64;
    function getPassedTime(): Int64;
  end;

implementation

{ TProcessDuration }

// «асекает врем€ начала операции
procedure TProcessDuration.start();
begin
  startTimestamp := now;
end;

// „истит врем€ начала операции и количество пройденных шагов в ней
procedure TProcessDuration.stop();
begin
  startTimestamp := 0;
  stepsCount := 0;
end;

// ƒобавл€ет один пройденный шаг операции
procedure TProcessDuration.addStep();
begin
  Inc(stepsCount);
end;

// ¬озвращает среднее количество шагов оперции в секунду
function TProcessDuration.getOPS(): real;
var passedTimeInSeconds: Int64;
begin
  passedTimeInSeconds := getPassedTime();

  if passedTimeInSeconds = 0 then
    Result := 0
  else Result := stepsCount / passedTimeInSeconds;
end;

// ¬озвращает оставшеес€ врем€ в секундах
function TProcessDuration.getRemainingTime(const stepCountRemain: integer): Int64;
begin
  if stepCountRemain < 0 then
    Result := 0
  else Result := round(getOPS * stepCountRemain);
end;

// ¬озвращает количество секунд с начала операции
function TProcessDuration.getPassedTime(): Int64;
begin
  Result := SecondsBetween(startTimestamp, now);
end;

end.
