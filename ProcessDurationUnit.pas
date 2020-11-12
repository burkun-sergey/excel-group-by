﻿unit ProcessDurationUnit;

interface

uses
  Classes, SysUtils, DateUtils;

type

  { TProcessDuration }

  // Класс для подсчета статистики по операции
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

// Засекает время начала операции
procedure TProcessDuration.start();
begin
  startTimestamp := now;
end;

// Чистит время начала операции и количество пройденных шагов в ней
procedure TProcessDuration.stop();
begin
  startTimestamp := 0;
  stepsCount := 0;
end;

// Добавляет один пройденный шаг операции
procedure TProcessDuration.addStep();
begin
  Inc(stepsCount);
end;

// Возвращает среднее количество шагов оперции в секунду
function TProcessDuration.getOPS(): real;
var passedTimeInSeconds: Int64;
begin
  passedTimeInSeconds := getPassedTime();

  if passedTimeInSeconds = 0 then
    Result := 0
  else Result := stepsCount / passedTimeInSeconds;
end;

// Возвращает оставшееся время в секундах
function TProcessDuration.getRemainingTime(const stepCountRemain: integer): Int64;
begin
  if stepCountRemain < 0 then
    Result := 0
  else Result := round(getOPS * stepCountRemain);
end;

// Возвращает количество секунд с начала операции
function TProcessDuration.getPassedTime(): Int64;
begin
  Result := SecondsBetween(startTimestamp, now);
end;

end.
