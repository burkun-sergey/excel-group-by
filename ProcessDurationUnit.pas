unit ProcessDurationUnit;

interface

uses
  Classes, SysUtils, DateUtils;

type

  { TProcessDuration }

  // ����� ��� �������� ���������� �� ��������
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

// �������� ����� ������ ��������
procedure TProcessDuration.start();
begin
  startTimestamp := now;
end;

// ������ ����� ������ �������� � ���������� ���������� ����� � ���
procedure TProcessDuration.stop();
begin
  startTimestamp := 0;
  stepsCount := 0;
end;

// ��������� ���� ���������� ��� ��������
procedure TProcessDuration.addStep();
begin
  Inc(stepsCount);
end;

// ���������� ������� ���������� ����� ������� � �������
function TProcessDuration.getOPS(): real;
var passedTimeInSeconds: Int64;
begin
  passedTimeInSeconds := getPassedTime();

  if passedTimeInSeconds = 0 then
    Result := 0
  else Result := stepsCount / passedTimeInSeconds;
end;

// ���������� ���������� ����� � ��������
function TProcessDuration.getRemainingTime(const stepCountRemain: integer): Int64;
begin
  if stepCountRemain < 0 then
    Result := 0
  else Result := round(getOPS * stepCountRemain);
end;

// ���������� ���������� ������ � ������ ��������
function TProcessDuration.getPassedTime(): Int64;
begin
  Result := SecondsBetween(startTimestamp, now);
end;

end.
