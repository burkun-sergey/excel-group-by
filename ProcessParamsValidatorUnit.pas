unit ProcessParamsValidatorUnit;

interface

uses ProcessParamsUnit, System.Classes, System.SysUtils;

function ValidateProcessParams(const params: ProcessParamsRecord): TStringList;

implementation

function ValidateProcessParams(const params: ProcessParamsRecord): TStringList;
begin
  Result := TStringList.Create;

  if not FileExists(params.srcFilename) then
    Result.Add('�������� ���� �� ������');

  if not DirectoryExists(ExtractFilePath(params.newFilename)) then
    Result.Add('����� ��� �������� ��������������� ����� �� ����������');

  if params.srcSheetNum <= 0 then
    Result.Add('������� ������ ����� ����� � �������');

  if (params.srcRangeType = pprtCustom) and (trim(params.srcRangeAddress) = '') then
    Result.Add('�� ������ ����� �������');

  if params.srcRangeStartFromRowNum <= 0 then
    Result.Add('������� ������� ������ ������ � ������� ������');

  if params.srcRangeColumnId <= 0 then
    Result.Add('������� ������ ����� ������� �������������� � ������� ������');

  if params.srcRangeColumnBarcode <= 0 then
    Result.Add('������� ������ ����� ������� ��������� � ������� ������');

  if (params.srcRangeColumnId > 0) and (params.srcRangeColumnBarcode > 0) and (params.srcRangeColumnId = params.srcRangeColumnBarcode) then
    Result.Add('������� � ��������� � �������������� �� ������ ���������');
end;

end.
