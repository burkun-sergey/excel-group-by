unit ProcessParamsValidatorUnit;

interface

uses ProcessParamsUnit, System.Classes, System.SysUtils;

function ValidateProcessParams(const params: ProcessParamsRecord): TStringList;

implementation

function ValidateProcessParams(const params: ProcessParamsRecord): TStringList;
begin
  Result := TStringList.Create;

  if not FileExists(params.srcFilename) then
    Result.Add('Исходный файл не найден');

  if not DirectoryExists(ExtractFilePath(params.newFilename)) then
    Result.Add('Папка для создания результирующего файла не существует');

  if params.srcSheetNum <= 0 then
    Result.Add('Неверно указан номер листа с данными');

  if (params.srcRangeType = pprtCustom) and (trim(params.srcRangeAddress) = '') then
    Result.Add('Не указан адрес области');

  if params.srcRangeStartFromRowNum <= 0 then
    Result.Add('Неверно указано начало строки в области данных');

  if params.srcRangeColumnId <= 0 then
    Result.Add('Неверно указан номер колонки идентификатора в области данных');

  if params.srcRangeColumnBarcode <= 0 then
    Result.Add('Неверно указан номер колонки штрихкода в области данных');

  if (params.srcRangeColumnId > 0) and (params.srcRangeColumnBarcode > 0) and (params.srcRangeColumnId = params.srcRangeColumnBarcode) then
    Result.Add('Колонка и штрихкода и идентификатора не должны совпадать');
end;

end.
