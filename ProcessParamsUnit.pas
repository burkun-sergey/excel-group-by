unit ProcessParamsUnit;

interface

uses ProcessDurationUnit;

type

  // ���������� �� ��������
  OperationInfRecord = record
    ops: real;
    stepNum,
    stepsCount,
    timePassedInSeconds,
    timeRemainInSeconds: int64;
    terminated: boolean;
    exceptionMessage: string;
  end;

  POperationInfRecord = ^OperationInfRecord;

  ProcessParamRangeType = (pprtCustom, pprtAuto);

  // ��������� ��� ��������
  ProcessParamsRecord = record
    srcFilename,
    newFilename: string;
    srcSheetNum: integer;
    srcRangeType: ProcessParamRangeType;
    srcRangeAddress: string;
    srcRangeStartFromRowNum: integer;
    srcRangeColumnId: integer;
    srcRangeColumnBarcode: integer;
    durationObj: TProcessDuration;
    operationInf: POperationInfRecord;
  end;

implementation

end.
