unit Exporter;

interface

uses
  System.SysUtils, Vcl.ComCtrls, Vcl.StdCtrls;

procedure ExportSplit(const IniPath: string; AProgress: TProgressBar; AMemo: TMemo = nil);

implementation

uses
  System.Classes, System.Variants, System.IOUtils,
  Winapi.ActiveX, Winapi.Windows, Winapi.Messages,
  IniFiles, ComObj, System.DateUtils,
  Vcl.Forms;

const
  COL_JOB_FIRST     = 1;   // A
  COL_JOB_LAST      = 5;   // E
  COL_PART_FIRST    = 6;   // F
  COL_PART_LAST     = 9;   // I
  COL_PROCESS_FIRST = 20;  // T
  COL_PROCESS_LAST  = 77;  // BY

  COL_PO            = 3;   // C (PO.No.) - เปลี่ยนให้ดึงจากช่อง C
  COL_F             = 6;   // F
  COL_O             = 15;  // O (StartDate)

{-------------------- Helpers --------------------}

procedure MemoStep(AMemo: TMemo; const S: string);
begin
  if Assigned(AMemo) then
  begin
    AMemo.Lines.Add(FormatDateTime('hh:nn:ss', Now) + '  ' + S);
    AMemo.SelStart := Length(AMemo.Text);
    AMemo.Perform(EM_SCROLLCARET, 0, 0);
    Application.ProcessMessages;
  end;
end;

function CsvEscape(const S: string): string;
var
  NeedsQuote: Boolean;
  R: string;
begin
  NeedsQuote := (Pos(',', S) > 0) or (Pos('"', S) > 0) or
                (Pos(#13, S) > 0) or (Pos(#10, S) > 0);
  R := StringReplace(S, '"', '""', [rfReplaceAll]);
  if NeedsQuote then
    Result := '"' + R + '"'
  else
    Result := R;
end;

function CsvJoin(const Fields: TArray<string>): string;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to High(Fields) do
  begin
    if i > 0 then Result := Result + ',';
    Result := Result + CsvEscape(Fields[i]);
  end;
end;

// ฟังก์ชันดึงค่าจาก Array แบบปลอดภัย (ป้องกัน Error Invalid Index)
function GetSafeVal(const DataArray: Variant; Row, Col, MaxCol: Integer): string;
begin
  if Col <= MaxCol then
    Result := VarToStr(DataArray[Row, Col])
  else
    Result := ''; // ถ้าคอลัมน์เกินกว่าที่มีในไฟล์ ให้คืนค่าว่าง
end;

// ฟังก์ชันสำหรับอ่านจาก Variant Array พร้อมส่ง MaxCol เข้าไปตรวจสอบ
function CsvLineFromArrayCols(const DataArray: Variant; const Row: Integer; const ColIdx: array of Integer; const MaxCol: Integer): string;
var
  i, C: Integer;
begin
  Result := '';
  for i := Low(ColIdx) to High(ColIdx) do
  begin
    C := ColIdx[i];
    if i > Low(ColIdx) then
      Result := Result + ',';

    Result := Result + CsvEscape(GetSafeVal(DataArray, Row, C, MaxCol));
  end;
end;

{-------------------- CSV Writers --------------------}

procedure SaveCSVJob(const DataArray: Variant; const MaxRow, MaxCol: Integer;
  const Path, LogPath: string; AProgress: TProgressBar; AMemo: TMemo);
var
  R, DummyIdx: Integer;
  SL, AddedPOs: TStringList;
  HeaderFields: TArray<string>;
  LineData, SDate, PoNo: string;
  ExcelDate: Variant;
begin
  MemoStep(AMemo, 'Start JOB → ' + Path);
  SL := TStringList.Create;
  AddedPOs := TStringList.Create;
  try
    // ตั้งให้ค้นหา PO ที่เคยใส่ไปแล้วได้ไวขึ้น
    AddedPOs.Sorted := True;

    // Header เหลือแค่ 4 คอลัมน์
    HeaderFields := TArray<string>.Create('CstmrCD','Cstmr.Name','PO.No.','StartDate');
    SL.Add(CsvJoin(HeaderFields));

    if Assigned(AProgress) then
    begin
      AProgress.Position := 0;
      AProgress.Max := MaxRow - 2;
    end;

    for R := 3 to MaxRow do
    begin
      // ตรวจสอบ PO No. ว่าซ้ำหรือไม่
      PoNo := GetSafeVal(DataArray, R, COL_PO, MaxCol);
      if (Trim(PoNo) = '') or AddedPOs.Find(PoNo, DummyIdx) then
      begin
        if Assigned(AProgress) then AProgress.Position := R - 2;
        Continue; // ข้ามบรรทัดนี้ไปเลยถ้า PO ซ้ำ
      end;

      // จดจำว่า PO นี้เพิ่มลงไปแล้ว
      AddedPOs.Add(PoNo);

      // ดึงข้อมูล 3 คอลัมน์แรก (A, B, C)
      LineData := CsvLineFromArrayCols(DataArray, R, [1, 2, COL_PO], MaxCol);

      SDate := '';
      if COL_O <= MaxCol then
      begin
        ExcelDate := DataArray[R, COL_O];
        if not VarIsNull(ExcelDate) and not VarIsEmpty(ExcelDate) then
        begin
          try
            SDate := FormatDateTime('dd/mm/yyyy', IncDay(VarToDateTime(ExcelDate), 3));
          except
            SDate := '';
          end;
        end;
      end;

      // เอาข้อมูลต่อกันแล้วใส่ลงใน CSV
      SL.Add(LineData + ',' + CsvEscape(SDate));

      if Assigned(AProgress) then AProgress.Position := R - 2;
      if (R mod 1000 = 0) then
        MemoStep(AMemo, Format('JOB processing row %d...', [R]));
    end;

    ForceDirectories(ExtractFileDir(Path));
    SL.SaveToFile(Path, TEncoding.UTF8);

    TFile.AppendAllText(LogPath, Format('[%s] JOB -> %s (header + %d unique rows)' + sLineBreak,
      [FormatDateTime('yyyy-mm-dd hh:nn:ss', Now), Path, SL.Count - 1]), TEncoding.UTF8);
    MemoStep(AMemo, Format('JOB completed (%d unique rows written)', [SL.Count - 1]));
  finally
    AddedPOs.Free;
    SL.Free;
  end;
end;

procedure SaveCSVPart(const DataArray: Variant; const MaxRow, MaxCol: Integer;
  const Path, LogPath: string; AProgress: TProgressBar; AMemo: TMemo);
var
  R: Integer;
  SL: TStringList;
  HeaderFields: TArray<string>;
begin
  MemoStep(AMemo, 'Start PART → ' + Path);
  SL := TStringList.Create;
  try
    // ✅ 1. แทรกชื่อ Header: Drawing No. (เดิมคือ RE) และ Ref No. (เดิมคือ ProductName)
    HeaderFields := TArray<string>.Create('PO.No.', 'Drawing No.', 'Ref No.', 'PartsName', 'Material', 'SizeRemarks', 'PartsQuantity');
    SL.Add(CsvJoin(HeaderFields));

    if Assigned(AProgress) then
    begin
      AProgress.Position := 0;
      AProgress.Max := MaxRow - 2;
    end;

    for R := 3 to MaxRow do
    begin
      // ✅ 2. เพิ่มคอลัมน์ 4 (RE) และ 5 (ProductName) เข้าไปต่อจาก COL_PO (3)
      SL.Add(CsvLineFromArrayCols(DataArray, R, [COL_PO, 4, 5, 6, 7, 8, 9], MaxCol));

      if Assigned(AProgress) then AProgress.Position := R - 2;
      if (R mod 1000 = 0) then
        MemoStep(AMemo, Format('PART processing row %d...', [R]));
    end;

    ForceDirectories(ExtractFileDir(Path));
    SL.SaveToFile(Path, TEncoding.UTF8);

    TFile.AppendAllText(LogPath, Format('[%s] PART -> %s (header + %d rows)' + sLineBreak,
      [FormatDateTime('yyyy-mm-dd hh:nn:ss', Now), Path, SL.Count - 1]), TEncoding.UTF8);
    MemoStep(AMemo, Format('PART completed (%d rows written)', [SL.Count - 1]));
  finally
    SL.Free;
  end;
end;

procedure SaveCSVProcess(const DataArray: Variant; const MaxRow, MaxCol: Integer;
  const Path, LogPath: string; AProgress: TProgressBar; AMemo: TMemo);
var
  R, k: Integer;
  SL: TStringList;
  PoNo, PartFig, ProcName, SetVal, MaVal: string;
  TripletStartCol, TripletEndCol, MaxTriples: Integer;
  ProcCol, SetCol, MaCol: Integer;
begin
  MemoStep(AMemo, 'Start PROCESS → ' + Path);
  SL := TStringList.Create;
  try
    SL.Add(CsvJoin(TArray<string>.Create('PO.No.','Part figure','process','set','ma')));

    TripletStartCol  := COL_PROCESS_FIRST + 1;
    TripletEndCol    := COL_PROCESS_LAST;

    if TripletEndCol <= MaxCol then
      MaxTriples := (TripletEndCol - TripletStartCol + 1) div 3
    else
      MaxTriples := (MaxCol - TripletStartCol + 1) div 3;

    if MaxTriples < 0 then MaxTriples := 0;

    if Assigned(AProgress) then
    begin
      AProgress.Position := 0;
      AProgress.Max := MaxRow - 2;
    end;

    for R := 3 to MaxRow do
    begin
      // ดึงผ่าน GetSafeVal ทุกจุดเพื่อความปลอดภัย
      PoNo    := GetSafeVal(DataArray, R, COL_PO, MaxCol);
      PartFig := GetSafeVal(DataArray, R, COL_F, MaxCol);

      if COL_PROCESS_FIRST <= MaxCol then
      begin
        ProcName := GetSafeVal(DataArray, R, COL_PROCESS_FIRST, MaxCol);
        if Trim(ProcName) <> '' then
          SL.Add(CsvJoin(TArray<string>.Create(PoNo, PartFig, ProcName, '0', '0')));
      end;

      for k := 0 to MaxTriples - 1 do
      begin
        ProcCol := TripletStartCol + (k * 3);
        SetCol  := ProcCol + 1;
        MaCol   := ProcCol + 2;

        if MaCol > MaxCol then Break;

        ProcName := GetSafeVal(DataArray, R, ProcCol, MaxCol);
        SetVal   := GetSafeVal(DataArray, R, SetCol, MaxCol);
        MaVal    := GetSafeVal(DataArray, R, MaCol, MaxCol);

        if Trim(SetVal) = '' then SetVal := '0';
        if Trim(MaVal)  = '' then MaVal  := '0';

        if Trim(ProcName) <> '' then
          SL.Add(CsvJoin(TArray<string>.Create(PoNo, PartFig, ProcName, SetVal, MaVal)));
      end;

      if Assigned(AProgress) then AProgress.Position := R - 2;
      if (R mod 1000 = 0) then
        MemoStep(AMemo, Format('PROCESS processing row %d...', [R]));
    end;

    ForceDirectories(ExtractFileDir(Path));
    SL.SaveToFile(Path, TEncoding.UTF8);

    TFile.AppendAllText(LogPath, Format('[%s] PROCESS -> %s (header + %d rows)' + sLineBreak,
      [FormatDateTime('yyyy-mm-dd hh:nn:ss', Now), Path, SL.Count - 1]), TEncoding.UTF8);
    MemoStep(AMemo, Format('PROCESS completed (%d rows)', [SL.Count - 1]));
  finally
    SL.Free;
  end;
end;

{-------------------- MAIN --------------------}

procedure ExportSplit(const IniPath: string; AProgress: TProgressBar; AMemo: TMemo = nil);
var
  Ini: TMemIniFile;
  InputFile, SheetName: string;
  OutJob, OutPart, OutProcess: string;
  LogPath, LogDir: string;
  Excel, WB, Sheet: OleVariant;
  DataArray: Variant;
  MaxRow, MaxCol: Integer;
begin
  Ini := TMemIniFile.Create(IniPath, TEncoding.UTF8);
  try
    InputFile := Ini.ReadString('Input','File','');
    // ดึงชื่อชีตจาก INI (ป้องกัน Error กรณีไฟล์ใช้ชื่อชีตอื่น)
    SheetName := Ini.ReadString('Input','Sheet','Sheet1');
    OutJob := Ini.ReadString('Output1','Path','');
    OutPart := Ini.ReadString('Output2','Path','');
    OutProcess := Ini.ReadString('Output3','Path','');
    LogPath := Ini.ReadString('Options','LogPath','');
    if Trim(LogPath)='' then
    begin
      LogDir := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0))) + 'LOG\';
      if not TDirectory.Exists(LogDir) then TDirectory.CreateDirectory(LogDir);
      LogPath := TPath.Combine(LogDir, 'export_log_' + FormatDateTime('yyyymmdd', Now) + '.txt');
    end;
  finally
    Ini.Free;
  end;

  if not FileExists(InputFile) then
    raise Exception.CreateFmt('Input file not found: %s', [InputFile]);

  CoInitialize(nil);
  try
    MemoStep(AMemo, 'Starting Excel Application...');
    Excel := CreateOleObject('Excel.Application');
    Excel.Visible := False;

    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := False;

    WB := Excel.Workbooks.Open(InputFile, False, True);
    try
      Sheet := WB.Worksheets[SheetName];
      MemoStep(AMemo, 'Reading all data from sheet: ' + SheetName + ' into memory (fast & safe mode)...');

      DataArray := Sheet.UsedRange.Value;
      MaxRow := VarArrayHighBound(DataArray, 1);
      MaxCol := VarArrayHighBound(DataArray, 2);

      MemoStep(AMemo, Format('Data loaded. Total Rows: %d, Total Cols: %d', [MaxRow, MaxCol]));

      if MaxRow >= 3 then
      begin
        if OutJob <> '' then SaveCSVJob(DataArray, MaxRow, MaxCol, OutJob, LogPath, AProgress, AMemo);
        if OutPart <> '' then SaveCSVPart(DataArray, MaxRow, MaxCol, OutPart, LogPath, AProgress, AMemo);
        if OutProcess <> '' then SaveCSVProcess(DataArray, MaxRow, MaxCol, OutProcess, LogPath, AProgress, AMemo);
      end
      else
        MemoStep(AMemo, 'Not enough data rows to process.');

      MemoStep(AMemo, 'Export completed successfully.');
    finally
      WB.Close(False);
      Excel.Quit;
    end;
  finally
    CoUninitialize;
  end;
end;

end.
