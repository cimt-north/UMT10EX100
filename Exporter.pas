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

  COL_M             = 13;  // M (PO.No.)
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

// ฟังก์ชันสำหรับอ่านจาก Variant Array (ทำงานเร็วกว่าอ่านทีละ Cell)
function CsvLineFromArrayCols(const DataArray: Variant; const Row: Integer; const ColIdx: array of Integer): string;
var
  i: Integer;
  V: Variant;
begin
  Result := '';
  for i := Low(ColIdx) to High(ColIdx) do
  begin
    V := DataArray[Row, ColIdx[i]];
    if i > Low(ColIdx) then
      Result := Result + ',';
    Result := Result + CsvEscape(VarToStr(V));
  end;
end;

{-------------------- CSV Writers --------------------}

procedure SaveCSVJob(const DataArray: Variant; const MaxRow: Integer;
  const Path, LogPath: string; AProgress: TProgressBar; AMemo: TMemo);
var
  R: Integer;
  SL: TStringList;
  HeaderFields: TArray<string>;
  LineData, SDate: string;
  ExcelDate: Variant;
begin
  MemoStep(AMemo, 'Start JOB → ' + Path);
  SL := TStringList.Create;
  try
    // ✅ เปลี่ยน Mfg.No. เป็น PO.No.
    HeaderFields := TArray<string>.Create('CstmrCD','Cstmr.Name','PO.No.','RE','ProductName','StartDate');
    SL.Add(CsvJoin(HeaderFields));

    if Assigned(AProgress) then
    begin
      AProgress.Position := 0;
      AProgress.Max := MaxRow - 2; // เริ่มข้อมูลที่แถว 3
    end;

    for R := 3 to MaxRow do
    begin
      // ดึงคอลัมน์ [1, 2, 13, 4, 5]
      LineData := CsvLineFromArrayCols(DataArray, R, [1, 2, COL_M, 4, 5]);

      // ดึงวันที่จาก Column O มาคำนวณ +3 วัน ทันที (ไม่ต้องเขียนแล้วเปิดใหม่)
      ExcelDate := DataArray[R, COL_O];
      SDate := '';
      if not VarIsNull(ExcelDate) and not VarIsEmpty(ExcelDate) then
      begin
        try
          SDate := FormatDateTime('dd/mm/yyyy', IncDay(VarToDateTime(ExcelDate), 3));
        except
          SDate := '';
        end;
      end;

      SL.Add(LineData + ',' + CsvEscape(SDate));

      if Assigned(AProgress) then AProgress.Position := R - 2;
      if (R mod 1000 = 0) then
        MemoStep(AMemo, Format('JOB processing row %d...', [R]));
    end;

    ForceDirectories(ExtractFileDir(Path));
    SL.SaveToFile(Path, TEncoding.UTF8);

    TFile.AppendAllText(LogPath, Format('[%s] JOB -> %s (header + %d rows)' + sLineBreak,
      [FormatDateTime('yyyy-mm-dd hh:nn:ss', Now), Path, SL.Count - 1]), TEncoding.UTF8);
    MemoStep(AMemo, Format('JOB completed (%d rows written)', [SL.Count - 1]));
  finally
    SL.Free;
  end;
end;

procedure SaveCSVPart(const DataArray: Variant; const MaxRow: Integer;
  const Path, LogPath: string; AProgress: TProgressBar; AMemo: TMemo);
var
  R: Integer;
  SL: TStringList;
  HeaderFields: TArray<string>;
begin
  MemoStep(AMemo, 'Start PART → ' + Path);
  SL := TStringList.Create;
  try
    // ✅ เปลี่ยน Mfg.No. เป็น PO.No.
    HeaderFields := TArray<string>.Create('PO.No.','PartsName','Material','SizeRemarks','PartsQuantity');
    SL.Add(CsvJoin(HeaderFields));

    if Assigned(AProgress) then
    begin
      AProgress.Position := 0;
      AProgress.Max := MaxRow - 2;
    end;

    for R := 3 to MaxRow do
    begin
      SL.Add(CsvLineFromArrayCols(DataArray, R, [COL_M, 6, 7, 8, 9]));

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
    // ✅ เปลี่ยน Mfg.No. เป็น PO.No.
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
      PoNo    := VarToStr(DataArray[R, COL_M]);
      PartFig := VarToStr(DataArray[R, COL_F]);

      // ตัวแรกสุด
      if COL_PROCESS_FIRST <= MaxCol then
      begin
        ProcName := VarToStr(DataArray[R, COL_PROCESS_FIRST]);
        if Trim(ProcName) <> '' then
          SL.Add(CsvJoin(TArray<string>.Create(PoNo, PartFig, ProcName, '0', '0')));
      end;

      // ตัวที่เหลือเป็น Triplet
      for k := 0 to MaxTriples - 1 do
      begin
        ProcCol := TripletStartCol + (k * 3);
        SetCol  := ProcCol + 1;
        MaCol   := ProcCol + 2;

        if MaCol > MaxCol then Break; // ป้องกัน index out of bounds

        ProcName := VarToStr(DataArray[R, ProcCol]);
        SetVal   := VarToStr(DataArray[R, SetCol]);
        MaVal    := VarToStr(DataArray[R, MaCol]);

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
    SheetName := 'ピックアップ';
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

    // ป้องกัน Popup กวนใจและช่วยให้ประมวลผลเร็วขึ้น
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := False;

    WB := Excel.Workbooks.Open(InputFile, False, True); // เปิดแบบ Read-Only
    try
      Sheet := WB.Worksheets[SheetName];
      MemoStep(AMemo, 'Reading all data from sheet: ' + SheetName + ' into memory (fast mode)...');

      // ✅ ดึงข้อมูลเข้า RAM รวดเดียวด้วย Variant Array
      DataArray := Sheet.UsedRange.Value;
      MaxRow := VarArrayHighBound(DataArray, 1);
      MaxCol := VarArrayHighBound(DataArray, 2);

      MemoStep(AMemo, Format('Data loaded. Total Rows: %d, Total Cols: %d', [MaxRow, MaxCol]));

      if MaxRow >= 3 then
      begin
        if OutJob <> '' then SaveCSVJob(DataArray, MaxRow, OutJob, LogPath, AProgress, AMemo);
        if OutPart <> '' then SaveCSVPart(DataArray, MaxRow, OutPart, LogPath, AProgress, AMemo);
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
