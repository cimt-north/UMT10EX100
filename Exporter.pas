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

  COL_M             = 13;  // M
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

function CsvLineFromRowCols(const Sheet: OleVariant; const Row: Integer; const ColIdx: array of Integer): string;
var
  i: Integer;
  V: OleVariant;
begin
  Result := '';
  for i := Low(ColIdx) to High(ColIdx) do
  begin
    V := Sheet.Cells[Row, ColIdx[i]].Value;
    if i > Low(ColIdx) then
      Result := Result + ',';
    Result := Result + CsvEscape(VarToStr(V));
  end;
end;

{-------------------- CSV Writers --------------------}

procedure SaveCSVWithCustomHeaderByCols(const SrcSheet: OleVariant;
  const ColIndices: array of Integer; const Path: string;
  const Title: string; const LogPath: string;
  AProgress: TProgressBar; AMemo: TMemo;
  const HeaderFields: TArray<string>);
var
  R: Integer;
  SL: TStringList;
  Total, Done: Integer;
  LastRow, StartRow, EndRow: Integer;
  WroteDataCount: Integer;
begin
  MemoStep(AMemo, 'Start ' + Title + ' → ' + Path);
  SL := TStringList.Create;
  try
    SL.Add(CsvJoin(HeaderFields));

    LastRow  := SrcSheet.UsedRange.Row + SrcSheet.UsedRange.Rows.Count - 1;
    StartRow := 3;
    EndRow   := LastRow;

    if EndRow >= StartRow then
      Total := 1 + (EndRow - StartRow + 1)
    else
      Total := 1;

    Done := 0;
    if Assigned(AProgress) then
    begin
      AProgress.Position := 0;
      AProgress.Max := Total;
    end;

    Inc(Done);
    if Assigned(AProgress) then AProgress.Position := Done;

    WroteDataCount := 0;
    if EndRow >= StartRow then
      for R := StartRow to EndRow do
      begin
        SL.Add(CsvLineFromRowCols(SrcSheet, R, ColIndices));
        Inc(WroteDataCount);
        Inc(Done);
        if Assigned(AProgress) then AProgress.Position := Done;
        if (R mod 50 = 0) then
          MemoStep(AMemo, Format('%s processing row %d...', [Title, R]));
      end;

    ForceDirectories(ExtractFileDir(Path));
    SL.SaveToFile(Path, TEncoding.UTF8);

    TFile.AppendAllText(LogPath,
      Format('[%s] %s -> %s (header + %d rows)',
        [FormatDateTime('yyyy-mm-dd hh:nn:ss', Now), Title, Path, WroteDataCount])
      + sLineBreak, TEncoding.UTF8);

    MemoStep(AMemo, Format('%s completed (%d rows written)', [Title, WroteDataCount]));
  finally
    SL.Free;
  end;
end;

{-------------------- PROCESS --------------------}

procedure SaveProcessAsLongCSV(const SrcSheet: OleVariant;
  const Path: string; const LogPath: string; AProgress: TProgressBar; AMemo: TMemo);
var
  SL: TStringList;
  LastRow, StartRow, EndRow: Integer;
  Total, Done, R: Integer;
  MfgNo, PartFig, ProcName, SetVal, MaVal: string;
  FirstOnlyProcCol, TripletStartCol, TripletEndCol, MaxTriples, k: Integer;
  ProcCol, SetCol, MaCol: Integer;
begin
  MemoStep(AMemo, 'Start PROCESS → ' + Path);
  SL := TStringList.Create;
  try
    SL.Add(CsvJoin(TArray<string>.Create('Mfg.No.','Part figure','process','set','ma')));
    LastRow := SrcSheet.UsedRange.Row + SrcSheet.UsedRange.Rows.Count - 1;
    StartRow := 3; EndRow := LastRow;
    FirstOnlyProcCol := COL_PROCESS_FIRST;
    TripletStartCol  := COL_PROCESS_FIRST + 1;
    TripletEndCol    := COL_PROCESS_LAST;
    if TripletEndCol >= TripletStartCol then
      MaxTriples := (TripletEndCol - TripletStartCol + 1) div 3 else MaxTriples := 0;

    if EndRow >= StartRow then
      Total := 1 + (EndRow - StartRow + 1) * (1 + MaxTriples)
    else Total := 1;

    Done := 0;
    if Assigned(AProgress) then
    begin
      AProgress.Position := 0;
      AProgress.Max := Total;
    end;

    Inc(Done);
    if Assigned(AProgress) then AProgress.Position := Done;

    for R := StartRow to EndRow do
    begin
      MfgNo   := VarToStr(SrcSheet.Cells[R, COL_M].Value);
      PartFig := VarToStr(SrcSheet.Cells[R, COL_F].Value);
      ProcName := VarToStr(SrcSheet.Cells[R, FirstOnlyProcCol].Value);
      if Trim(ProcName) <> '' then
        SL.Add(CsvJoin(TArray<string>.Create(MfgNo, PartFig, ProcName, '0', '0')));

      for k := 0 to MaxTriples-1 do
      begin
        ProcCol := TripletStartCol + k*3;
        SetCol  := ProcCol + 1;
        MaCol   := ProcCol + 2;
        ProcName := VarToStr(SrcSheet.Cells[R, ProcCol].Value);
        SetVal := VarToStr(SrcSheet.Cells[R, SetCol].Value);
        MaVal  := VarToStr(SrcSheet.Cells[R, MaCol].Value);
        if Trim(SetVal) = '' then SetVal := '0';
        if Trim(MaVal)  = '' then MaVal  := '0';
        if Trim(ProcName) <> '' then
          SL.Add(CsvJoin(TArray<string>.Create(MfgNo, PartFig, ProcName, SetVal, MaVal)));
      end;
    end;

    SL.SaveToFile(Path, TEncoding.UTF8);
    MemoStep(AMemo, Format('PROCESS completed (%d rows)', [SL.Count-1]));
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
  HeaderJob, HeaderPart: TArray<string>;
  SL: TStringList;
  i: Integer;
  ExcelDate: Variant;
  SDate: string;
begin
  HeaderJob := TArray<string>.Create('CstmrCD','Cstmr.Name','Mfg.No.','RE','ProductName');
  HeaderPart := TArray<string>.Create('Mfg.No.','PartsName','Material','SizeRemarks','PartsQuantity');

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
      LogDir := IncludeTrailingPathDelimiter(ExtractFilePath(ParamStr(0)))+'LOG\';
      if not TDirectory.Exists(LogDir) then TDirectory.CreateDirectory(LogDir);
      LogPath := TPath.Combine(LogDir,'export_log_'+FormatDateTime('yyyymmdd',Now)+'.txt');
    end;
  finally
    Ini.Free;
  end;

  if not FileExists(InputFile) then
    raise Exception.CreateFmt('Input file not found: %s',[InputFile]);

  CoInitialize(nil);
  try
    Excel := CreateOleObject('Excel.Application');
    Excel.Visible := False;
    WB := Excel.Workbooks.Open(InputFile);
    try
      Sheet := WB.Worksheets[SheetName];
      MemoStep(AMemo,'Using sheet: '+SheetName);

      // ===== JOB =====
      if OutJob <> '' then
      begin
        SaveCSVWithCustomHeaderByCols(Sheet,[1,2,COL_M,4,5],OutJob,'JOB',LogPath,AProgress,AMemo,HeaderJob);

        // ✅ Read StartDate from column O and +3 days
        SL := TStringList.Create;
        try
          SL.LoadFromFile(OutJob,TEncoding.UTF8);
          if SL.Count>0 then
          begin
            SL[0] := SL[0]+',StartDate';
            for i:=1 to SL.Count-1 do
            begin
              ExcelDate := Sheet.Cells[i+2,COL_O].Value; // Column O
              if not VarIsNull(ExcelDate) and not VarIsEmpty(ExcelDate) then
              begin
                try
                  SDate := FormatDateTime('dd/mm/yyyy',IncDay(VarToDateTime(ExcelDate),3));
                except
                  SDate := '';
                end;
              end
              else
                SDate := '';
              SL[i] := SL[i]+','+SDate;
            end;
          end;
          SL.SaveToFile(OutJob,TEncoding.UTF8);
        finally
          SL.Free;
        end;
        MemoStep(AMemo,'Added StartDate (column O +3 days) to JOB file');
      end;

      // ===== PART =====
      if OutPart <> '' then
        SaveCSVWithCustomHeaderByCols(Sheet,[COL_M,6,7,8,9],OutPart,'PART',LogPath,AProgress,AMemo,HeaderPart);

      // ===== PROCESS =====
      if OutProcess <> '' then
        SaveProcessAsLongCSV(Sheet,OutProcess,LogPath,AProgress,AMemo);

      MemoStep(AMemo,'Export completed successfully.');
    finally
      WB.Close(False);
      Excel.Quit;
    end;
  finally
    CoUninitialize;
  end;
end;

end.

