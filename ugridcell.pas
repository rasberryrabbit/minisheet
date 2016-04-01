unit uGridCell;

{$mode objfpc}{$H+}

{$i usheettabconf.inc}

interface

uses
  Classes, SysUtils,
  Grids { {$ifdef WINDOWS},messages{$endif} },Graphics, contnrs;

type

  TEditHistoryItem = class
    public
      CellPos:TPoint;
      Buffer:string;
  end;

  { TEditHistory }

  TEditHistory = class(TObjectStack)
    public
      destructor Destroy; override;

      procedure Clear;

      procedure PushHistory(const Pos:TPoint;const s:string);
      function PopHistory:TEditHistoryItem;
      function PeekHistory:TEditHistoryItem;
      function FindHistory(ACol, ARow:Integer):TEditHistoryItem;
  end;

  { TStrGridCell }

  TStrGridCell=class(TStringGrid)
    private
      procedure CompactCells(var ValidCol: Integer; var ValidRow: Integer);
      function GetRowValidMax:Integer;
      function GetCellFromStr(var s: string; delimiter: char=','): string;
      procedure GridColRowInserted(Sender: TObject; IsColumn: Boolean; sIndex,
        tIndex: Integer);
      procedure GridDrawCell(Sender: TObject; aCol, aRow: Integer;
        aRect: TRect; aState: TGridDrawState);
      {$ifdef USE_SPREADCELL}
      procedure GridGetEditText(Sender: TObject; ACol, ARow: Integer;
        var Value: string);
      {$endif}
      procedure GridValidateEntry(sender: TObject; aCol, aRow: Integer;
        const OldValue: string; var NewValue: String);
      {
      {$ifdef WINDOWS}
      procedure GridIMEStartComposition(var msg:TMessage); message WM_IME_STARTCOMPOSITION;
      procedure GridIMECompisiton(var msg:TMessage); message WM_IME_COMPOSITION;
      {$endif}
      }
      procedure SetSelectColorMask(Value:Integer);
    protected
      // selected text
      SelTextColor:TColor;
      NormalTextColor:TColor;
      FHistoryStack:TEditHistory;
      FRedoStack:TEditHistory;
      FDoHistory:Boolean;
      FSelectColorMask:Integer;

      function CopyToBuffer(const R: TGridRect): String;
      procedure DoCopyToClipboard; override;
      procedure KeyDown(var Key: Word; Shift: TShiftState); override;
      procedure DoCutToClipboard; override;
      procedure DoPasteFromClipboard; override;
      procedure AutoAdjustColumn(aCol: Integer); override;
      function MakeSmartCell(const str: string; aCol, aRow: Integer): string;
      function GetEditText(aCol, aRow: Integer): string; override;
      procedure SetCells(ACol, ARow: Integer; const AValue: string); override;
      function SelectionSetTextEx(txt:string;StartCol,StartRow,ColLimit,RowLimit:Integer):TPoint;
      procedure DoEnter; override;
    public
      CopySolved:Boolean;
      AutoSizeCell:Boolean;
      MaxCellWidth:Integer;
      // cell copy
      LastCopyCell:TPoint;

      constructor Create(AOwner: TComponent); override;
      destructor Destroy; override;
      procedure PasteFromClipboardCell;
      procedure CutToClipboardCell;
      procedure UnDelete;
      procedure RedoDelete;
      procedure UnDeleteCell;
      procedure AddEditHistory;
      procedure LoadFromCSVStream(Stream:TStream;delimiter:char=',');
      procedure LoadFromCSVFile(const Filename:string;delimiter:char=',');
      procedure SaveToCSVStream(Stream:TStream;delimiter:char=',');
      procedure SaveToCSVFile(const Filename:string;delimiter:char=',');
      procedure ExportToCSVStream(Stream:TStream;delimiter:char=',');
      procedure ExportToCSVFile(const Filename:string;delimiter:char=',');
      procedure ExportToExcelFile(Filename: string);
      procedure ImportFromExcelFile(const Filename:string);

      property EditHistory:TEditHistory read FHistoryStack;
      property RedoHistory:TEditHistory read FRedoStack;
      property SelectColorMask:Integer read FSelectColorMask write SetSelectColorMask;
      property RowValidMax:Integer read GetRowValidMax;
  end;

  procedure UpdateColLabel(Grid: TStringGrid; s, e: Integer);

implementation

uses Clipbrd, types, usheetTab, uCellFormulaPaser, LCLType,
  fpspreadsheet{$ifdef WINDOWS}, LCLIntf{$endif}, FPCanvas,
  fpsTypes; // fpshpreadsheet 1.6

const
  CMaxHistory=500;

function GetWorkSheet(Grid: TStrGridCell): TCellTabSheet;
begin
  if Grid<>nil then
    Result:=TCellTabSheet(Grid.Parent);
end;

procedure UpdateColLabel(Grid: TStringGrid; s, e: Integer);
var
  i:Integer;
begin
  if s=-1 then
    s:=0;
  if e=-1 then
    e:=Grid.ColCount-1;
  for i:=s to e do
    Grid.Cells[i,0]:=MakeColumnStr(i-1);
end;

function Is_OneCellBuffer(const s:string):Boolean;
var
  strp:PChar;
begin
  result:=True;
  strp:=PChar(s);
  while strp^<>#0 do begin
    if strp^=#9 then begin
      Result:=False;
      break;
    end;
    Inc(strp);
  end;
end;

{ TEditHistory }

destructor TEditHistory.Destroy;
begin
  while Count>0 do
    TEditHistoryItem(Pop).Free;
  inherited Destroy;
end;

procedure TEditHistory.Clear;
begin
  while Count>0 do
    TEditHistoryItem(pop).Free;
end;

procedure TEditHistory.PushHistory(const Pos: TPoint; const s: string);
var
  newobj:TEditHistoryItem;
begin
  if Count>CMaxHistory then begin
    newobj:=TEditHistoryItem(List.Items[0]);
    List.Delete(0);
    newobj.Free;
  end;

  newobj:=TEditHistoryItem.Create;
  try
    newobj.CellPos:=Pos;
    newobj.Buffer:=s;
    Push(newobj);
  except
    newobj.Free;
  end;
end;

function TEditHistory.PopHistory: TEditHistoryItem;
begin
  if Count>0 then
    Result:=TEditHistoryItem(Pop)
    else
      Result:=nil;
end;

function TEditHistory.PeekHistory: TEditHistoryItem;
begin
  if Count>0 then
    Result:=TEditHistoryItem(Peek)
    else
      Result:=nil;
end;

function TEditHistory.FindHistory(ACol, ARow: Integer): TEditHistoryItem;
var
  i:Integer;
  temp:TEditHistoryItem;
begin
  // find ACol, ARow previous text, and pop item
  Result:=nil;
  if Count>0 then begin
    i:=Count-1;
    while i>=0 do begin
      temp:=TEditHistoryItem(List.Items[i]);
      if (temp.CellPos.x=ACol) and (temp.CellPos.y=ARow) then begin
        Result:=Temp;
        List.Delete(i);
        break;
      end;
      Dec(i);
    end;
  end;
end;


{ TStrGridCell }

procedure TStrGridCell.CompactCells(var ValidCol: Integer;
  var ValidRow: Integer);
var
  R:TStringList;
begin
  // compact Rows
  R:=TStringList.Create;
  try
    ValidRow:=RowCount-1;
    while ValidRow>1 do begin
      R.Assign(Rows[ValidRow]);
      R.Delimiter:=#32;
      R.Delete(0);
      if Trim(R.DelimitedText)<>'' then begin
        break;
      end;
      Dec(ValidRow);
    end;

    ValidCol:=ColCount-1;
    while ValidCol>1 do begin
      R.Assign(Cols[ValidCol]);
      R.Delimiter:=#32;
      R.Delete(0);
      if Trim(R.DelimitedText)<>'' then begin
        break;
      end;
      Dec(ValidCol);
    end;

  finally
    R.Free;
  end;
end;

function TStrGridCell.GetRowValidMax: Integer;
var
  x:Integer;
begin
  CompactCells(x,Result);
  if Result<=0 then
    Result:=1;
end;

function TStrGridCell.GetCellFromStr(var s: string; delimiter:char=','): string;
var
  i, j, l:Integer;
  ch:char;
begin
  l:=Length(s);
  i:=1;
  j:=0;
  Result:='';
  while i<=l do begin
    ch:=s[i];
    if (j=0) and (ch=delimiter) then begin
      break;
    end else
    if s[i]='"' then begin
      if j=0 then
        Inc(j)
        else begin
          Dec(j);
          if i<l then begin
            Inc(i);
            ch:=s[i];
            if ch='"' then begin
              Inc(j);
              Result:=Result+'"';
            end else
              continue;
          end;
        end;
    end else
      Result:=Result+ch;
    Inc(i);
  end;
  if i>l then
    i:=l;
  Delete(s,1,i);
end;

procedure TStrGridCell.DoCopyToClipboard;
var
  SelStr: String;
  R:TGridRect;
begin
  R:=Selection;
  LastCopyCell:=R.TopLeft;

  SelStr:=CopyToBuffer(R);
  Clipboard.AsText := SelStr;
end;

procedure TStrGridCell.KeyDown(var Key: Word; Shift: TShiftState);
var
  selupd:Boolean;
  rec:TRect;
  pp:TPoint;
  lastedit:TEditHistoryItem;
begin
  // editor exit keys
  selupd:=Key in [VK_LEFT,VK_RIGHT,VK_UP,VK_DOWN,VK_PRIOR,VK_NEXT];
  // increase column
  if Key=VK_RIGHT then begin
    if ColCount<255 then
      if Col=(ColCount-1) then begin
        ColCount:=ColCount+1;
        if ColCount-VisibleColCount>LeftCol then
          LeftCol:=LeftCol+1;
      end;
  end;
  // popup menu
  if (Key=VK_APPS) and (not EditorMode) then begin
    if Assigned(PopupMenu) then begin
      rec:=CellRect(Col,Row);
      pp:=ControlToScreen(Point(rec.Right,rec.Top));
      PopupMenu.PopUp(pp.x,pp.y);
    end;
  end else
  if (Key=VK_DELETE) and (not EditorMode) then begin
    if Shift=[ssShift] then begin
      UnDelete;
    end else
    if Shift=[] then begin
      // delete
      FHistoryStack.PushHistory(Point(Selection.Left,Selection.Top),CopyToBuffer(Selection));
      Clean(Selection,[]);
    end;
    selupd:=(Shift=[]) or (Shift=[ssShift]);
  end else
  if (not EditorKey) and
     ((Key in [VK_V,VK_X]) and (Shift=[ssCtrl,ssShift])) then
    // disable default functions
  else
  if (not EditorKey) and
     ((Key=VK_X) and (Shift=[ssCtrl])) then
       // cut
       DoCutToClipboard
  else
  if EditorKey and (Key=VK_ESCAPE) and (Shift = []) then begin
    // Handle ESC key in Editor
    if (FHistoryStack.PeekHistory<>nil) and
       (FHistoryStack.PeekHistory.CellPos.x=Col) and
       (FHistoryStack.PeekHistory.CellPos.y=Row) then
    begin
      lastedit:=FHistoryStack.PopHistory;
      try
      Editor.SetTextBuf(pchar(lastedit.Buffer));
      finally
        lastedit.Free;
      end;
    end;
    EditorMode:=False;
  end
  else
  inherited KeyDown(Key, Shift);
  if (not EditorKey) and selupd then
    SelectCell(Col,Row);
end;

procedure TStrGridCell.DoCutToClipboard;
begin
  if EditingAllowed(Col) then begin
    DoCopyToClipboard;
    FHistoryStack.PushHistory(Point(Selection.Left,Selection.Top),CopyToBuffer(Selection));
    Clean(Selection,[]);
    if not EditorKey then
       SelectCell(Col,Row);
  end;
end;

procedure TStrGridCell.DoPasteFromClipboard;
begin
  FDoHistory:=True;
  try
    inherited DoPasteFromClipboard;
  finally
    FDoHistory:=False;
  end;
  if not EditorKey then
     SelectCell(Col,Row);
end;

procedure TStrGridCell.PasteFromClipboardCell;
var
  txt:string;
  rpos:TPoint;
  iCol,iRow,iRight,iBottom:Integer;
  bOneCell:Boolean;
begin
  if EditingAllowed(Col) and Clipboard.HasFormat(CF_TEXT) then begin
    iCol:=Selection.Left;
    rpos.x:=0;
    FDoHistory:=True;
    bOneCell:=(Selection.Right-Selection.Left=0) and (Selection.Bottom-Selection.Top=0);
    try
      if bOneCell then begin
        iRight:=ColCount-1;
        iBottom:=RowCount-1;
      end else begin
        iRight:=Selection.Right;
        iBottom:=Selection.Bottom;
      end;
      while (iCol<=iRight) or bOneCell do begin
        iRow:=Selection.Top;
        while (iRow<=Selection.Bottom) or bOneCell do begin
          txt:=MakeSmartCell(Clipboard.AsText,iCol,iRow);
          rpos:=SelectionSetTextEx(txt,iCol,iRow,iRight,iBottom);
          if rpos.y=0 then
            Inc(iRow)
            else
              Inc(iRow,rpos.y);
          if bOneCell then
            break;
        end;
        if rpos.x=0 then
          Inc(iCol)
          else
            Inc(iCol,rpos.x);
        if bOneCell then
          break;
      end;
    finally
      FDoHistory:=False;
    end;
    if not EditorKey then
       SelectCell(Col,Row);
  end;
end;

procedure TStrGridCell.CutToClipboardCell;
begin
  if EditingAllowed(Col) then begin
    CopySolved:=False;
    DoCopyToClipboard;
    FHistoryStack.PushHistory(Point(Selection.Left,Selection.Top),CopyToBuffer(Selection));
    CopySolved:=True;
    Clean(Selection,[]);
    if not EditorKey then
      SelectCell(Col,Row);
  end;
end;

procedure TStrGridCell.AutoAdjustColumn(aCol: Integer);
var
  i,W: Integer;
  Ts: TSize;
  TmpCanvas: TCanvas;
  C: TGridColumn;
  txt:string;
begin
  if (aCol<1) or (aCol>ColCount-1) then
    Exit;

  tmpCanvas := GetWorkingCanvas(Canvas);

  C := ColumnFromGridColumn(aCol);

  try
    W:=0;
    for i := 1 to RowCount-1 do begin

      if C<>nil then begin
        if i<FixedRows then
          tmpCanvas.Font := C.Title.Font
        else
          tmpCanvas.Font := C.Font;
      end else begin
        if i<FixedRows then
          tmpCanvas.Font := TitleFont
        else
          tmpCanvas.Font := Font;
      end;

      if (i=0) and (FixedRows>0) and (C<>nil) then
        Ts := TmpCanvas.TextExtent(C.Title.Caption)
      else begin
        txt:=Cells[aCol,i];
        if (txt<>'') and (txt[1] in ['''','>','<']) then
          Delete(txt,1,1)
          else
            if not GetWorkSheet(self).Solver.CheckNumber(txt) then
              GetWorkSheet(self).SolveFormula(txt);
        Ts := TmpCanvas.TextExtent(txt);
      end;

      if Ts.Cx>W then
        W := Ts.Cx;
    end;
  finally
    if tmpCanvas<>Canvas then
      FreeWorkingCanvas(tmpCanvas);
  end;

  if W=0 then
    W := DefaultColWidth
  else
    W := W + 8;

  if MaxCellWidth<>0 then
    if W>MaxCellWidth then
      W:=MaxCellWidth;

  if W<DefaultColWidth then
    W:=DefaultColWidth;

  ColWidths[aCol] := W;
end;

function TStrGridCell.MakeSmartCell(const str: string;aCol,aRow:Integer): string;
begin
  Result:=ConvertCell(LastCopyCell,Point(aCol,aRow),str);
end;

function TStrGridCell.GetEditText(aCol, aRow: Integer): string;
begin
  Result:=inherited GetEditText(aCol, aRow);
  FHistoryStack.PushHistory(Point(ACol,ARow),Result);
end;

procedure TStrGridCell.SetCells(ACol, ARow: Integer; const AValue: string);
begin
  if FDoHistory then
    FHistoryStack.PushHistory(Point(ACol,ARow),Cells[ACol,ARow]);
  inherited SetCells(ACol, ARow, AValue);
end;

function TStrGridCell.SelectionSetTextEx(txt: string; StartCol, StartRow,
  ColLimit, RowLimit: Integer): TPoint;
var
  L,SubL: TStringList;
  i,j: Integer;
  procedure CollectCols(const S: String);
  var
    P,Ini: PChar;
    St: String;
  begin
    Subl.Clear;
    P := Pchar(S);
    if P<>nil then
      while P^<>#0 do begin
        ini := P;
        while (P^<>#0) and (P^<>#9) do
          Inc(P);
        SetLength(St, P-Ini);
        Move(Ini^,St[1],P-Ini);
        SubL.Add(St);
        if P^<>#0 then
          Inc(P);
      end;
  end;
begin
  L := TStringList.Create;
  SubL := TStringList.Create;
  Result:=Point(0,0);
  try
    L.Text := txt;
    Result.y:=L.Count;
    for j:=0 to L.Count-1 do begin
      if (j+StartRow >= RowCount) or (j+StartRow>RowLimit) then
        break;
      CollectCols(L[j]);
      if SubL.Count>Result.x then
        Result.x:=SubL.Count;
      for i:=0 to SubL.Count-1 do
        if (i+StartCol<ColCount) and (i+StartCol<=ColLimit) and (not GetColumnReadonly(i+StartCol)) then
           Cells[i + StartCol, j + StartRow] := SubL[i];
    end;
  finally
    SubL.Free;
    L.Free;
  end;
end;

procedure TStrGridCell.DoEnter;
begin
  inherited DoEnter;
  SelectCell(Col,Row);
end;

constructor TStrGridCell.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  CopySolved:=True;
  LastCopyCell:=Point(1,1);
  AutoSizeCell:=False;
  MaxCellWidth:=0;
  AutoAdvance:=aaDown;
  OnDrawCell:=@GridDrawCell;
  OnValidateEntry:=@GridValidateEntry;
  OnColRowInserted:=@GridColRowInserted;
  FDoHistory:=False;
  SetSelectColorMask($00FFAA00);
  FHistoryStack:=TEditHistory.Create;
  FRedoStack:=TEditHistory.Create;
  {$ifdef USE_SPREADCELL}
  OnGetEditText:=@GridGetEditText;
  {$endif}
end;

destructor TStrGridCell.Destroy;
begin
  FHistoryStack.Free;
  FRedoStack.Free;
  inherited Destroy;
end;

procedure TStrGridCell.LoadFromCSVStream(Stream: TStream; delimiter: char);
var
  R:TStringList;
  s:string;
  i,j:Integer;
begin
  RowCount:=30;
  ColCount:=8;
  Clean(1,1,7,29,[]);
  R:=TStringList.Create;
  try
    R.LoadFromStream(Stream);
    for i:=0 to R.Count-1 do begin
      s:=R.Strings[i];
      if (s<>'') and (i>=RowCount) then
        RowCount:=RowCount+1;
      j:=1;
      while s<>'' do begin
        if j>=ColCount then begin
          ColCount:=j+1;
        end;
        Cells[j,i+1]:=GetCellFromStr(s,delimiter);
        Inc(j);
      end;
    end;
    FHistoryStack.Clear;
    FRedoStack.Clear;
  finally
    R.Free;
  end;
  SelectCell(FixedCols,FixedRows);
end;

procedure TStrGridCell.LoadFromCSVFile(const Filename: string; delimiter: char);
var
  fs:TFileStream;
begin
  fs:=TFileStream.Create(Filename,fmOpenRead or fmShareDenyWrite);
  try
    LoadFromCSVStream(fs,delimiter);
    FHistoryStack.Clear;
    FRedoStack.Clear;
  finally
    fs.Free;
  end;
end;

procedure TStrGridCell.SaveToCSVStream(Stream: TStream; delimiter: char);
var
  NewLine,R:TStringList;
  j,ValidRow,ValidCol,i:Integer;
begin
  CompactCells(ValidCol, ValidRow);

  NewLine:=TStringList.Create;
  try
    R:=TStringList.Create;
    try
      for j:=1 to ValidRow do begin
        R.Assign(Rows[j]);
        if R.Count>0 then begin
          i:=R.Count-1;
          while i>ValidCol do begin
            R.Delete(i);
            Dec(i);
          end;
        end;
        if R.Count>0 then
          R.Delete(0);
        R.StrictDelimiter:=False;
        R.Delimiter:=delimiter;
        NewLine.Add(R.DelimitedText);
      end;
    finally
      R.Free;
    end;
    NewLine.SaveToStream(Stream);
  finally
    NewLine.Free;
  end;
end;

procedure TStrGridCell.SaveToCSVFile(const Filename: string; delimiter: char);
var
  fs:TFileStream;
begin
  fs:=TFileStream.Create(Filename,fmOpenWrite or fmCreate or fmShareDenyWrite);
  try
    SaveToCSVStream(fs,delimiter);
  finally
    fs.Free;
  end;
end;

procedure TStrGridCell.ExportToCSVStream(Stream: TStream; delimiter: char);
var
  NewLine,R:TStringList;
  j,ValidRow,ValidCol,i:Integer;
  txt:string;
begin
  CompactCells(ValidCol, ValidRow);

  NewLine:=TStringList.Create;
  try
    R:=TStringList.Create;
    try
      for j:=1 to ValidRow do begin
        R.Assign(Rows[j]);
        if R.Count>0 then begin
          i:=R.Count-1;
          while i>ValidCol do begin
            R.Delete(i);
            Dec(i);
          end;
        end;
        if R.Count>0 then
          R.Delete(0);
        for i:=0 to R.Count-1 do begin
          txt:=R.Strings[i];
          if txt<>'' then
            if not(txt[1] in ['''','>','<']) then begin
              if not GetWorkSheet(self).Solver.CheckNumber(txt) then
                GetWorkSheet(self).SolveFormula(txt);
            end else
              delete(txt,1,1);
          R.Strings[i]:=txt;
        end;
        R.StrictDelimiter:=False;
        R.Delimiter:=delimiter;
        NewLine.Add(R.DelimitedText);
      end;
    finally
      R.Free;
    end;
    NewLine.SaveToStream(Stream);
  finally
    NewLine.Free;
  end;
end;

procedure TStrGridCell.ExportToCSVFile(const Filename: string; delimiter: char);
var
  fs:TFileStream;
begin
  fs:=TFileStream.Create(Filename,fmOpenWrite or fmCreate or fmShareDenyWrite);
  try
    ExportToCSVStream(fs,delimiter);
  finally
    fs.Free;
  end;
end;

procedure TStrGridCell.ExportToExcelFile(Filename: string);
var
  WorkB:TsWorkbook;
  Sheet:TsWorksheet;
  j,ValidRow,ValidCol,i:Integer;
  txt:string;
  fm:TsSpreadsheetFormat;
begin
  CompactCells(ValidCol, ValidRow);

  WorkB:=TsWorkbook.Create;
  try
    Sheet:=WorkB.AddWorksheet('Sheet1');

    for j:=1 to ValidRow do
      for i:=1 to ValidCol do begin
        txt:=Cells[i,j];
        if txt<>'' then begin
          if not (txt[1] in ['''','>','<']) then begin
            if not GetWorkSheet(self).Solver.CheckNumber(txt) then
              GetWorkSheet(self).SolveFormula(txt);
          end else
            Delete(txt,1,1);
          Sheet.WriteUTF8Text(j-1,i-1,txt);
        end;
      end;

    // default = .ods
    if not WorkB.GetFormatFromFileHeader{GetFormatFromFileName}(Filename,fm) then begin
      fm:=sfOpenDocument;
      ChangeFileExt(Filename,STR_OPENDOCUMENT_CALC_EXTENSION);
    end;
    WorkB.WriteToFile(Filename,fm,True);
  finally
    WorkB.Free;
  end;
end;

procedure TStrGridCell.ImportFromExcelFile(const Filename: string);
var
  workb:TsWorkbook;
  sheet:TsWorksheet;
  cell:PCell;
  lrow,lcol,i:Integer;
  txt:string;
begin
  workb:=TsWorkbook.Create;
  try
    workb.ReadFromFile(Filename);
    if workb.GetWorksheetCount>0 then begin
      sheet:=workb.GetWorksheetByIndex(0);
      if (sheet<>nil) and (sheet.GetCellCount>0) then begin
        lcol:=sheet.GetLastColNumber+2;
        lrow:=sheet.GetLastRowNumber+2;
        if lcol>ColCount then
          ColCount:=lcol;
        if lrow>RowCount then
          RowCount:=lrow;
        Clean(1,1,ColCount-1,RowCount-1,[]);

        {
        cell:=sheet.GetFirstCell();

        for i:=0 to sheet.GetCellCount-1 do begin
          lrow:=cell^.Row;
          lcol:=cell^.Col;
          Cells[lcol+1,lrow+1]:=sheet.ReadAsUTF8Text(lrow,lcol);

          cell:=sheet.GetNextCell();
        end;
        }
        // 1.6
        for cell in sheet.Cells do begin
          lrow:=cell^.Row;
          lcol:=cell^.Col;
          Cells[lcol+1,lrow+1]:=sheet.ReadAsUTF8Text(lrow,lcol);
        end;
      end;
    end;
    FHistoryStack.Clear;
    FRedoStack.Clear;
  finally
    workb.Free;
  end;
  SelectCell(FixedCols,FixedRows);
end;

procedure TStrGridCell.GridDrawCell(Sender: TObject; aCol, aRow: Integer;
  aRect: TRect; aState: TGridDrawState);
const
  _RightCellMargin=5;
var
  {$ifdef USE_SPREADCELL}
  cell:PCell;
  {$endif}
  txt,inf:string;
  ts:TTextStyle;
begin
  if (aCol>0) and (aRow>0) then begin
    FillChar(ts,sizeof(TTextStyle),0);
    ts.Alignment:=taLeftJustify;
    ts.Layout:=tlCenter;
    ts.SingleLine:=True;
    ts.EndEllipsis:=True;
    {$ifdef USE_SPREADCELL}
    cell:=GetWorkSheet(Grid).Sheet.GetCell(aRow-1,aCol-1);
    txt:=cell^.UTF8StringValue;
    {$else}
    txt:=Cells[aCol,aRow];
    {$endif}
    if txt<>'' then
      // text justify
      if not (txt[1] in ['''','>','<']) then begin
        ts.Alignment:=taRightJustify;
        if not GetWorkSheet(self).Solver.CheckNumber(txt) then
          if 0<>GetWorkSheet(self).SolveFormula(txt) then
            ts.Alignment:=taLeftJustify;
      end
      else begin
        if txt[1]='>' then
          ts.Alignment:=taRightJustify
          else
            if txt[1]='<' then
              ts.Alignment:=taCenter;
        Delete(txt,1,1);
      end;
    if gdSelected in aState then begin
      Canvas.Brush.Color:=SelectedColor;
      Canvas.Font.Color:=SelTextColor;
    end else begin
      Canvas.Brush.Color:=Color;
      Canvas.Font.Color:=NormalTextColor;
    end;
    Canvas.FillRect(aRect);
    if ts.Alignment=taRightJustify then
      Dec(aRect.Right,_RightCellMargin);
    Inc(aRect.Left);
    Canvas.TextRect(aRect,aRect.Left,aRect.Top,txt,ts);
  end else
    DefaultDrawCell(aCol,aRow,aRect,aState);
end;

procedure TStrGridCell.GridValidateEntry(sender: TObject; aCol, aRow: Integer;
  const OldValue: string; var NewValue: String);
{$ifdef USE_SPREADCELL}
var
  cell:PCell;
  works:TsWorksheet;
  wsolve:TCellFormula;
{$endif}
begin
{$ifdef USE_SPREADCELL}
  if (aCol>0) and (aRow>0) then begin
    works:=GetWorkSheet(Sender as TStringGrid).Sheet;
    wsolve:=GetWorkSheet(sender as TStringGrid).Solver;
    if WorkS<>nil then begin
      cell:=WorkS.GetCell(aRow-1,aCol-1);
      cell^.ContentType:=cctUTF8String;
      cell^.UTF8StringValue:=NewValue;
    end;
  end;
{$endif}
  if AutoSizeCell then
    AutoSizeColumn(aCol);
  Invalidate;
  if ((AutoAdvance=aaDown) and (aRow=RowCount-1)) or
     ((AutoAdvance=aaRight) and (aCol=ColCount-1)) then
       MoveExtend(False,aCol,aRow);

end;

(*
{$ifdef WINDOWS}
procedure TStrGridCell.GridIMEStartComposition(var msg: TMessage);
begin
  EditorMode:=True;
  if InplaceEditor<>nil then
    msg.Result:=SendMessage(InplaceEditor.Handle,msg.msg,msg.wParam,msg.lParam);
end;

procedure TStrGridCell.GridIMECompisiton(var msg: TMessage);
begin
  if InplaceEditor<>nil then
    msg.Result:=SendMessage(InplaceEditor.Handle,msg.msg,msg.wParam,msg.lParam);
end;
{$endif}
*)

procedure TStrGridCell.SetSelectColorMask(Value: Integer);
var
  cl:TColor;
begin
  if (Value and $00FFFFFF)<>FSelectColorMask then begin
    FSelectColorMask:=(Value and $00FFFFFF);
    cl:=ColorToRGB(Color);
    SelectedColor:=cl and FSelectColorMask;
    SelTextColor:=(not cl) and (not FSelectColorMask) and $00FFFFFF;
    NormalTextColor:=(not cl) and $00FFFFFF;
  end;
end;

function TStrGridCell.CopyToBuffer(const R: TGridRect): String;
var
  k: LongInt;
  j: LongInt;
  i: LongInt;
  txt: String;
begin
  Result := '';
  for i:=R.Top to R.Bottom do begin
    for j:=R.Left to R.Right do begin
      if Columns.Enabled and (j>=FirstGridColumn) then begin
        k := ColumnIndexFromGridColumn(j);
        if not Columns[k].Visible then
          continue;
        if (i=0) then
          Result := Result + Columns[k].Title.Caption
        else
          Result := Result + Cells[j, i];
      end else begin
        txt:=Cells[j, i];
        if CopySolved and (txt<>'') then
          if not (txt[1] in ['''', '>', '<']) then begin
            if not GetWorkSheet(self).Solver.CheckNumber(txt) then
              GetWorkSheet(self).SolveFormula(txt);
          end;
        Result := Result + txt;
      end;
      if j<>R.Right then
        Result := Result + #9;
    end;
    Result := Result + #13#10;
  end;
end;

procedure TStrGridCell.UnDelete;
var
  ehist:TEditHistoryItem;
begin
  if FHistoryStack.Count>0 then begin
    ehist:=FHistoryStack.PopHistory;
    try
      Col:=ehist.CellPos.x;
      Row:=ehist.CellPos.y;
      if not Is_OneCellBuffer(ehist.Buffer) then
         SelectionSetText(ehist.Buffer)
         else begin
           FRedoStack.PushHistory(ehist.CellPos,Cells[ehist.CellPos.x,ehist.CellPos.y]);
           Cells[Col,Row]:=ehist.Buffer;
         end;
      SelectCell(Col,Row);
    finally
      ehist.Free;
    end;
  end;
end;

procedure TStrGridCell.RedoDelete;
var
  ehist:TEditHistoryItem;
begin
  if FRedoStack.Count>0 then begin
    ehist:=FRedoStack.PopHistory;
    try
      Col:=ehist.CellPos.x;
      Row:=ehist.CellPos.y;
      if not Is_OneCellBuffer(ehist.Buffer) then
         SelectionSetText(ehist.Buffer)
         else begin
           FHistoryStack.PushHistory(ehist.CellPos,Cells[ehist.CellPos.x,ehist.CellPos.y]);
           Cells[Col,Row]:=ehist.Buffer;
         end;
      SelectCell(Col,Row);
    finally
      ehist.Free;
    end;
  end;
end;

procedure TStrGridCell.UnDeleteCell;
var
  ehist:TEditHistoryItem;
begin
  ehist:=FHistoryStack.FindHistory(Col,Row);
  if ehist<>nil then begin
    try
      if not Is_OneCellBuffer(ehist.Buffer) then
         SelectionSetText(ehist.Buffer)
         else begin
           FRedoStack.PushHistory(ehist.CellPos,Cells[ehist.CellPos.x,ehist.CellPos.y]);
           Cells[Col,Row]:=ehist.Buffer;
         end;
      SelectCell(Col,Row);
    finally
      ehist.Free;
    end;
  end;
end;

procedure TStrGridCell.AddEditHistory;
begin
  FHistoryStack.PushHistory(Point(Col,Row),Cells[Col,Row]);
end;

procedure TStrGridCell.GridColRowInserted(Sender: TObject; IsColumn: Boolean;
  sIndex, tIndex: Integer);
begin
  if IsColumn then
    UpdateColLabel(self,sIndex,tIndex);
end;

{$ifdef USE_SPREADCELL}
procedure TStrGridCell.GridGetEditText(Sender: TObject; ACol, ARow: Integer;
  var Value: string);
var
  cell:PCell;
  works:TsWorksheet;
begin
  works:=GetWorkSheet(self).Sheet;
  if works<>nil then begin
    cell:=works.GetCell(ARow-1,ACol-1);
    Value:=cell^.UTF8StringValue;
  end;
end;
{$endif}



end.

