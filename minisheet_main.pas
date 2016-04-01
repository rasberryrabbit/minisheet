unit minisheet_main;

{$mode objfpc}{$H+}

{$i usheettabconf.inc}

interface

uses
  Classes, SysUtils, FileUtil, SynEdit, SynCompletion, ExtendedNotebook,
  LR_PGrid, LR_Class, LR_DSet, Forms, Controls, Graphics, Dialogs, ActnList,
  Menus, ExtCtrls, StdCtrls, Buttons, ComCtrls, Grids, StdActns, usheetTab,
  uGridCell, types, LCLType;

type

  { TFormMain }

  TFormMain = class(TForm)
    ActionUnDelCell: TAction;
    ActionRedoDelete: TAction;
    ActionPrintNoFrame: TAction;
    ActionPrintNoValidRow: TAction;
    ActionPrintColLeft: TAction;
    ActionUndelete: TAction;
    ActionCutCells: TAction;
    ActionPasteCell: TAction;
    ActionImportExcel: TAction;
    ActionExportExcel: TAction;
    ActionAutoDown: TAction;
    ActionSetCSVDelimiter: TAction;
    ActionExportCSV: TAction;
    ActionPrint: TAction;
    ActionSheetNext: TAction;
    ActionSheetPrev: TAction;
    ActionCopyCells: TAction;
    ActionCopyCalced: TAction;
    FileOpen1: TFileOpen;
    frReport1: TfrReport;
    frUserDataset1: TfrUserDataset;
    GridFuncCells: TAction;
    ActionAutoCellWidth2: TAction;
    ActionAutoCellWidth: TAction;
    ActionCloseSheet: TAction;
    ActionNewSheet: TAction;
    ActionList1: TActionList;
    FileSaveAs1: TFileSaveAs;
    ImageList1: TImageList;
    MenuItem10: TMenuItem;
    MenuItem11: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem19: TMenuItem;
    MenuItem20: TMenuItem;
    MenuItem21: TMenuItem;
    MenuItem22: TMenuItem;
    MenuItem23: TMenuItem;
    MenuItem24: TMenuItem;
    MenuItem25: TMenuItem;
    MenuItem26: TMenuItem;
    MenuItem27: TMenuItem;
    MenuItem28: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    MenuItem32: TMenuItem;
    MenuItem33: TMenuItem;
    MenuItem34: TMenuItem;
    MenuItem35: TMenuItem;
    MenuItem36: TMenuItem;
    MenuItem37: TMenuItem;
    MenuItem38: TMenuItem;
    MenuItem39: TMenuItem;
    MenuItem40: TMenuItem;
    MenuItem41: TMenuItem;
    MenuItem42: TMenuItem;
    MenuItem43: TMenuItem;
    MenuItem44: TMenuItem;
    MenuItem45: TMenuItem;
    MenuItem46: TMenuItem;
    MenuItem47: TMenuItem;
    MenuItem48: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem9: TMenuItem;
    NoteSheet: TExtendedNotebook;
    FileExit1: TFileExit;
    MainMenu1: TMainMenu;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    Panel2: TPanel;
    PopupGrid: TPopupMenu;
    PopupTab: TPopupMenu;
    SaveDialog1: TSaveDialog;
    StatusBar1: TStatusBar;
    SynCompletion1: TSynCompletion;
    SynEdit1: TSynEdit;
    ToolBar1: TToolBar;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    procedure ActionAutoCellWidth2Execute(Sender: TObject);
    procedure ActionAutoCellWidthExecute(Sender: TObject);
    procedure ActionAutoDownExecute(Sender: TObject);
    procedure ActionCloseSheetExecute(Sender: TObject);
    procedure ActionCopyCalcedExecute(Sender: TObject);
    procedure ActionCopyCellsExecute(Sender: TObject);
    procedure ActionCutCellsExecute(Sender: TObject);
    procedure ActionExportCSVExecute(Sender: TObject);
    procedure ActionExportExcelExecute(Sender: TObject);
    procedure ActionImportExcelExecute(Sender: TObject);
    procedure ActionList1Update(AAction: TBasicAction; var Handled: Boolean);
    procedure ActionNewSheetExecute(Sender: TObject);
    procedure ActionPasteCellExecute(Sender: TObject);
    procedure ActionPrintExecute(Sender: TObject);
    procedure ActionPrintNoValidRowExecute(Sender: TObject);
    procedure ActionRedoDeleteExecute(Sender: TObject);
    procedure ActionRedoDeleteUpdate(Sender: TObject);
    procedure ActionSetCSVDelimiterExecute(Sender: TObject);
    procedure ActionSheetNextExecute(Sender: TObject);
    procedure ActionSheetPrevExecute(Sender: TObject);
    procedure ActionUnDelCellExecute(Sender: TObject);
    procedure ActionUndeleteExecute(Sender: TObject);
    procedure ActionUndeleteUpdate(Sender: TObject);
    procedure FileOpen1Accept(Sender: TObject);
    procedure FileSaveAs1Accept(Sender: TObject);
    procedure FileSaveAs1BeforeExecute(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormDropFiles(Sender: TObject; const FileNames: array of String);
    procedure FormShow(Sender: TObject);
    procedure frReport1EnterRect(Memo: TStringList; View: TfrView);
    procedure frReport1GetValue(const ParName: String; var ParValue: Variant);
    procedure frUserDataset1CheckEOF(Sender: TObject; var Eof: Boolean);
    procedure frUserDataset1First(Sender: TObject);
    procedure frUserDataset1Next(Sender: TObject);
    procedure GridColRowDeleted(Sender: TObject; IsColumn: Boolean;
      sIndex, tIndex: Integer);
    procedure GridFuncCellsExecute(Sender: TObject);
    procedure GridSelectCell(Sender: TObject; aCol, aRow: Integer;
      var CanSelect: Boolean);
    procedure NoteSheetChange(Sender: TObject);
    procedure NoteSheetMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GridMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure GridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SynEdit1Enter(Sender: TObject);
    procedure SynEdit1Exit(Sender: TObject);
    procedure SynEdit1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState
      );
  private
    { private declarations }
    APage:TTabSheet;
  public
    NewEdit,OldEdit:TSynEdit;

    { public declarations }
    procedure CreateNewSheet;
    function GetNewSheetName:string;
    function GetWorkSheet(Grid:TStrGridCell):TCellTabSheet;
    function GetGrid(Tab:TTabSheet):TStrGridCell;
    function GetValidCell(Grid:TStrGridCell;aCol,aRow:Integer):string;
    procedure UpdateFileCaption(s:string);
    procedure ReCreateEdit;
  end;

var
  FormMain: TFormMain;

implementation

uses uCellFormulaPaser, fpspreadsheet, fpsallformats, uformcellfunc, DefaultTranslator,
  Translations, gettext, LCLMessageGlue;

{$R *.lfm}

{$R minisheet_rep.rc}

resourcestring
  rsSIsEqualToDe = '"%s" is equal to DecimalSeparator';
  rsInputCSVDeli = 'Input CSV Delimiter character';

var
  CSVDelimiter:char=',';
  ReportRow,ReportRowValid,ReportStartCol:Integer;

{ TFormMain }


procedure TFormMain.SynEdit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key=VK_RETURN) or (Key=VK_UP) or (Key=VK_DOWN) or (Key=VK_TAB) then begin
    with TCellTabSheet(NoteSheet.ActivePage).Solver.Grid do begin
      Cells[Col,Row]:=TrimRight(NewEdit.Lines.Text);
      EditorMode:=False;
      SetFocus;
    end;
    if Key<>VK_RETURN then
      LCLSendKeyDownEvent(TCellTabSheet(NoteSheet.ActivePage).Solver.Grid,Key,1,True,False);
    Key:=0;
  end else if Key=VK_ESCAPE then begin
    Key:=0;
    with TCellTabSheet(NoteSheet.ActivePage).Solver.Grid do begin
      NewEdit.LineText:=TrimRight(Cells[Col,Row]);
      EditorMode:=False;
      SetFocus;
    end;
  end else if Key=VK_LEFT then begin
    if (NewEdit.SelStart=1) and (NewEdit.SelEnd<2) then begin
      with TCellTabSheet(NoteSheet.ActivePage).Solver.Grid do begin
        Cells[Col,Row]:=TrimRight(NewEdit.Lines.Text);
        EditorMode:=False;
        SetFocus;
      end;
      LCLSendKeyDownEvent(TCellTabSheet(NoteSheet.ActivePage).Solver.Grid,Key,1,True,False);
      Key:=0;
    end;
  end;
end;

procedure TFormMain.CreateNewSheet;
var
  Sheet:TCellTabSheet;
  NewGrid:TStrGridCell;
  s:string;
begin
  s:=GetNewSheetName;
  Sheet:=TCellTabSheet.Create(NoteSheet);
  try
    Sheet.Name:=s;
    Sheet.PageControl:=NoteSheet;

    NewGrid:=TStrGridCell.Create(Sheet);
    NewGrid.Name:='Grid_'+s;
    NewGrid.Parent:=Sheet;
    NewGrid.Align:=alClient;
    NewGrid.Options:=[goDrawFocusSelected,goAutoAddRows,goAutoAddRowsSkipContentCheck,goCellEllipsis,
          goColSizing,goEditing,goFixedColSizing,goFixedRowNumbering,goFixedHorzLine,
          goFixedVertLine,goRangeSelect,goTabs,goSmoothScroll,goVertLine,goHorzLine];
    NewGrid.DefaultColWidth:=80;
    NewGrid.OnColRowDeleted:=@GridColRowDeleted;
    NewGrid.OnSelectCell:=@GridSelectCell;
    NewGrid.OnMouseMove:=@GridMouseMove;
    NewGrid.OnKeyDown:=@GridKeyDown;
    NewGrid.PopupMenu:=PopupGrid;
    NewGrid.Flat:=True;
    NewGrid.UseXORFeatures:=True;
    NewGrid.Color:=clWhite;
    NewGrid.AutoSizeCell:=ActionAutoCellWidth2.Checked;
    NewGrid.ColCount:=8;
    NewGrid.RowCount:=30;
    UpdateColLabel(NewGrid,-1,-1);
    if ActionAutoDown.Checked then
      NewGrid.AutoAdvance:=aaDown
      else
        NewGrid.AutoAdvance:=aaRight;
    NewGrid.Modified:=False;
    NoteSheet.ActivePage:=Sheet;
  except
    Sheet.Free;
  end;
end;

function TFormMain.GetNewSheetName: string;
var
  i,j:Integer;
  s:string;
  HasName:Boolean;
begin
  i:=1;
  while i<10000 do begin
    s:='Sheet'+IntToStr(i);
    HasName:=False;
    for j:=0 to NoteSheet.PageCount-1 do begin
      if NoteSheet.Pages[j].Caption=s then begin
        HasName:=True;
        break;
      end;
    end;
    if not HasName then
      break;
    Inc(i);
  end;
  Result:=s;
end;

function TFormMain.GetWorkSheet(Grid: TStrGridCell): TCellTabSheet;
begin
  if Grid<>nil then
    Result:=TCellTabSheet(Grid.Owner);
end;

function TFormMain.GetGrid(Tab: TTabSheet): TStrGridCell;
var
  i:Integer;
begin
  Result:=nil;
  for i:=0 to Tab.ComponentCount-1 do
    if Tab.Components[i] is TStrGridCell then begin
      Result:=Tab.Components[i] as TStrGridCell;
      break;
    end;
end;

function TFormMain.GetValidCell(Grid: TStrGridCell; aCol, aRow: Integer): string;
begin
  Result:='';
  if not ((aCol>=Grid.ColCount) or (aRow>=Grid.RowCount)) then begin
    Result:=Grid.Cells[aCol,aRow];
  end;
end;

procedure TFormMain.UpdateFileCaption(s: string);
begin
  StatusBar1.Panels.Items[2].Text:=ExtractFileName(s);
end;

procedure TFormMain.ReCreateEdit;
var
  i:Integer;
begin
  if OldEdit<>nil then
    FreeAndNil(OldEdit);
  for i:=0 to Panel2.ControlCount-1 do
    if Panel2.Controls[i] is TSynEdit then
      OldEdit:=TSynEdit(Panel2.Controls[i]);
  newedit:=TSynEdit.Create(self);
  newedit.Parent:=Panel2;
  newedit.WantTabs:=False;
  newedit.ScrollBars:=ssNone;
  newedit.RightGutter.Visible:=False;
  newedit.Gutter.Visible:=False;
  newedit.Options:=newedit.Options+[eoHideRightMargin];
  newedit.Align:=alClient;
  newedit.OnEnter:=@SynEdit1Enter;
  newedit.OnExit:=@SynEdit1Exit;
  newedit.OnKeyDown:=@SynEdit1KeyDown;
  SynCompletion1.Editor:=NewEdit;
end;

procedure TFormMain.GridFuncCellsExecute(Sender: TObject);
const
  sumfunc='%s(%s:%s)';
var
  s:string;
  grect:TGridRect;
  cellfunc:TFormCellFunc;
  grid:TStringGrid;
begin
  grid:=TCellTabSheet(NoteSheet.ActivePage).Solver.Grid;
  grect:=grid.Selection;

  cellfunc:=TFormCellFunc.Create(self);
  try
    cellfunc.Grid:=grid;
    if mrOK=cellfunc.ShowModal then begin
      if cellfunc.DestCell.y>0 then begin
        s:=format(sumfunc,[cellfunc.CBFunc.Text,MakeColumnStr(grect.Left-1)+IntToStr(grect.Top),
          MakeColumnStr(grect.Right-1)+IntToStr(grect.Bottom)]);
        grid.Cells[cellfunc.DestCell.x,cellfunc.DestCell.y]:=s;
      end;
    end;
  finally
    cellfunc.Free;
  end;
end;

procedure TFormMain.GridSelectCell(Sender: TObject; aCol, aRow: Integer;
  var CanSelect: Boolean);
{$ifdef USE_SPREADCELL}
var
  cell:PCell;
  works:TsWorksheet;
{$endif}
begin
{$ifdef USE_SPREADCELL}
  works:=GetWorkSheet(Sender as TStringGrid).Sheet;
  if WorkS<>nil then begin
    cell:=WorkS.GetCell(aRow-1,aCol-1);
    SynEdit1.Text:=cell^.UTF8StringValue;
  end;
{$else}
  NewEdit.Text:=(Sender as TStringGrid).Cells[aCol,aRow];
  StatusBar1.Panels[0].Text:=MakeColumnStr(aCol-1)+IntToStr(aRow);
{$endif}
end;

procedure TFormMain.NoteSheetChange(Sender: TObject);
begin
  if GetGrid(NoteSheet.ActivePage)<>nil then begin
    GetGrid(NoteSheet.ActivePage).AutoSizeCell:=ActionAutoCellWidth2.Checked;
    if ActionAutoDown.Checked then
      GetGrid(NoteSheet.ActivePage).AutoAdvance:=aaDown
      else
        GetGrid(NoteSheet.ActivePage).AutoAdvance:=aaRight;
  end;
  UpdateFileCaption(TCellTabSheet(NoteSheet.ActivePage).Filename);
end;


procedure TFormMain.NoteSheetMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
  i:Integer;
  p:TPoint;
begin
  i:=NoteSheet.TabIndexAtClientPos(Point(X,Y));
  if i<>-1 then begin
    NoteSheet.ActivePageIndex:=i;
    //GetGrid(NoteSheet.ActivePage).SetFocus;
    if Button = mbRight then begin
      p:=Point(X,Y);
      p:=ClientToScreen(p);
      PopupTab.PopUp(p.x+10,p.y-20);
    end;
  end;
end;

procedure TFormMain.GridMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
var
  p:TPoint;
begin
  {
  p:=(Sender as TStringGrid).MouseToCell(Point(X,Y));
  if (p.x>0) and (p.y>0) then
    StatusBar1.Panels[1].Text:=Format('%s%d',[MakeColumnStr(p.x-1),p.y])
    else
      StatusBar1.Panels[1].Text:='';
  }
end;

procedure TFormMain.GridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (not TStrGridCell(Sender).EditorMode) then begin
    if (Key=VK_F2) then begin
      NewEdit.SetFocus;
      NewEdit.SelectAll;
      TStrGridCell(Sender).AddEditHistory;
      Key:=0;
    end;
  end;
end;

procedure TFormMain.SynEdit1Enter(Sender: TObject);
begin
  APage:=NoteSheet.ActivePage;
end;

procedure TFormMain.SynEdit1Exit(Sender: TObject);
begin
  if NoteSheet.IndexOf(APage)<>-1 then
    with GetGrid(APage) do
      Cells[Col,Row]:=NewEdit.Text;
  ReCreateEdit;
  with GetGrid(APage) do
    NewEdit.Text:=Cells[Col,Row];
end;

procedure TFormMain.FormShow(Sender: TObject);
begin
  OldEdit:=nil;
  ReCreateEdit;
  CreateNewSheet;
end;

function GetColNo(const s:string):string;
var
  i,l:Integer;
begin
  Result:='';
  l:=length(s);
  i:=1;
  while i<=l do begin
    if s[i] in ['0'..'9'] then
      Result:=Result+s[i]
      else
        if Result<>'' then
          break;
    inc(i);
  end;
end;

procedure TFormMain.frReport1EnterRect(Memo: TStringList; View: TfrView);
var
  me:TfrMemoView;
  txt:string;
  aCol:Integer;
  grid:TStrGridCell;
begin
  if View is TfrMemoView then begin
    me:=TfrMemoView(View);
    aCol:=StrToIntDef(GetColNo(me.Memo.Text),0);
    grid:=GetGrid(NoteSheet.ActivePage);
    if (aCol>0) and (aCol<grid.ColCount) then begin
      txt:=grid.Cells[aCol,ReportRow];
      if txt<>'' then
        // text justify
        if not (txt[1] in ['''','>','<']) then begin
          me.Alignment:=taRightJustify;
          if not GetWorkSheet(grid).Solver.CheckNumber(txt) then
            if 0<>GetWorkSheet(grid).SolveFormula(txt) then
              me.Alignment:=taLeftJustify;
        end
        else begin
          if me<>nil then
            if txt[1]='>' then
              me.Alignment:=taRightJustify
              else
                if txt[1]='<' then
                  me.Alignment:=taCenter
                  else
                    me.Alignment:=taLeftJustify;
          Delete(txt,1,1);
        end;
    end;
  end;
end;

procedure TFormMain.frReport1GetValue(const ParName: String;
  var ParValue: Variant);
var
  grid:TStrGridCell;
  aCol:Integer;
  txt:string;
begin
  if CompareText(Copy(ParName,1,4),'cell')=0 then begin
    aCol:=StrToIntDef(Copy(ParName,5,5),0);
    grid:=GetGrid(NoteSheet.ActivePage);
    if (aCol>0) and (aCol<grid.ColCount) then begin
        txt:=grid.Cells[aCol,ReportRow];
        if txt<>'' then
          if not (txt[1] in ['''','>','<']) then begin
            if not GetWorkSheet(grid).Solver.CheckNumber(txt) then
              GetWorkSheet(grid).SolveFormula(txt);
          end
          else begin
            Delete(txt,1,1);
          end;
        ParValue:=txt;
      end else
        ParValue:='';
  end;
end;

procedure TFormMain.frUserDataset1CheckEOF(Sender: TObject; var Eof: Boolean);
begin
  Eof:=ReportRow>ReportRowValid;
end;

procedure TFormMain.frUserDataset1First(Sender: TObject);
begin
  ReportRow:=1;
end;

procedure TFormMain.frUserDataset1Next(Sender: TObject);
begin
  Inc(ReportRow);
end;

procedure TFormMain.GridColRowDeleted(Sender: TObject; IsColumn: Boolean;
  sIndex, tIndex: Integer);
begin

end;

procedure TFormMain.FormCreate(Sender: TObject);
var
  lng,lngf:string;
begin
  if DecimalSeparator=',' then
    CSVDelimiter:=';';

  GetLanguageIDs(lng,lngf);
  Translations.TranslateUnitResourceStrings('LCLStrConsts', 'lclstrconsts.%s.po', lng, lngf);
end;

procedure TFormMain.ActionNewSheetExecute(Sender: TObject);
begin
  CreateNewSheet;
end;

procedure TFormMain.ActionPasteCellExecute(Sender: TObject);
begin
  GetGrid(NoteSheet.ActivePage).PasteFromClipboardCell;
end;

procedure TFormMain.ActionPrintExecute(Sender: TObject);
var
  memo:TfrMemoView;
  band:TfrBandView;
  lwidth,lleft,ltop:double;
  ipos,rmargin:Integer;
  grid:TStrGridCell;
  pdata:TResourceStream;
begin
  pdata:=TResourceStream.Create(HINSTANCE,'REPORTDATA',RT_RCDATA);
  try
    frReport1.LoadFromXMLStream(pdata);
  finally
    pdata.Free;
  end;
  //frReport1.LoadFromFile('minireport.lrf');
  band:=TfrBandView(frReport1.FindObject('masterdata1'));
  if band<>nil then begin
    lleft:=frReport1.Pages[0].LeftMargin;
    rmargin:=frReport1.Pages[0].RightMargin;
    ltop:=band.Top;
    grid:=GetGrid(NoteSheet.ActivePage);
    lwidth:=grid.DefaultColWidth;
    if ActionPrintNoValidRow.Checked then
      ReportRowValid:=grid.RowCount-1
      else
        ReportRowValid:=grid.RowValidMax;
    // column start
    if Sender=ActionPrintColLeft then
      ReportStartCol:=grid.LeftCol
      else
        ReportStartCol:=1;
    ipos:=ReportStartCol;
    // add memoview
    while lleft<rmargin do begin
      lwidth:=grid.ColWidths[ipos]+10;
      if lleft+lwidth>rmargin then
        lwidth:=rmargin-lleft;
      // 1
      memo:=TfrMemoView(frCreateObject(gtMemo,'',frReport1.Pages[0]));
      memo.CreateUniqueName;
      memo.Left:=lleft;
      memo.Top:=ltop;
      memo.Width:=lwidth;
      memo.Alignment:=taRightJustify;
      memo.Height:=18;
      if not ActionPrintNoFrame.Checked then
        memo.Frames:=[frbLeft,frbTop,frbRight,frbBottom]
        else
          memo.Frames:=[];
      memo.Memo.Text:=format('[cell%d]',[ipos]);
      memo.WordWrap:=False;
      frReport1.Pages[0].Objects.Add(memo);
      Inc(ipos);
      lleft:=lleft+lwidth;
    end;
  end;
  frReport1.ShowReport;
end;

procedure TFormMain.ActionPrintNoValidRowExecute(Sender: TObject);
begin
  TAction(Sender).Checked:=not TAction(Sender).Checked;
end;

procedure TFormMain.ActionRedoDeleteExecute(Sender: TObject);
begin
  GetGrid(NoteSheet.ActivePage).RedoDelete;
end;

procedure TFormMain.ActionRedoDeleteUpdate(Sender: TObject);
begin
  with GetGrid(NoteSheet.ActivePage) do begin
    TAction(Sender).Enabled:=(not EditorMode) and (RedoHistory.Count>0);
    //TAction(Sender).Caption:=Format('Undelete %d',[GetGrid(NoteSheet.ActivePage).EditHistory.Count]);
  end;
end;

procedure TFormMain.ActionSetCSVDelimiterExecute(Sender: TObject);
var
  newval:string;
begin
  newval:=CSVDelimiter;
  if InputQuery('minisheet', rsInputCSVDeli, newval) then begin
    if newval<>'' then
      if newval<>DecimalSeparator then
        CSVDelimiter:=newval[1]
        else
          ShowMessageFmt(rsSIsEqualToDe, [newval]);
  end;
end;

procedure TFormMain.ActionSheetNextExecute(Sender: TObject);
var i:Integer;
begin
  i:=NoteSheet.ActivePageIndex;
  Inc(i);
  if i>=NoteSheet.PageCount then
    i:=0;
  NoteSheet.PageIndex:=i;
end;

procedure TFormMain.ActionSheetPrevExecute(Sender: TObject);
var i:Integer;
begin
  i:=NoteSheet.ActivePageIndex;
  Dec(i);
  if i<0 then
    i:=NoteSheet.PageCount-1;
  NoteSheet.PageIndex:=i;
end;

procedure TFormMain.ActionUnDelCellExecute(Sender: TObject);
begin
  GetGrid(NoteSheet.ActivePage).UnDeleteCell;
end;

procedure TFormMain.ActionUndeleteExecute(Sender: TObject);
begin
  GetGrid(NoteSheet.ActivePage).UnDelete;
end;

procedure TFormMain.ActionUndeleteUpdate(Sender: TObject);
begin
  with GetGrid(NoteSheet.ActivePage) do begin
    TAction(Sender).Enabled:=(not EditorMode) and (EditHistory.Count>0);
    //TAction(Sender).Caption:=Format('Undelete %d',[GetGrid(NoteSheet.ActivePage).EditHistory.Count]);
  end;
end;

procedure TFormMain.FileOpen1Accept(Sender: TObject);
begin
  try
    if GetGrid(NoteSheet.ActivePage).Modified or (TCellTabSheet(NoteSheet.ActivePage).Filename<>'') then
      CreateNewSheet;
    GetGrid(NoteSheet.ActivePage).LoadFromCSVFile(Utf8ToAnsi(FileOpen1.Dialog.FileName),CSVDelimiter);
    TCellTabSheet(NoteSheet.ActivePage).Filename:=FileOpen1.Dialog.FileName;
    GetGrid(NoteSheet.ActivePage).Modified:=False;
    UpdateFileCaption(TCellTabSheet(NoteSheet.ActivePage).Filename);
  except
    on e:Exception do begin
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TFormMain.ActionCloseSheetExecute(Sender: TObject);
var
  Grid:TStrGridCell;
  ix:Integer;
  sfile:string;
begin
  Grid:=GetGrid(NoteSheet.ActivePage);
  ix:=mrNone;
  if Grid<>nil then
    if Grid.Modified then begin
      ix:=GetWorkSheet(Grid).CheckClose(mrNone);
      if ix in [mrYes,mrYesToAll] then begin
        sfile:=GetWorkSheet(Grid).Filename;
        try
        if sfile<>'' then
          grid.SaveToCSVFile(Utf8ToAnsi(sfile),CSVDelimiter)
          else
            FileSaveAs1.Execute;
        except
          on e:Exception do begin
            ShowMessage(e.Message);
          end;
        end;
      end;
    end;
  if ix<>mrCancel then
    NoteSheet.ActivePage.Free;
  if NoteSheet.PageCount=0 then
    CreateNewSheet;
  GetGrid(NoteSheet.ActivePage).SetFocus;
end;

procedure TFormMain.ActionCopyCalcedExecute(Sender: TObject);
begin
  TCellTabSheet(NoteSheet.ActivePage).Solver.Grid.CopyToClipboard(True);
end;

procedure TFormMain.ActionCopyCellsExecute(Sender: TObject);
begin
  TCellTabSheet(NoteSheet.ActivePage).Solver.Grid.CopySolved:=False;
  TCellTabSheet(NoteSheet.ActivePage).Solver.Grid.CopyToClipboard(True);
  TCellTabSheet(NoteSheet.ActivePage).Solver.Grid.CopySolved:=True;
end;

procedure TFormMain.ActionCutCellsExecute(Sender: TObject);
begin
  TCellTabSheet(NoteSheet.ActivePage).Solver.Grid.CutToClipboardCell;
end;

procedure TFormMain.ActionExportCSVExecute(Sender: TObject);
begin
  SaveDialog1.FileName:=ChangeFileExt(TCellTabSheet(NoteSheet.ActivePage).Filename,'_export.csv');
  try
    if SaveDialog1.Execute then
      GetGrid(NoteSheet.ActivePage).ExportToCSVFile(Utf8ToAnsi(SaveDialog1.FileName));
  except
    on e:Exception do begin
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TFormMain.ActionExportExcelExecute(Sender: TObject);
begin
  SaveDialog1.FileName:=ChangeFileExt(TCellTabSheet(NoteSheet.ActivePage).Filename,'_export.xls');
  try
    if SaveDialog1.Execute then
      GetGrid(NoteSheet.ActivePage).ExportToExcelFile(Utf8ToAnsi(SaveDialog1.FileName));
  except
    on e:Exception do begin
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TFormMain.ActionImportExcelExecute(Sender: TObject);
begin
  try
    if OpenDialog1.Execute then begin
      GetGrid(NoteSheet.ActivePage).ImportFromExcelFile(Utf8ToAnsi(OpenDialog1.FileName));
    end;
  except
    on e:Exception do begin
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TFormMain.ActionList1Update(AAction: TBasicAction;
  var Handled: Boolean);
begin

end;

procedure TFormMain.ActionAutoCellWidthExecute(Sender: TObject);
begin
  TCellTabSheet(NoteSheet.ActivePage).Solver.Grid.AutoSizeColumns;
end;

procedure TFormMain.ActionAutoDownExecute(Sender: TObject);
begin
  TAction(Sender).Checked:=not TAction(Sender).Checked;
  if TAction(Sender).Checked then
    GetGrid(TTabSheet(NoteSheet.ActivePage)).AutoAdvance:=aaDown
    else
      GetGrid(TTabSheet(NoteSheet.ActivePage)).AutoAdvance:=aaRight;
end;

procedure TFormMain.ActionAutoCellWidth2Execute(Sender: TObject);
begin
  TAction(Sender).Checked:=not TAction(Sender).Checked;
  GetGrid(TTabSheet(NoteSheet.ActivePage)).AutoSizeCell:=TAction(Sender).Checked;
end;

procedure TFormMain.FileSaveAs1Accept(Sender: TObject);
begin
  GetGrid(NoteSheet.ActivePage).SaveToCSVFile(Utf8ToAnsi(FileSaveAs1.Dialog.FileName),CSVDelimiter);
  GetGrid(NoteSheet.ActivePage).Modified:=False;
  TCellTabSheet(NoteSheet.ActivePage).Filename:=FileSaveAs1.Dialog.FileName;
  UpdateFileCaption(TCellTabSheet(NoteSheet.ActivePage).Filename);
end;

procedure TFormMain.FileSaveAs1BeforeExecute(Sender: TObject);
begin
  FileSaveAs1.Dialog.FileName:=TCellTabSheet(NoteSheet.ActivePage).Filename;
end;

procedure TFormMain.FormCloseQuery(Sender: TObject; var CanClose: boolean);
var
  ix,i:Integer;
  grid:TStrGridCell;
  sheet:TTabSheet;
  sfile:string;
begin
  ix:=mrNone;
  while NoteSheet.PageCount>0 do begin
    grid:=GetGrid(NoteSheet.ActivePage);
    if Grid<>nil then
      if Grid.Modified then begin
        ix:=GetWorkSheet(Grid).CheckClose(ix);
        if ix=mrCancel then
          break;
        if ix in [mrYes,mrYesToAll] then begin
          sfile:=GetWorkSheet(Grid).Filename;
          try
          if sfile<>'' then
            grid.SaveToCSVFile(Utf8ToAnsi(sfile),CSVDelimiter)
            else
              FileSaveAs1.Execute;
          except
            on e:Exception do begin
              ShowMessage(e.Message);
              break;
            end;
          end;
        end;
      end;
    NoteSheet.ActivePage.Free;
  end;
  CanClose:=NoteSheet.PageCount=0;
end;

procedure TFormMain.FormDestroy(Sender: TObject);
begin

end;

procedure TFormMain.FormDropFiles(Sender: TObject;
  const FileNames: array of String);
var
  i:Integer;
begin
  for i:=Low(FileNames) to High(FileNames) do begin
    try
      CreateNewSheet;
      FileOpen1.Dialog.FileName:=FileNames[i];
      FileOpen1Accept(nil);
    except
      on e:Exception do
        ShowMessage(e.Message);
    end;
  end;
  GetGrid(NoteSheet.ActivePage).SetFocus;
end;


end.

