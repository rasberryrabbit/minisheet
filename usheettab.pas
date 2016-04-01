unit usheetTab;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, ComCtrls,
  uCellFormulaPaser, fpspreadsheet;

type

  { TCellTabSheet }

  TCellTabSheet=class(TTabSheet)
    private
      //FWorkBook:TsWorkbook;
      //FSheet:TsWorksheet;
      FSolver:TCellFormula;
    protected
      procedure Notification(AComponent: TComponent; Operation: TOperation);
        override;
    public
      Filename:string;

      constructor Create(TheOwner: TComponent); override;
      destructor Destroy; override;

      function SolveFormula(var s: string): Integer;

      function CheckClose(DefAction:Integer):Integer;

      //property WorkBook:TsWorkbook read FWorkBook;
      //property Sheet:TsWorksheet read FSheet;
      property Solver:TCellFormula read FSolver;
  end;


implementation

uses uGridCell, Dialogs, Controls;

resourcestring
  rsSheetSIsModi = '"%s" is modified. Save this sheet?';


{ TCellTabSheet }

procedure TCellTabSheet.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  if AComponent is TStrGridCell then begin
    if Operation=opInsert then
      if Solver<>nil then
        Solver.Grid:=TStrGridCell(AComponent);
    if Operation=opRemove then
      if Solver<>nil then
        if Solver.Grid=AComponent then
          Solver.Grid:=nil;
  end;
  inherited Notification(AComponent, Operation);
end;

constructor TCellTabSheet.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  //FSheet:=nil;
  Filename:='';
  FSolver:=TCellFormula.Create;
  //FWorkBook:=TsWorkbook.Create;
end;

destructor TCellTabSheet.Destroy;
begin
  FSolver.Free;
  //FWorkBook.Free;
  inherited Destroy;
end;

function TCellTabSheet.SolveFormula(var s: string): Integer;
var
  rets:string;
  //debug:string;
begin
  Result:=-1;
  try
    FSolver.Data.Data:=self;
    FSolver.ClearCellRef;            // clear formula cell references
    FSolver.Data.DataFunc:=FSolver;  // root solver
    Result:=FSolver.DoGenerateRPNQueue(s);
    if Result=0 then begin
      //debug:=FSolver.DumpRPNQueue;
      Result:=FSolver.SolveRPNQueue(rets);
    end;
    if Result=0 then
      s:=rets;
  except
  end;
end;

// defaction = 0(query)
function TCellTabSheet.CheckClose(DefAction: Integer): Integer;
begin
  Result:=DefAction;
  if Solver.Grid.Modified then begin
    if (DefAction in [mrYes,mrNo,mrNone]) then begin
      if Owner<>nil then begin
        if TPageControl(Owner).PageCount>1 then
          Result:=QuestionDlg('minisheet',
                            Format(rsSheetSIsModi, [Caption]),
                            mtConfirmation,
                            [mrYes,mrNo,mrCancel,mrYesToAll,mrNoToAll],
                            0)
          else
          Result:=QuestionDlg('minisheet',
                            Format(rsSheetSIsModi, [Caption]),
                            mtConfirmation,
                            [mrYes,mrNo,mrCancel],
                            0);
      end else
        Result:=mrNo;
    end;
  end;
end;

end.

