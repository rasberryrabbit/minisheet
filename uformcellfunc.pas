unit uformcellfunc;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls,
  Grids;

type

  { TFormCellFunc }

  TFormCellFunc = class(TForm)
    Button1: TButton;
    Button2: TButton;
    CBFunc: TComboBox;
    CheckBox1: TCheckBox;
    EditCell: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure CheckBox1Change(Sender: TObject);
    procedure EditCellChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
    Grid:TStringGrid;
    DestCell:TPoint;
  end;

var
  FormCellFunc: TFormCellFunc;

implementation

uses uCellFormulaPaser;

{$R *.lfm}

{ TFormCellFunc }

procedure TFormCellFunc.FormShow(Sender: TObject);
begin
  CheckBox1Change(nil);
end;

procedure TFormCellFunc.EditCellChange(Sender: TObject);
var
  crow:Integer;
  s:string;
begin
  s:=UpperCase(EditCell.Text);
  if (s<>'') and (_CheckCellPos(s,1,length(s))>length(s)) then begin
    crow:=AlphaToColRow(s);
    DestCell.x:=(crow shr 16)+1;
    DestCell.y:=(crow and $FFFF)+1;
  end else
    DestCell.y:=0;
  if DestCell.x>=Grid.ColCount then
     DestCell.y:=0;
  if DestCell.y>=Grid.RowCount then
     DestCell.y:=0;
  Button1.Enabled:=DestCell.y>0;
  if ((DestCell.x>=Grid.Selection.Left) and (DestCell.x<=Grid.Selection.Right)) and
     ((DestCell.y>=Grid.Selection.Top) and (DestCell.y<=Grid.Selection.Bottom)) then
     begin
       Label2.Font.Color:=clRed;
       Button1.Enabled:=False;
     end else begin
       Label2.Font.Color:=clBlack;
       Button1.Enabled:=True;
     end;
end;

procedure TFormCellFunc.CheckBox1Change(Sender: TObject);
var
  i,j:Integer;
  rc:TGridRect;
begin
  rc:=Grid.Selection;
  DestCell.x:=rc.Right;
  DestCell.y:=0;
  i:=rc.Bottom;
  if (not CheckBox1.Checked) and
      (rc.Bottom>=Grid.RowCount-1) then
    if (DestCell.x<Grid.ColCount-1) then begin
      Inc(DestCell.x);
      i:=rc.Top;
    end;
  if (i<>rc.Top) and (i<Grid.RowCount-1) then
    Inc(i)
    else begin
      i:=rc.Top;
      if rc.Right<(Grid.ColCount-1) then
        DestCell.x:=rc.Right+1;
    end;
  j:=1;
  while j>0 do begin
    while i<Grid.RowCount do begin
      if Grid.Cells[DestCell.x,i]='' then begin
        DestCell.y:=i;
        break;
      end;
      inc(i);
    end;
    if DestCell.y=0 then begin
      if DestCell.x<>rc.Right then
        DestCell.x:=rc.Right
        else
          if rc.Right<Grid.ColCount-1 then
            DestCell.x:=rc.Right+1;
      if DestCell.x<>rc.Right then
        i:=rc.Top
        else
          if rc.Bottom<Grid.RowCount-1 then
            i:=rc.Bottom+1
            else break;
    end else break;
    Dec(j);
  end;
  if DestCell.y=0 then
    DestCell.y:=rc.Top;
  EditCell.Text:=MakeColumnStr(DestCell.x-1)+IntToStr(DestCell.y);
end;

end.

