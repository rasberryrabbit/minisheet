unit uCellFormulaPaser;

{$mode objfpc}{$H+}

{$i usheettabconf.inc}

interface

uses
  Classes, SysUtils, uSimplefmParser, uGridCell;

type

  { TCellFormula }

  TCellFormula=class(TSimplefmParser)
    private
      FGrid:TStrGridCell;
      FRefList:TFPList;
      function AtToken(Token:TSimpleTokenVar;const str:string;CPos:Integer;var NextPos:Integer):boolean;
    protected
    public
      constructor Create;
      destructor Destroy; override;

      // check cell reference contains formula
      procedure AddCellRef(cell:Pointer);
      procedure DelCellRef(cell:Pointer);
      function CheckRef(cell:Pointer):Integer;
      procedure ClearCellRef;

      function CheckNumber(const s:string):Boolean;

      property Grid:TStrGridCell read FGrid write FGrid;
  end;


  function ConvertCell(const BasePos,CurrPos:TPoint;const str:string):string;

  function MakeColumnStr(index:Integer):string;
  function AlphaToColRow(const s:string):Integer;
  function _CheckCellPos(const cellstr: string;startpos,len:Integer): Integer;

implementation

uses contnrs, fpspreadsheet, minisheet_main, mp_ratio, math;

function ConvertCell(const BasePos,CurrPos: TPoint; const str: string): string;
var
  val:TSimpleTokenVar;
  i,p,j,l,tostart,x,y:Integer;
  tok:string;
  IsRgn:Boolean;
begin
  Result:='';
  l:=length(str);
  i:=1;
  // relative position
  x:=CurrPos.x-BasePos.x;
  y:=CurrPos.y-BasePos.y;

  val:=TSimpleTokenVar.Create;
  try

    while i<>0 do begin
      // skip space
      while i<l do begin
        if str[i]<#32 then
          Result:=Result+str[i]
          else
            break;
        Inc(i);
      end;
      // get token
      tostart:=i;
      i:=GetTokenFromStr(str,tok,i);
      if i=0 then
        break;
      // get type
      try
        val.tokenType:=toNone;
        val.Text:=tok;
      except
        val.tokenType:=toNone;
      end;
      if val.tokenType=toNone then begin
        p:=_CheckCellPos(str,tostart,l);
        if p<>-1 then begin
          IsRgn:=False;
          j:=p;
          // one cell
          p:=AlphaToColRow(UpperCase(Copy(str,tostart,p-tostart)));
          Result:=Result+MakeColumnStr((p shr 16)+x)+IntToStr((p and $ffff)+y+1);
          i:=j;
          // cell range
          if j<l then
            if str[j]=':' then begin
              Inc(j);
              Result:=Result+':';
              if j<l then begin
                p:=_CheckCellPos(str,j,l);
                if p<>-1 then begin
                  IsRgn:=True;
                  i:=p;
                  p:=AlphaToColRow(UpperCase(Copy(str,j,p-j)));
                  Result:=Result+MakeColumnStr((p shr 16)+x)+IntToStr((p and $ffff)+y+1);
                end;
              end;
            end;
          continue;
        end;
      end;
      Result:=Result+Copy(str,tostart,i-tostart);
    end;
  finally
    val.Free;
  end;
end;

function MakeColumnStr(index:Integer):string;
const
  _AlphaLen = ord('Z')-ord('A')+1;
var
  ix,ia,il:Integer;
begin
  Result:='';
  ix:=index;
  il:=0;
  if ix>-1 then
  repeat
    ia:=ix mod _AlphaLen;
    if il>0 then
      Dec(ia);
    Result:=chr(ord('A')+ia)+Result;
    ix:=ix div _AlphaLen;
    Inc(il);
  until ix=0;
end;

procedure _CheckRange(var x,y:Integer);
var
  z:Integer;
begin
  if x>y then begin
    z:=x;
    x:=y;
    y:=z;
  end;
end;

// A1 -> 0,0
function AlphaToColRow(const s:string):Integer;
const
  __alphacell = ord('Z')-ord('A')+1;
var
  i,l,col,di:Integer;
begin
  l:=Length(s);
  i:=1;
  while i<=l do begin
     if not (s[i] in ['A'..'Z']) then
       break;
     Inc(i);
  end;
  // get row
  Result:=StrToIntDef(Copy(s,i,l-i+1),1)-1;
  Dec(i);
  col:=0;
  l:=1;
  repeat
    di:=ord(s[i])-ord('A');
    if l>1 then // decrease 1 if not lowerest digit
      Inc(di);
    col:=col+di*l;
    l:=l*__alphacell;
    Dec(i);
  until i=0;
  // col max = 255, row max = 65535
  Result:=Result or ((col {$ifdef USE_SPREADCELL} and $ff{$endif}) shl 16);
end;

// read cell value
function _ReadCell(aRow, aCol: Integer; Data: PSimplefmData;MustGet:Boolean=True): string;
var
  solver, rsolver: TCellFormula;
  retv: string;
  cell: {$ifdef USE_SPREADCELL}PCell{$else}Pointer{$endif};
  {$ifdef USE_SPREADCELL}sheet:TsWorksheet;{$endif}
begin
    {$ifdef USE_SPREADCELL}
    sheet:=TsWorksheet(Data^.Data);
    {$endif}
    if (aRow>=0) and (aCol>=0) then begin
      rsolver:=TCellFormula(Data^.DataFunc);
      {$ifdef USE_SPREADCELL}
      cell:=sheet.GetCell(aRow,aCol);
      retv:=cell^.UTF8StringValue;
      {$else}
      Inc(aCol); Inc(aRow); // skip fixed col, row
      cell:=Pointer(PtrInt((aCol shl 16) or (aRow and $FFFF)));
      if (aCol<(rsolver.Grid).ColCount) and (aRow<(rsolver.Grid).RowCount) then
        retv:=(rsolver.Grid).Cells[aCol,aRow]
        else
          retv:='';
      {$endif}
      if ((retv<>'') and (retv[1]='=')) and
         not IsNumericSigned(retv) then begin
        // prevent self reference infinite recursive call
        if rsolver.CheckRef(cell)=-1 then begin
          rsolver.AddCellRef(cell);
          try
            solver:=TCellFormula.Create;
            try
              solver.Data:=Data^;
              solver.Grid:=rsolver.Grid;
              if 0<>solver.DoGenerateRPNQueue(retv) then
                retv:=''
                else
                  if 0<>solver.SolveRPNQueue(retv) then
                    retv:='';
            finally
              solver.Free;
              rsolver.DelCellRef(cell);
            end;
          except
            retv:='';
          end;
        end else
          retv:='';
      end;
    end else
      retv:='';
  if retv='' then
    if MustGet then
      retv:='0'
        else
          retv:=' ';
  Result:=retv;
end;

// get one cell - user defined function
function GetCellValue(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer;
var
  aCol,aRow:Integer;
begin
  Result:=0;
  if Data<>nil then begin
      aRow:=StrToInt(R.Text);
      aCol:=aRow shr 16;
      aRow:=aRow and $FFFF;
      R.Text:=_ReadCell(aRow, aCol, Data, False);
  end;
end;

// return cell value ending pos
function _CheckCellPos(const cellstr: string;startpos,len:Integer): Integer;
var
  alpos,NumPos:Integer;
begin
  Result:=startpos;
  alpos:=0;
  NumPos:=0;
  while Result<=len do begin
    if cellstr[Result] in ['A'..'Z','a'..'z'] then begin
      if NumPos<>0 then begin
        Result:=-1;
        break;
      end;
      if alpos=0 then
        alpos:=Result;
    end else if cellstr[Result] in ['0'..'9'] then begin
      if alpos=0 then begin
        Result:=-1;
        break;
      end;
      if NumPos=0 then
        NumPos:=Result;
    end else begin
      break;
    end;
    Inc(Result);
  end;
  if Result<>-1 then
    if (alpos>=NumPos) or (alpos=0) or (Result-startpos<=1) or (NumPos-alpos>3) then
      Result:=-1;
end;

// ---------------------- funciton definition --------------------------

function func_Sum(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer;
var
  SCell:TSimpleTokenVar;
  x,y,top,left,bom,right:Integer;
begin
  Result:=0;
  try
    sCell:=TSimpleTokenVar(Stk.Pop);
    try
      bom:=StrToInt(R.Text);
      right:=bom shr 16;
      bom:=bom and $FFFF;
      top:=StrToInt(SCell.Text);
      left:=top shr 16;
      top:=top and $FFFF;
      _CheckRange(left,right);
      _CheckRange(top,bom);
      R.Text:='0';
      for y:=top to bom do
        for x:=left to right do begin
           SCell.tokenType:=toNone;
           SCell.Text:=_ReadCell(y,x,Data,False);
           if SCell.tokenType=toNumeric then
             mpr_add(R.tokenValue,SCell.tokenValue,R.tokenValue);
        end;
    finally
      SCell.Free;
    end;
  except
    R.tokenType:=toNone;
  end;
end;

function func_Avg(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer;
var
  SCell:TSimpleTokenVar;
  x,y,top,left,bom,right,z:Integer;
begin
  Result:=0;
  try
    sCell:=TSimpleTokenVar(Stk.Pop);
    try
      bom:=StrToInt(R.Text);
      right:=bom shr 16;
      bom:=bom and $FFFF;
      top:=StrToInt(SCell.Text);
      left:=top shr 16;
      top:=top and $FFFF;
      _CheckRange(left,right);
      _CheckRange(top,bom);
      R.Text:='0';
      z:=0;
      for y:=top to bom do
        for x:=left to right do begin
           SCell.tokenType:=toNone;
           SCell.Text:=_ReadCell(y,x,Data,False);
           if SCell.tokenType=toNumeric then begin
             Inc(z);
             mpr_add(R.tokenValue,SCell.tokenValue,R.tokenValue);
           end;
        end;
        if z>0 then begin
          SCell.Text:=IntToStr(z);
          if not mpr_is0(SCell.tokenValue) then
            mpr_div(R.tokenValue,SCell.tokenValue,R.tokenValue);
        end;
    finally
      SCell.Free;
    end;
  except
    R.tokenType:=toNone;
  end;
end;

function func_Min(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer;
var
  SCell:TSimpleTokenVar;
  x,y,top,left,bom,right:Integer;
begin
  Result:=0;
  try
    sCell:=TSimpleTokenVar(Stk.Pop);
    try
      bom:=StrToInt(R.Text);
      right:=bom shr 16;
      bom:=bom and $FFFF;
      top:=StrToInt(SCell.Text);
      left:=top shr 16;
      top:=top and $FFFF;
      _CheckRange(left,right);
      _CheckRange(top,bom);
      R.tokenType:=toNone;
      for y:=top to bom do
        for x:=left to right do begin
           SCell.tokenType:=toNone;
           SCell.Text:=_ReadCell(y,x,Data,False);
           if SCell.tokenType=toNumeric then
             if R.tokenType=toNone then begin
               mpr_copy(SCell.tokenValue,R.tokenValue);
               R.tokenType:=toNumeric;
             end else
               if mpr_is_lt(SCell.tokenValue,R.tokenValue) then
                 mpr_copy(SCell.tokenValue,R.tokenValue);
        end;
    finally
      SCell.Free;
    end;
  except
    R.tokenType:=toNone;
  end;
end;

function func_Max(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer;
var
  SCell:TSimpleTokenVar;
  x,y,top,left,bom,right:Integer;
begin
  Result:=0;
  try
    sCell:=TSimpleTokenVar(Stk.Pop);
    try
      bom:=StrToInt(R.Text);
      right:=bom shr 16;
      bom:=bom and $FFFF;
      top:=StrToInt(SCell.Text);
      left:=top shr 16;
      top:=top and $FFFF;
      _CheckRange(left,right);
      _CheckRange(top,bom);
      R.tokenType:=toNone;
      for y:=top to bom do
        for x:=left to right do begin
           SCell.tokenType:=toNone;
           SCell.Text:=_ReadCell(y,x,Data,False);
           if SCell.tokenType=toNumeric then
             if R.tokenType=toNone then begin
               mpr_copy(SCell.tokenValue,R.tokenValue);
               R.tokenType:=toNumeric;
             end else
               if mpr_is_gt(SCell.tokenValue,R.tokenValue) then
                 mpr_copy(SCell.tokenValue,R.tokenValue);
        end;
    finally
      SCell.Free;
    end;
  except
    R.tokenType:=toNone;
  end;
end;

function func_Count(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer;
var
  SCell:TSimpleTokenVar;
  x,y,top,left,bom,right:Integer;
begin
  Result:=0;
  try
    sCell:=TSimpleTokenVar(Stk.Pop);
    try
      bom:=StrToInt(R.Text);
      right:=bom shr 16;
      bom:=bom and $FFFF;
      top:=StrToInt(SCell.Text);
      left:=top shr 16;
      top:=top and $FFFF;
      _CheckRange(left,right);
      _CheckRange(top,bom);
      R.Text:='0';
      for y:=top to bom do
        for x:=left to right do begin
           SCell.tokenType:=toNone;
           SCell.Text:=_ReadCell(y,x,Data,False);
           if SCell.tokenType=toNumeric then begin
             mpr_set_int(SCell.tokenValue,1,1);
             mpr_add(R.tokenValue,SCell.tokenValue,R.tokenValue);
           end;
        end;
    finally
      SCell.Free;
    end;
  except
    R.tokenType:=toNone;
  end;
end;

function func_Min2(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  b:TSimpleTokenVar;
begin
  if Stk.Count>0 then begin
    b := TSimpleTokenVar(Stk.Pop);
    try
      if mpr_is_lt(b.tokenValue,R.tokenValue) then
        mpr_copy(b.tokenValue,R.tokenValue);
    finally
      b.Free;
    end;
  end else
    R.tokenType:=toNone;
  Result:=0;
end;

function func_Max2(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  b:TSimpleTokenVar;
begin
  if Stk.Count>0 then begin
    b := TSimpleTokenVar(Stk.Pop);
    try
      if mpr_is_gt(b.tokenValue,R.tokenValue) then
        mpr_copy(b.tokenValue,R.tokenValue);
    finally
      b.Free;
    end;
  end else
    R.tokenType:=toNone;
  Result:=0;
end;

function func_sin(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=sin(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_sinh(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=sinh(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_cos(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=cos(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_cosh(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=cosh(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;


function func_tan(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=tan(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_tanh(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=tanh(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_ln(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=ln(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_sqrt(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=sqrt(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_arctan(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=arctan(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_arctanh(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=arctanh(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_radtodeg(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=radtodeg(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_degtorad(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=degtorad(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_cotan(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=cotan(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_log10(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=log10(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_log2(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=log2(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_logn(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  b:TSimpleTokenVar;
  c1,c2:Double;
begin
  if Stk.Count>0 then begin
    b := TSimpleTokenVar(Stk.Pop);
    try
      c1:=mpr_todouble(R.tokenValue); // log c1 (c2)
      c2:=mpr_todouble(b.tokenValue);
      c1:=logn(c1,c2);
      mpr_read_double(R.tokenValue,c1);
    finally
      b.Free;
    end;
  end else
    R.tokenType:=toNone;
  Result:=0;
end;

function func_lnxp1(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=lnxp1(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_arcsin(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=arcsin(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_arcsinh(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=arcsinh(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_arccos(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=arccos(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

function func_arccosh(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=arccosh(dbl);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;

// ifequ(avalue,cvalue,equval,notval)
function func_ifequ(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  eqval,cval,aval:TSimpleTokenVar;
begin
  if Stk.Count>2 then begin
    Result:=0;
    eqval:=TSimpleTokenVar(Stk.Pop);
    if eqval.tokenType<>toNumeric then
      Result:=-9;
    cval:=TSimpleTokenVar(Stk.Pop);
    if cval.tokenType<>toNumeric then
      Result:=-9;
    aval:=TSimpleTokenVar(Stk.Pop);
    if aval.tokenType<>toNumeric then
      Result:=-9;
    if Result=0 then
      if mpr_is_eq(aval.tokenValue,cval.tokenValue) then
        mpr_copy(eqval.tokenValue,R.tokenValue);
  end else
    Result:=-9;
end;

// (avalue,cvalue,equval,notval)
function func_ifgt(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  eqval,cval,aval:TSimpleTokenVar;
begin
  if Stk.Count>2 then begin
    Result:=0;
    eqval:=TSimpleTokenVar(Stk.Pop);
    if eqval.tokenType<>toNumeric then
      Result:=-9;
    cval:=TSimpleTokenVar(Stk.Pop);
    if cval.tokenType<>toNumeric then
      Result:=-9;
    aval:=TSimpleTokenVar(Stk.Pop);
    if aval.tokenType<>toNumeric then
      Result:=-9;
    if Result=0 then
      if mpr_is_gt(aval.tokenValue,cval.tokenValue) then
        mpr_copy(eqval.tokenValue,R.tokenValue);
  end else
    Result:=-9;
end;

// (avalue,cvalue,equval,notval)
function func_iflt(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  eqval,cval,aval:TSimpleTokenVar;
begin
  if Stk.Count>2 then begin
    Result:=0;
    eqval:=TSimpleTokenVar(Stk.Pop);
    if eqval.tokenType<>toNumeric then
      Result:=-9;
    cval:=TSimpleTokenVar(Stk.Pop);
    if cval.tokenType<>toNumeric then
      Result:=-9;
    aval:=TSimpleTokenVar(Stk.Pop);
    if aval.tokenType<>toNumeric then
      Result:=-9;
    if Result=0 then
      if mpr_is_lt(aval.tokenValue,cval.tokenValue) then
        mpr_copy(eqval.tokenValue,R.tokenValue);
  end else
    Result:=-9;
end;

// (avalue,cvalue,equval,notval)
function func_ifgte(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  eqval,cval,aval:TSimpleTokenVar;
begin
  if Stk.Count>2 then begin
    Result:=0;
    eqval:=TSimpleTokenVar(Stk.Pop);
    if eqval.tokenType<>toNumeric then
      Result:=-9;
    cval:=TSimpleTokenVar(Stk.Pop);
    if cval.tokenType<>toNumeric then
      Result:=-9;
    aval:=TSimpleTokenVar(Stk.Pop);
    if aval.tokenType<>toNumeric then
      Result:=-9;
    if Result=0 then
      if mpr_is_ge(aval.tokenValue,cval.tokenValue) then
        mpr_copy(eqval.tokenValue,R.tokenValue);
  end else
    Result:=-9;
end;

// (avalue,cvalue,equval,notval)
function func_iflte(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  eqval,cval,aval:TSimpleTokenVar;
begin
  if Stk.Count>2 then begin
    Result:=0;
    eqval:=TSimpleTokenVar(Stk.Pop);
    if eqval.tokenType<>toNumeric then
      Result:=-9;
    cval:=TSimpleTokenVar(Stk.Pop);
    if cval.tokenType<>toNumeric then
      Result:=-9;
    aval:=TSimpleTokenVar(Stk.Pop);
    if aval.tokenType<>toNumeric then
      Result:=-9;
    if Result=0 then
      if mpr_is_le(aval.tokenValue,cval.tokenValue) then
        mpr_copy(eqval.tokenValue,R.tokenValue);
  end else
    Result:=-9;
end;

function func_RoundBank(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  dbl:Double;
begin
  dbl:=mpr_todouble(R.tokenValue);
  dbl:=RoundTo(dbl,-3);
  mpr_read_double(R.tokenValue,dbl);
  Result:=0;
end;


//-------------------------------------------------------------


{ TCellFormula }

function TCellFormula.AtToken(Token: TSimpleTokenVar; const str: string;
  CPos: Integer; var NextPos: Integer): boolean;
var
  i,j,l:Integer;
  tval:TSimpleTokenVar;
  IsRgn:Boolean;
begin
  Result:=False;
  if Token.tokenType=toNone then begin
    l:=Length(str);
    i:=_CheckCellPos(str,CPos,l);
    if i<>-1 then begin
      IsRgn:=False;
      j:=i;
      // one cell
      i:=AlphaToColRow(UpperCase(Copy(str,CPos,i-CPos)));
      tval:=TSimpleTokenVar.Create;
      tval.Text:=IntToStr(i);
      tval.tokenType:=toNumeric;
      RPNQueue.Push(tval);
      NextPos:=j;
      // cell range
      if j<l then
        if str[j]=':' then begin
          Inc(j);
          if j<l then begin
            i:=_CheckCellPos(str,j,l);
            if i<>-1 then begin
              IsRgn:=True;
              NextPos:=i;
              i:=AlphaToColRow(UpperCase(Copy(str,j,i-j)));
              Token.Text:=IntToStr(i);
            end;
          end;
        end;
      { push functions on operator stack,
        skip default to read next token }
      if not IsRgn then begin
        Token.Text:='GetCellValue';
        OperatorStk.Push(Token);
        Result:=True;
      end;
    end;
  end;
end;

constructor TCellFormula.Create;
begin
  inherited Create;
  OnToken:=@AtToken;
  FRefList:=TFPList.Create;
end;

destructor TCellFormula.Destroy;
begin
  FRefList.Free;
  inherited Destroy;
end;

procedure TCellFormula.AddCellRef(cell: Pointer);
begin
  FRefList.Add(cell);
end;

procedure TCellFormula.DelCellRef(cell: Pointer);
var
  i:Integer;
begin
  i:=FRefList.IndexOf(cell);
  if i<>-1 then
    FRefList.Delete(i);
end;

function TCellFormula.CheckRef(cell: Pointer): Integer;
begin
  Result:=FRefList.IndexOf(cell);
end;

procedure TCellFormula.ClearCellRef;
begin
  FRefList.Clear;
end;

function TCellFormula.CheckNumber(const s: string): Boolean;
begin
  Result:=IsNumericSigned(s);
end;


initialization
  Simple_AddFunction(-1,'GetCellValue',1,@GetCellValue);
  Simple_AddFunction(-1,'Sum',1,@func_Sum);
  Simple_AddFunction(-1,'Average',1,@func_Avg);
  Simple_AddFunction(-1,'Min',1,@func_Min);
  Simple_AddFunction(-1,'Max',1,@func_Max);
  Simple_AddFunction(-1,'Count',1,@func_Count);
  Simple_AddFunction(-1,'Min2',2,@func_Min2);
  Simple_AddFunction(-1,'Max2',2,@func_Max2);
  // math functions
  Simple_AddFunction(-1,'degtorad',1,@func_degtorad);
  Simple_AddFunction(-1,'radtodeg',1,@func_radtodeg);
  Simple_AddFunction(-1,'ln',1,@func_ln);
  Simple_AddFunction(-1,'lnxp1',1,@func_lnxp1);
  Simple_AddFunction(-1,'log10',1,@func_log10);
  Simple_AddFunction(-1,'log2',1,@func_log2);
  Simple_AddFunction(-1,'logn',2,@func_logn);
  Simple_AddFunction(-1,'sqrt',1,@func_sqrt);
  Simple_AddFunction(-1,'sin',1,@func_sin);
  Simple_AddFunction(-1,'sinh',1,@func_sinh);
  Simple_AddFunction(-1,'cos',1,@func_cos);
  Simple_AddFunction(-1,'cosh',1,@func_cosh);
  Simple_AddFunction(-1,'cotan',1,@func_cotan);
  Simple_AddFunction(-1,'tan',1,@func_tan);
  Simple_AddFunction(-1,'tanh',1,@func_tanh);
  Simple_AddFunction(-1,'arctan',1,@func_arctan);
  Simple_AddFunction(-1,'arcsin',1,@func_arcsin);
  Simple_AddFunction(-1,'arccos',1,@func_arccos);
  Simple_AddFunction(-1,'arctanh',1,@func_arctanh);
  Simple_AddFunction(-1,'arcsinh',1,@func_arcsinh);
  Simple_AddFunction(-1,'arccosh',1,@func_arccosh);
  // equal, great, less
  Simple_AddFunction(-1,'ifequ',4,@func_ifequ);
  Simple_AddFunction(-1,'ifgt',4,@func_ifgt);
  Simple_AddFunction(-1,'iflt',4,@func_iflt);
  Simple_AddFunction(-1,'ifgte',4,@func_ifgte);
  Simple_AddFunction(-1,'iflte',4,@func_iflte);
  //
  Simple_AddFunction(-1,'roundb',1,@func_RoundBank);


end.

