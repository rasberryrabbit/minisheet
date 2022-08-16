unit uSimplefmParser;

interface

// mini rpn parser for mpr by Do-wan Kim
//
// http://en.wikipedia.org/wiki/Shunting-yard_algorithm
//
// ----------------------------------------------------
// 1. DoGenerateRPNQueue
// 2. SolveRPNQueue = 0, success
// Result in text


{.$define _THREADSAFEMODE}

{$define SIMPLEPARSER_USE_THREAD}

uses classes, contnrs, SysUtils
{$ifdef _THREADSAFEMODE} ,syncobjs {$endif}
  , mp_types, mp_base, mp_ratio;

const
    _Simplefunc_NameLen = 31;

type
  ESimpleParseException = class(Exception);

  TEnhObjQueue = class(TObjectQueue)
  published
    property List;
  end;

  // Main Class

  TSimplefmData = record
    Data:Pointer;       // pointer to sheet
    DataFunc:Pointer;   // used in cell function
  end;
  PSimplefmData = ^TSimplefmData;

  TSimpleTokenVar=class;

  { TSimplefmParser }

  TOnSimplefmParserToken= function(Token:TSimpleTokenVar;const str:string;Start:Integer;var NextPos:Integer):boolean of object;

  TSimplefmParser = class
    private
      FErrcode:Integer;
      FSmartZero : Boolean;
      FPrecision : Integer;
      FResultStack : TObjectStack;
      FOperatorStack : TObjectStack;
      FRPNQueue : TEnhObjQueue;
      FOnToken:TOnSimplefmParserToken;
      procedure ClearRPNQueue;
      procedure ClearOperatorStack;
      procedure ClearResultStack;
      function GetResultText:string;
    protected
      {$ifdef SIMPLEPARSER_USE_THREAD}
      function CalcExpNum(var c:TSimpleTokenVar;expint:Integer):Boolean;
      {$endif}
    public
      Data:TSimplefmData;
      Timeout:Integer;

      constructor Create;
      destructor Destroy; override;

      // Generate RPN
      function DoGenerateRPNQueue(const str:string):Integer;
      // dump RPN Queue
      function DumpRPNQueue:string;
      // solve RPN Queue
      function SolveRPNQueue(out Rep: string): Integer;
      // paise error
      procedure RaiseError;

    property Precision: Integer read FPrecision write FPrecision; // Max Precision
    property SmartZero: Boolean read FSmartZero write FSmartZero; // remove useless '0'
    property Result: string read GetResultText;
    property OperatorStk:TObjectStack read FOperatorStack;
    property RPNQueue:TEnhObjQueue read FRPNQueue;
    property OnToken:TOnSimplefmParserToken read FOnToken write FOnToken;
  end;

  // Operator
  TSimpleOperatorAssociate = ( OpAssoNone, OpAssoLeftRight, OpAssoRightLeft );
  // token
  TSimpleTokenType = ( toNone, toNumeric, toOperator, toFunc, toPar );
  // token class

  { TSimpleTokenVar }

  TSimpleTokenVar = class
    private
      FSmartZero : Boolean;
      FPrecision : Integer;
      function ConvertToString : string;
      procedure StrToValue(const Value : string);
      function CheckOperator:Boolean;
      function CheckNumber:Boolean;
      function CheckParenthesis: Boolean;
    protected
      function IsNumberValue(const str:string;exp:PInteger):Boolean; virtual;
      function IsOperatorValue(const str:string):Boolean; virtual;
      function IsParenthesisValue(const str:string):Boolean; virtual;
    public
      tokenType  : TSimpleTokenType;
      tokenValue : mp_rat; // rational

      OperatorAsso : TSimpleOperatorAssociate;
      OperatorPrec : Integer;
      OperatorName : array[0.._Simplefunc_NameLen] of Ansichar;
      FuncId, FuncParam : Integer;
//      InternalFunc : Boolean;

      constructor Create;
      destructor Destroy; override;

    property Text: string read ConvertToString write StrToValue;
    property IsOperator: Boolean read CheckOperator;
    property IsNumber: Boolean read CheckNumber;
    property IsParenthesis: Boolean read CheckParenthesis;
    property Precision: Integer read FPrecision write FPrecision;
    property SmartZero: Boolean read FSmartZero write FSmartZero;
  end;

  { function definition
    R:Result,
    Stk:Result stack,
    Data.Data:User Data }
  TSimpleFuncProc = function (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer;
  { check basic Numeric value }
  function IsNumericSigned(const str:string;expmark:PInteger=nil):Boolean;
  { add user function
    Id=-1 , Auto Generate ID
    Name: unique name of function
  }
  function Simple_AddFunction(Id:Integer;const Name:string;ParamLen:Integer;Func:TSimpleFuncProc):Integer;

  function GetTokenFromStr(const str:string; out Token:string; Index:Integer):Integer;
{ function parameter separator,
  DecimalSeparator also used in functions
  Your localization DecimalSeparator is ',', you changes it to another char.
}
var
  _Parser_ParamSeparator:char=',';

implementation

uses BTypes;

resourcestring
  rsSyntaxError = 'Syntax Error';
  rsModularOpera = 'modular operator only support integer';
  rsNeedMoreFunc = 'need more function parameter(s)';
  rsFunctionIsNo = 'function is not found';
  rsRootOnlySupp = 'Root only support integer';
  rsTooManyFunct = 'too many function parameter(s)';
  rsSignDigitMus = 'sign digit must be "+" or "-"';
  rsCannotInsert = 'Cannot insert RPN queue';
  rsCalcurationE = 'Calcuration Error';

const
  _Parenthesis = ['(',')','{','}','[',']'];
  _ParenthesisLeft = ['(','{','['];

var
   ValueList : TList;
   {$ifdef _THREADSAFEMODE}
   _CritSection : TCriticalSection;
   {$endif}
   FuncTree: TFPHashList;

type
  {$ifdef SIMPLEPARSER_USE_THREAD}
  { TSimpleFMThread }

  TSimpleFMThread=class(TThread)
    private
      DoneFlag:Boolean;
    protected
      procedure Execute; override;
    public
      EventWait:PRTLEvent;
      // exponent
      expint:Integer;
      Val:TSimpleTokenVar;

      constructor Create(CreateSuspended: Boolean; const StackSize: SizeUInt=
        DefaultStackSize);
      destructor Destroy; override;
      function WaitDone(TimeOut:Integer):Boolean;
  end;
  {$endif}

  // define operators
  TSimpleOperatorProperty = record
    OpStr : array[0..4] of AnsiChar;
    Assoc : TSimpleOperatorAssociate;
    Prece : Integer;
  end;
  // define function
  TFunctionDef = record
    Id : Integer;
    Name : array[0.._Simplefunc_NameLen] of AnsiChar;
    Param : Integer;
    Func : TSimpleFuncProc;
  end;
  PFunctionDef=^TFunctionDef;

const
  // operator property
  OpPropArray : array[0..5] of TSimpleOperatorProperty = (
    ( OpStr:'+'; Assoc:OpAssoLeftRight; Prece:2; ),
    ( OpStr:'-'; Assoc:OpAssoLeftRight; Prece:2; ),
    ( OpStr:'*'; Assoc:OpAssoLeftRight; Prece:3; ),
    ( OpStr:'/'; Assoc:OpAssoLeftRight; Prece:3; ),
    ( OpStr:'%'; Assoc:OpAssoLeftRight; Prece:3; ),
    ( OpStr:'^'; Assoc:OpAssoRightLeft; Prece:4; )
  );

  // function definitions, 1-based
  FID_ABS = 1;
  FID_NEG = 2;
  FID_RECIP = 3;
  FID_FRAC = 4;
  FID_TRUNC = 5;
  FID_ROUND = 6;
  FID_FLOOR = 7;
  FID_CEIL = 8;
  FID_ROOT = 9;
  FID_PI = 10;
  FID_ADD = 11;

  { ===== default functions ====== }
    function f_ABS  (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_NEG  (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_RECIP(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_FRAC (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_TRUNC(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_ROUND(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_FLOOR(var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_CEIL (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_ROOT (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_PI   (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
    function f_ADD  (var R:TSimpleTokenVar;Stk:TObjectStack;Data:PSimplefmData):Integer; forward;
  { ===================== }


  // default functions record
const
  _DefaultFuncCount = 11;
  _FuncList : array[0.._DefaultFuncCount-1] of TFunctionDef = (
   ( Id: FID_ABS; Name: 'ABS'; Param: 1; Func: @f_ABS ),      // |n|
   ( Id: FID_NEG; Name: 'NEG'; Param: 1; Func: @f_NEG ),      // -n
   ( Id: FID_RECIP; Name: 'RECIP'; Param: 1; Func: @f_RECIP ),  // 1/n
   ( Id: FID_FRAC; Name:  'FRAC'; param: 1; Func: @f_FRAC ),   // Frac
   ( Id: FID_TRUNC; Name:  'TRUNC'; param: 1; Func: @f_TRUNC ), // trunc
   ( Id: FID_ROUND; Name:  'ROUND'; param: 1; Func: @f_ROUND ), // round
   ( Id: FID_FLOOR; Name:  'FLOOR'; param: 1; Func: @f_FLOOR ), // floor
   ( Id: FID_CEIL; Name:  'CEIL'; param: 1; Func: @f_CEIL ),   // ceil
   ( Id: FID_ROOT; Name:  'ROOT'; param: 1; Func: @f_ROOT ),   // root, integer
   ( Id: FID_PI; Name:  'PI'; param: 0; Func: @f_PI ),       // PI, test
   ( Id: FID_ADD; Name:  'ADD'; param: 2; Func: @f_ADD )      // add, test function
  );

  // precision
  _DefaultPrecision = 40;
  // Pi
  PI_STR = '3.14159265358979323846264338327950288419716939937510';

  // add function
  function Simple_AddFunction(Id:Integer;const Name:string;ParamLen:Integer;Func:TSimpleFuncProc):Integer;
  var
    newfunc:PFunctionDef;
    newfname:string;
  begin
    Result:=-1;
    newfunc:=nil;
    newfname:=UpperCase(Copy(Name,1,_Simplefunc_NameLen));
    if FuncTree.Find(newfname)=nil then begin
      GetMem(newfunc,sizeof(TFunctionDef));
      try
        if id=-1 then
          id:=FuncTree.Count+1
          else
            newfunc^.Id:=Id;
        strlcopy(newfunc^.Name,pchar(newfname),_Simplefunc_NameLen);
        newfunc^.Param:=ParamLen;
        newfunc^.Func:=Func;
        Result:=FuncTree.Add(newfname,newfunc)+2;
      except
        Freemem(newfunc);
      end;
    end;
  end;

  // initialize default functions
  procedure InitFuncList;
  var
    i:Integer;
  begin
    for i:=0 to _DefaultFuncCount-1 do
      Simple_AddFunction(_FuncList[i].Id,_FuncList[i].Name,_FuncList[i].Param,_FuncList[i].Func);
  end;

  // operator property
  function GetOperatorProp(const opstr:string;out OpAsso:TSimpleOperatorAssociate;
      out OpPrec:Integer):Boolean;
  var
    i : Integer;
  begin
    Result := False;
    for i:=Low(OpPropArray) to High(OpPropArray) do begin
      if opstr=OpPropArray[i].OpStr then begin
        OpAsso := OpPropArray[i].Assoc;
        OpPrec := OpPropArray[i].Prece;
        Result := True;
        break;
      end;
    end;
  end;

  // get function id from token
  function _GetFuncID(const Name:string;var Param:Integer):Integer;
  var
    func:PFunctionDef;
  begin
    func:=FuncTree.Find(UpperCase(Name));
    if func<>nil then begin
      Result:=func^.Id;
      Param:=func^.Param;
      end else Result:=0;
  end;

  // alloc numeric value
  procedure _TokenAlloc(var Value : mp_rat);
  begin
    mpr_init(Value);
    {$ifdef _THREADSAFEMODE}
    _CritSection.Acquire;
    try
    {$endif}
    ValueList.Add(@Value);
    {$ifdef _THREADSAFEMODE}
    finally
    _CritSection.Release;
    end;
    {$endif}
  end;
  // free numeric value
  procedure _TokenFreeValue(var Value : mp_rat);
  begin
    {$ifdef _THREADSAFEMODE}
    _CritSection.Acquire;
    try
    {$endif}
    ValueList.Extract(@Value);
    {$ifdef _THREADSAFEMODE}
    finally
    _CritSection.Release;
    end;
    {$endif}
    mpr_clear(Value);
  end;
  // copy value
  procedure _TokenCopyValue(var New: mp_rat; const Old : mp_rat);
  begin
    mpr_copy(Old,New);
  end;
  // convert to value
  procedure _StrToMPRAT(const str:string; var value : mp_rat);
  begin
    mpr_read_float_decimal(value,pchar8(str));
  end;

  // string functions, Index is 1-based;
  // get token from string, return next token pos
  function GetTokenFromStr(const str:string; out Token:string; Index:Integer):Integer;
  var
    Len : Integer;
    ch : AnsiChar;
    TokenType : Integer;
    FindToken, expos : Boolean;
  begin
    Len := Length(str);
    Result := Index;
    TokenType := 0;
    expos:=False;
    FindToken := False;
    Token := '';

    while Result<=Len do begin
      ch := str[Result];
      // number, function
      if (ch in ['0'..'9','A'..'Z','a'..'z','_']) or (ch=DecimalSeparator) then begin
        if (TokenType<>3) and FindToken then
          Break
          else begin
            TokenType := 3;
            FindToken := True;
          end;
        // is exponent number?
        if (not expos) and (ch in ['e','E']) then begin
          expos:=true;
          if IsNumericSigned(Token) then
            FindToken:=False;
        end;
        Token := Token + ch;
      end else
        // space
        if ch <= #32 then begin
          if FindToken then
            Break;
          // skip space
          while Result<=Len do begin
            if str[Result] > #32 then begin
              Dec(Result);
              Break;
            end;
            Inc(Result);
          end;
        end else
          // operator or else
          begin
            // make exponent number
            if (not FindToken) and (TokenType=3) and (ch in ['+','-']) then
            begin
              Token:=Token+ch;
              Inc(Result);
              FindToken:=True;
              Continue;
            end else
            if (TokenType<>4) and FindToken then
              Break
              else begin
                TokenType := 4;
                FindToken := True;
              end;
            // operator
            if ch in ['+','-','*','/','%','^','='] then
              TokenType := 5;
            // Parenthesis
            if (ch in _Parenthesis) or (ch=_Parser_ParamSeparator) then
              TokenType := 6;
            Token := Token + ch;
          end;
      Inc(Result);
    end;
    //
    if not FindToken then
      Result := 0;
  end;

  // check basic number format
  function IsNumericSigned(const str:string;expmark:PInteger=nil):Boolean;
  var
    i, Len, signpos, dotpos, numpos, expos : Integer;
  begin
    Result := True;
    Len := Length(str);
    i := 1;
    signpos:=0;
    dotpos:=0;
    numpos:=0;
    expos:=0;
    if expmark<>nil then
      expmark^:=0;
    while i<=Len do begin
      if str[i] in ['0'..'9'] then begin // digit
        if numpos=0 then
          numpos:=i;
      end else
      if str[i] in ['+','-'] then begin // signed digit, only one
        if (signpos=0) and (numpos=0) then
          signpos:=i
          else begin
            Result:=False;
            break;
          end;
      end else if str[i]=DecimalSeparator then begin // dot, only one
        if (dotpos=0) and (expos=0) then
          dotpos:=i
          else begin
            Result:=False;
            break;
          end;
      end else if str[i] in ['e','E'] then begin // exponent
        if (numpos=0) or (expos<>0) then begin
          Result:=False;
          break;
        end;
        if (signpos>numpos) or
           ((dotpos<>0) and (signpos>dotpos)) then begin
          Result:=False;
          break;
        end;
        numpos:=0;
        signpos:=0;
        expos:=i;
        if expmark<>nil then
          expmark^:=i;
      end else begin // non-digit and signed
        Result:=False;
        break;
      end;
      Inc(i);
    end;
    if Result then
      if (numpos=0) or
         (numpos<signpos) or
         ((expos=0) and (dotpos<>0) and (signpos>dotpos))then
        Result:=False;
  end;

{ TSimpleFMThread }

  procedure mp_expt_int_b(const a: mp_int; b: longint; var c: mp_int;skip:PBoolean);
    {-Calculate c = a^b, b>=0}
  var
    x: mp_int;
  begin
    if mp_error<>MP_OKAY then exit;

    if b<0 then begin
          raise MPXRange.Create('mp_expt_int: b < 0');
    end;

    {easy outs}
    if b=0 then begin
      mp_set1(c);
      exit;
    end
    else if mp_iszero(a) then begin
      mp_zero(c);
      exit;
    end
    else if b=1 then begin
      mp_copy(a,c);
      exit;
    end
    else if b=2 then begin
      mp_sqr(a,c);
      exit;
    end;

    {make local copy of a}
    mp_init_copy(x,a);
    if mp_error<>MP_OKAY then exit;

    mp_set1(c);
    {b >2, therefore loop is executed at least once}
    while (mp_error=MP_OKAY) and (not skip^) do begin
      if odd(b) then mp_mul(c,x,c);
      b := b shr 1;
      if b=0 then break;
      mp_sqr(x,x);
    end;
    mp_clear(x);
  end;

  procedure mpr_expt_b(const a: mp_rat; b: longint; var c: mp_rat; skip:PBoolean);
    {-Calculate c = a^b, a<>0 for b<0, 0^0=1}
  begin
    if mp_error<>MP_OKAY then exit;

    {Handle a=0}
    if a.den.used=0 then begin
      if b<0 then begin
            raise MPXRange.Create('mpr_expt: a=0, b<0');
      end
      else if b>0 then mpr_zero(c)
      else mpr_set_int(c,1,1);
      exit;
    end;

    {Setup for exponentiation and handle easy cases}
    if b<0 then begin
      mpr_inv(a,c);
      b := abs(b);
    end
    else begin
      if b=0 then mpr_set_int(c,1,1) else mpr_copy(a,c);
    end;
    if b<=1 then exit;

    {here we have to calculate c^b, b>1. Since rats are always stored}
    {in lowest terms, calculate c.num^b/c.den^b without normalization}
    mp_expt_int_b(c.num, b, c.num, skip);
    mp_expt_int_b(c.den, b, c.den, skip);
  end;


procedure TSimpleFMThread.Execute;
begin
  DoneFlag:=False;
  RTLeventSetEvent(EventWait);
  try
    // exp = 131071
    if mp_is1(Val.tokenValue.den) then begin
      mp_expt_int_b(Val.tokenValue.num,expint and $1FFFF,Val.tokenValue.num,@Terminated);
    end
    else begin
      mpr_expt_b(Val.tokenValue,expint and $1FFFF,Val.tokenValue,@Terminated);
    end;
  except
  end;
  RTLeventResetEvent(EventWait);
  DoneFlag:=True;
end;

constructor TSimpleFMThread.Create(CreateSuspended: Boolean;
  const StackSize: SizeUInt);
begin
  inherited Create(CreateSuspended,StackSize);
  Priority:=tpHighest;
  FreeOnTerminate:=False;
  EventWait:=RTLEventCreate;
end;

destructor TSimpleFMThread.Destroy;
begin
  RTLeventdestroy(EventWait);
  inherited Destroy;
end;

function TSimpleFMThread.WaitDone(TimeOut: Integer): Boolean;
begin
  RTLeventWaitFor(EventWait,TimeOut);
  Result:=DoneFlag;
end;


{ TSimpleTokenVar }

function TSimpleTokenVar.CheckNumber: Boolean;
begin
  Result := tokenType = toNumeric;
end;

function TSimpleTokenVar.CheckOperator: Boolean;
begin
  Result := tokenType = toOperator;
end;

function TSimpleTokenVar.CheckParenthesis: Boolean;
begin
  Result := (OperatorName[0] in _Parenthesis) or (OperatorName[0]=_Parser_ParamSeparator);
end;

function TSimpleTokenVar.IsNumberValue(const str: string; exp: PInteger
  ): Boolean;
begin
  Result:=IsNumericSigned(str,exp);
end;

function TSimpleTokenVar.IsOperatorValue(const str: string): Boolean;
var
  i : Integer;
begin
  Result := False;
  for i:=Low(OpPropArray) to High(OpPropArray) do begin
    if str=OpPropArray[i].OpStr then begin
      Result := True;
      break;
    end;
  end;
end;

function TSimpleTokenVar.IsParenthesisValue(const str: string): Boolean;
begin
  Result:=(str<>'') and ((str[1] in _Parenthesis) or (str[1]=_Parser_ParamSeparator));
end;

function TSimpleTokenVar.ConvertToString: string;
var
  iLen : Integer;
  OutputStr : string;
begin
  case tokenType of
  toNumeric : begin
        // no precision
        if mp_is1(tokenValue.den) then
          OutputStr:=mpr_tofloat_astr(tokenValue,10,0)
        else begin
        // has precision
          OutputStr :=mpr_tofloat_astr(tokenValue,10,FPrecision);
          if FSmartZero and (FPrecision>0) then begin
            iLen := Length(OutputStr);
            while iLen>0 do begin
                  if OutputStr[iLen]<>'0' then begin
                    if OutputStr[iLen]<>DecimalSeparator then
                      Inc(iLen);
                    OutputStr:=Copy(OutputStr,1,iLen-1);
                    Break;
                  end;
                  Dec(iLen);
            end;
          end;
        end;
        Result := OutputStr;
    end;
  toOperator, toFunc : begin
      Result := OperatorName;
    end;
  else
    Result := '';
  end;
end;

constructor TSimpleTokenVar.Create;
begin
  inherited;
  tokenType := toNone;
  FPrecision := _DefaultPrecision;
  FSmartZero := True;
  FuncId:=0;
  FuncParam:=0;
//  InternalFunc:=False;
  _TokenAlloc(tokenValue);
end;

destructor TSimpleTokenVar.Destroy;
begin
  _TokenFreeValue(tokenValue);
  inherited;
end;

procedure TSimpleTokenVar.StrToValue(const Value: string);
var
  exp:Integer;
  rtemp:Double;
begin
   if IsNumberValue(Value,@exp) then begin
     if exp=0 then begin
       // decimal
       _StrToMPRAT(Value,tokenValue);
     end else begin
       // float
       rtemp:=StrToFloatDef(Value,0.0);
       mpr_read_double(tokenValue,rtemp);
     end;
     tokenType:=toNumeric;
   end else begin
     mpr_zero(tokenValue);
     StrLCopy(@(OperatorName[0]),PAnsichar(Value),_Simplefunc_NameLen);
     // convert token to tokenvar
     if IsOperatorValue(Value) then begin
       if GetOperatorProp(Value,OperatorAsso,OperatorPrec) then
       begin
         tokenType := toOperator;
       end;
    end else
    if IsParenthesisValue(Value) then begin
      tokenType:=toPar;
    end else begin
      // func and string
      FuncId:=_GetFuncID(Value, FuncParam);
      // function id is 1-based
      if FuncId<>0 then begin
        tokenType:=toFunc;
        {
        // internal func
        InternalFunc:=FuncParam<0;
        if InternalFunc then
          FuncParam:=-FuncParam;
        }
      end;
    end;
  end;
end;

{ TSimplefmParser }

procedure TSimplefmParser.ClearOperatorStack;
begin
  set_mp_error(MP_OKAY);
  while FOperatorStack.Count>0 do
    TSimpleTokenVar(FOperatorStack.Pop).Free;
end;

procedure TSimplefmParser.ClearResultStack;
begin
  set_mp_error(MP_OKAY);
  while FResultStack.Count>0 do
    TSimpleTokenVar(FResultStack.Pop).Free;
end;

function TSimplefmParser.GetResultText: string;
begin
  if FResultStack.Count=1 then
     Result:=TSimpleTokenVar(FResultStack.Peek).Text
     else Result:='';
end;

{$ifdef SIMPLEPARSER_USE_THREAD}
function TSimplefmParser.CalcExpNum(var c: TSimpleTokenVar; expint: Integer
  ): Boolean;
var
  calcThrd:TSimpleFMThread;
  i:Integer;
  FLastEvent:PRTLEvent;
  CalcSkip:Boolean;
begin
  Result:=True;
  CalcSkip:=False;
  calcThrd:=TSimpleFMThread.Create(True);
  calcThrd.expint:=expint;
  calcThrd.Val:=c;
  FLastEvent:=calcThrd.EventWait;
  calcThrd.Resume;
  i:=0;
  // delay about Timeout milisecond
  while not calcThrd.WaitDone(1) do begin
    Inc(i);
    if i>=Timeout then begin
      if not CalcSkip then begin
        CalcSkip:=True;
        i:=0;
        calcThrd.Terminate;
        end else begin
          // never happens
          Result:=False;
          RTLeventdestroy(FLastEvent);
          break;
        end;
    end;
  end;
  if Result then begin
    c:=calcThrd.Val;
    calcThrd.Free;
  end;
end;
{$endif}

procedure TSimplefmParser.ClearRPNQueue;
begin
  while FRPNQueue.Count>0 do
    TSimpleTokenVar(FRPNQueue.Pop).Free;
end;

constructor TSimplefmParser.Create;
begin
  FPrecision := _DefaultPrecision;
  FErrcode:=0;
  Timeout:=800; // freeze thread if too low, > 500
  FSmartZero := True;
  Data.DataFunc:=nil;
  Data.Data:=nil;
  //
  FResultStack := TObjectStack.Create;
  FOperatorStack := TObjectStack.Create;
  FRPNQueue := TEnhObjQueue.Create;
  mp_fract_sep:=DecimalSeparator;
  if DecimalSeparator=',' then
    _Parser_ParamSeparator:=';';
end;

destructor TSimplefmParser.Destroy;
begin
  inherited;
  //
  ClearRPNQueue;
  ClearOperatorStack;
  ClearResultStack;
  //
  FResultStack.Free;
  FOperatorStack.Free;
  FRPNQueue.Free;
end;

// Generate RPN Queue
function TSimplefmParser.DoGenerateRPNQueue(const str: string): Integer;
const MaxParamChkDeep=64;
type
  TParamChk = array[0..MaxParamChkDeep] of Integer;
  PParamChk = ^TParamChk;
  TOpenCloseChk = array[0..MaxParamChkDeep] of Integer;
  POpenCloseChk = ^TOpenCloseChk;
var
  i, Len, ParamChkDeep, OpenCloseCount, TokStart, ftokenpos: Integer;
  tokn : string;
  value, tempvar : TSimpleTokenVar;
  PrevOp : TSimpleTokenVar;
  prevtk : TSimpleTokenType;
  ParamChk: PParamChk;
  OpenCloseChk: POpenCloseChk;
  _tempvalue : char;
begin
  Result:=0;
  i := 1;
  prevtk:=toNone;
  ftokenpos:=0;
  ParamChkDeep:=0;
  Len := Length(str);
  FErrcode:=0;

  ClearOperatorStack;
  ClearRPNQueue;

  New(ParamChk);
  New(OpenCloseChk);
  try
    OpenCloseCount:=0;
    OpenCloseChk^[0]:=0;


    while i<>0 do begin
      i := GetTokenFromStr(str,tokn,i);
      if i=0 then
        break;
      // check formula starting
      if ftokenpos=0 then begin
        if tokn='=' then begin
          ftokenpos:=i;
          continue;
        end else begin
          FErrcode:=-1;
          exit(-1);
        end;
      end;
      value := TSimpleTokenVar.Create;
      try
        value.Precision := FPrecision;
        value.SmartZero := FSmartZero;
        value.Text := tokn;
      except
        value.tokenType := toNone;
      end;
      TokStart:=i-length(tokn);
      if Assigned(FOnToken) then
        if FOnToken(value,str,TokStart,i) then begin
          prevtk:=value.tokenType;
          Continue;
        end;
      case value.tokenType of
      toNumeric : begin
                  FRPNQueue.Push(value);
                  prevtk:=toNumeric;
                  end;
      toOperator : begin
                    if (prevtk<>toNumeric) and (prevtk<>toFunc) then begin
                      // sign digit
                      if value.OperatorName[0] in ['+','-'] then begin
                       tempvar:=TSimpleTokenVar.Create;
                       try
                         tempvar.Precision:=FPrecision;
                         tempvar.SmartZero:=FSmartZero;
                         tempvar.tokenType:=toNumeric;
                         mpr_zero(tempvar.tokenValue);
                         FRPNQueue.Push(tempvar);
                       except
                         tempvar.Free;
                         FErrcode:=-8;
                         exit(FErrcode);
                       end;
                      end else begin
                        FErrcode:=-7;
                        exit(FErrcode);
                      end;
                    end else begin
                      // operator
                      while FOperatorStack.Count>0 do begin
                        PrevOp := TSimpleTokenVar(FOperatorStack.Peek);
                        if ((value.OperatorAsso = OpAssoLeftRight) and
                          (value.OperatorPrec <= PrevOp.OperatorPrec)) or
                          ((value.OperatorAsso = OpAssoRightLeft) and
                          (value.OperatorPrec < PrevOp.OperatorPrec)) or
                          (PrevOp.tokenType = toFunc)
                          then
                          begin
                            PrevOp := TSimpleTokenVar(FOperatorStack.Pop);
                            FRPNQueue.Push(PrevOp);
                          end
                          else
                            break;
                      end;
                    end;
                    FOperatorStack.Push(value);
                    prevtk:=toOperator;
                   end;
      toFunc : begin
                if value.FuncParam>0 then begin
                  FOperatorStack.Push(value);
                  // check function parameter count
                  if (ParamChkDeep<MaxParamChkDeep) then begin
                    Inc(ParamChkDeep);
                    ParamChk^[ParamChkDeep]:=value.FuncParam;
                  end;
                end else begin
                  FRPNQueue.Push(value);
                end;
                prevtk:=toFunc;
               end;
      toPar :
               begin
                // Parenthesis
                if value.OperatorName[0] in _ParenthesisLeft then begin
                  // 'number(...' case
                  if prevtk=toNumeric then begin
                    tempvar:=TSimpleTokenVar.Create;
                    tempvar.Text:='*';
                    FOperatorStack.Push(tempvar);
                  end;
                  FOperatorStack.Push(value);
                  // Parenthesis level
                  if prevtk=toFunc then
                    OpenCloseChk^[ParamChkDeep]:=OpenCloseCount;
                  Inc(OpenCloseCount);
                  prevtk:=toNone;
                end else begin
                  while FOperatorStack.Count>0 do begin
                    PrevOp := TSimpleTokenVar(FOperatorStack.Pop);
                    if PrevOp.OperatorName[0] in _ParenthesisLeft then begin
                      PrevOp.Free;
                      if prevtk=toOperator then begin
                        FErrcode:=-1;
                        exit(FErrcode);
                      end;
                      // check function parameter count
                      if ParamChkDeep>0 then
                        if FOperatorStack.Count>0 then
                          if TSimpleTokenVar(FOperatorStack.Peek).tokenType=toFunc then
                            if ParamChk^[ParamChkDeep]<=0 then begin
                              FErrcode:=-6;
                              exit(FErrcode);
                            end;
                      break;
                    end else begin
                      FRPNQueue.Push(PrevOp);
                    end;
                  end;
                  // Parenthesis check
                  if OpenCloseCount>0 then
                     Dec(OpenCloseCount)
                     else
                     begin
                       FErrcode:=-1;
                       exit(FErrcode);
                     end;
                  prevtk:=toNone;
                  // ','='()' split parameters, function(arg1)(arg2)...
                  if value.OperatorName[0]=_Parser_ParamSeparator then begin
                     value.OperatorName[0]:='(';
                     FOperatorStack.Push(value);
                     Inc(OpenCloseCount);
                     // check parameters length
                     if ParamChkDeep>0 then begin
                       if ParamChk^[ParamChkDeep]>0 then
                         Dec(ParamChk^[ParamChkDeep]);
                       if ParamChk^[ParamChkDeep]=0 then begin
                         FErrcode:=-6;
                         exit(FErrcode);
                       end;
                     end;
                  end else begin
                    prevtk:=toNumeric;
                     _tempvalue:=value.OperatorName[0];
                     value.Free;
                     // check closing function's parameter
                     if ParamChkDeep>0 then
                        if OpenCloseChk^[ParamChkDeep]=OpenCloseCount then
                          if (ParamChk^[ParamChkDeep]=1)
                              and (_tempvalue=')') then begin
                                  Dec(ParamChkDeep);
                                end else begin
                                  FErrcode:=-3;
                                  exit(FErrcode);
                                end;
                  end;
                end;
               end;
      toNone : begin
                 // skip parsing
                 value.Free;
                 Result:=-1;
                 break;
               end;
      end;
    end;
    // check last token
    if prevtk<>toOperator then begin
     if OpenCloseCount<>0 then begin
        Result:=-1;
        exit(FErrcode);
      end;
     // store remained operators to RPN queue
      while FOperatorStack.Count>0 do begin
        PrevOp := TSimpleTokenVar(FOperatorStack.Pop);
        FRPNQueue.Push(PrevOp);
      end;
    end else begin
      FErrcode:=-1;
      exit(FErrcode);
    end;
  finally
    Dispose(ParamChk);
    Dispose(OpenCloseChk);
  end;
end;

function TSimplefmParser.DumpRPNQueue: string;
var
  value : TSimpleTokenVar;
  i : Integer;
begin
  Result:='';
  if FRPNQueue.Count>0 then begin
    i := FRPNQueue.Count - 1;
    while i>=0 do begin
      value := TSimpleTokenVar(FRPNQueue.List[i]);
      if Result <> '' then
        Result := Result + ' ';
      Result := Result + value.Text;
      Dec(i);
    end;
  end;
end;

{ ------ functions ---------------------------------- }
function f_ABS(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
begin
  mpr_abs(R.tokenValue,R.tokenValue);
  Result:=0;
end;

function f_NEG(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
begin
  mpr_chs(R.tokenValue,R.tokenValue);
  Result:=0;
end;

function f_RECIP(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
begin
  mpr_inv(R.tokenValue,R.tokenValue);
  Result:=0;
end;

function f_FRAC(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
begin
  mpr_frac(R.tokenValue,R.tokenValue);
  Result:=0;
end;

function f_TRUNC(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  tmpint:mp_int;
begin
  mp_init(tmpint);
  mpr_trunc(R.tokenValue,tmpint);
  mp_copy(tmpint,R.tokenValue.num);
  mp_set_int(R.tokenValue.den,1);
  mp_clear(tmpint);
  Result:=0;
end;

function f_ROUND(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  tmpint:mp_int;
begin
  mp_init(tmpint);
  mpr_round(R.tokenValue,tmpint);
  mp_copy(tmpint,R.tokenValue.num);
  mp_set_int(R.tokenValue.den,1);
  mp_clear(tmpint);
  Result:=0;
end;

function f_FLOOR(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  tmpint:mp_int;
begin
  mp_init(tmpint);
  mpr_floor(R.tokenValue,tmpint);
  mp_copy(tmpint,R.tokenValue.num);
  mp_set_int(R.tokenValue.den,1);
  mp_clear(tmpint);
  Result:=0;
end;

function f_CEIL(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  tmpint:mp_int;
begin
  mp_init(tmpint);
  mpr_ceil(R.tokenValue,tmpint);
  mp_copy(tmpint,R.tokenValue.num);
  mp_set_int(R.tokenValue.den,1);
  mp_clear(tmpint);
  Result:=0;
end;

function f_ROOT(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  tmpint:mp_int;
begin
  Result:=0;
  if mp_is1(R.tokenValue.den) then begin
    mp_init(tmpint);
    mp_n_root(R.tokenValue.num,2,tmpint);
    mp_copy(tmpint,R.tokenValue.num);
    mp_clear(tmpint);
  end else
    Result:=-5;
end;

function f_PI(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
begin
  R.Text:=PI_STR;
  Result:=0;
end;

function f_ADD(var R: TSimpleTokenVar; Stk: TObjectStack; Data: PSimplefmData
  ): Integer;
var
  b:TSimpleTokenVar;
begin
  Result:=0;
  if Stk.Count>0 then begin
    b := TSimpleTokenVar(Stk.Pop);
    try
    mpr_add(R.tokenValue,b.tokenValue,R.tokenValue);
    finally
      b.Free;
    end;
  end else
    Result:=-3;
end;
{ ------------------------------------------------------- }


// solve RPN Queue
function TSimplefmParser.SolveRPNQueue(out Rep:string): Integer;
var
  tmpval : TSimpleTokenVar;
  c , b : TSimpleTokenVar;
  expnum,fret : Integer;
  func : PFunctionDef;
begin
  Result := 0;
  FErrcode:=0;
  ClearResultStack;

  while FRPNQueue.Count>0 do begin
    tmpval := TSimpleTokenVar(FRPNQueue.Pop);
    case tmpval.tokenType of
    // number
    toNumeric : FResultStack.Push(tmpval);
    // function
    toFunc :    if FResultStack.Count>=tmpval.FuncParam then begin
                  if tmpval.FuncParam>0 then
                     c := TSimpleTokenVar(FResultStack.Pop)
                     else begin
                          c:= TSimpleTokenVar.Create;
                          c.tokenType:=toNumeric;
                          c.Precision:=FPrecision;
                          c.SmartZero:=FSmartZero;
                     end;
                  try
                    func:=FuncTree.Find(UpperCase(tmpval.OperatorName));
                    if func<>nil then begin
                      fret:=func^.Func(c,FResultStack,@Data);
                      if fret<>0 then begin
                        FErrcode:=fret;
                        exit(FErrcode);
                      end;
                    end else begin
                      FErrcode:=-4;
                      exit(FErrcode);
                    end;
                    FResultStack.Push(c);
                  except
                    c.Free;
                    raise;
                  end;
                  tmpval.Free;
                end else begin
                  FErrcode:=-3;
                  exit(FErrcode);
                end;
    // operator
    toOperator : if FResultStack.Count>0 then begin
                try
                  // op2
                  b := TSimpleTokenVar(FResultStack.Pop);
                  // op1
                  if FResultStack.Count>0 then
                    c := TSimpleTokenVar(FResultStack.Pop)
                    else begin
                      c := TSimpleTokenVar.Create;
                      c.Precision := FPrecision;
                      c.SmartZero := FSmartZero;
                      c.tokenType := toNumeric;
                      mpr_zero(c.tokenValue);
                    end;

                  try
                    // process operator
                    if tmpval.OperatorName='+' then
                      mpr_add(c.tokenValue,b.tokenValue,c.tokenValue)
                    else
                    if tmpval.OperatorName='-' then
                      mpr_sub(c.tokenValue,b.tokenValue,c.tokenValue)
                    else
                    if tmpval.OperatorName='*' then
                      mpr_mul(c.tokenValue,b.tokenValue,c.tokenValue)
                    else
                    if tmpval.OperatorName='/' then
                      mpr_div(c.tokenValue,b.tokenValue,c.tokenValue)
                    else
                    if tmpval.OperatorName='%' then begin
                      if mp_is1(c.tokenValue.den) and mp_is1(b.tokenValue.den) then
                         mp_mod(c.tokenValue.num,b.tokenValue.num,c.tokenValue.num)
                         else begin
                           FErrcode:=-2;
                           exit(FErrcode);
                         end;
                    end else
                    if tmpval.OperatorName='^' then begin
                      expnum:=mp_get_int(b.tokenValue.num);
                      if (expnum<0) or (expnum>$1FFFF) then begin
                        FErrcode:=-9;
                        exit(FErrcode);
                      end;
                      {$ifndef SIMPLEPARSER_USE_THREAD}
                      if mp_is1(c.tokenValue.den) then begin
                        mp_expt_int(c.tokenValue.num,expnum,c.tokenValue.num);
                      end
                      else begin
                        mpr_expt(c.tokenValue,expnum,c.tokenValue);
                      end;
                      {$else}
                      if not CalcExpNum(c,expnum) then begin
                        FErrcode:=-9;
                        exit(FErrcode);
                      end;
                      {$endif}
                    end;
                    // store to result stack
                    FResultStack.Push(c);
                  except
                    c.Free;
                    FErrcode:=-9;
                    exit(FErrcode);
                  end;
                finally
                  b.Free;
                end;
                tmpval.Free;
             end else begin
               FErrcode:=-1;
               exit(FErrcode);
             end;
    end;
  end;
  // result
  if FResultStack.Count=1 then
    Rep := TSimpleTokenVar(FResultStack.Peek).Text
    else
      Result := -9;
end;

procedure TSimplefmParser.RaiseError;
var
  errmsg:string;
begin
  case FErrcode of
  -1: errmsg:=rsSyntaxError;
  -2: errmsg:=rsModularOpera;
  -3: errmsg:=rsNeedMoreFunc;
  -4: errmsg:=rsFunctionIsNo;
  -5: errmsg:=rsRootOnlySupp;
  -6: errmsg:=rsTooManyFunct;
  -7: errmsg:=rsSignDigitMus;
  -8: errmsg:=rsCannotInsert;
  -9: errmsg:=rsCalcurationE;
  end;
  if FErrcode<>0 then
    raise ESimpleParseException.Create(errmsg);
end;

procedure FreeValueList;
begin
  while ValueList.Count>0 do
    _TokenFreeValue(mp_rat(ValueList.Last^));
end;


initialization
  ValueList := TList.Create;
  {$ifdef _THREADSAFEMODE}
  _CritSection:=TCriticalSection.Create;
  {$endif}
  FuncTree:=TFPHashList.Create;
  InitFuncList;

finalization
  FreeValueList;
  {$ifdef _THREADSAFEMODE}
  _CritSection.Free;
  {$endif}
  FuncTree.Free;

end.
