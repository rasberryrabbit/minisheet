unit gridsime;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Grids {$ifdef MSWINDOWS}, messages {$endif};

type

    { TGridHan }

    TStringGrid = class(Grids.TStringGrid)
      private
        {$ifdef MSWINDOWS}
        procedure WMIMEStartComposition(var Msg:TMessage); message WM_IME_STARTCOMPOSITION;
        procedure WMIMEComposition(var Msg:TMessage); message WM_IME_COMPOSITION;
        {$endif}
      protected
      public
    end;

implementation



{$ifdef MSWINDOWS}
uses
  imm, Windows;

procedure TStringGrid.WMIMEStartComposition(var Msg: TMessage);
begin
  EditorMode:=True;
  if Assigned(InplaceEditor) then
    SendMessageW(InplaceEditor.Handle,WM_IME_STARTCOMPOSITION,Msg.wParam,Msg.lParam);
  Msg.Result:=-1;
end;

procedure TStringGrid.WMIMEComposition(var Msg: TMessage);
begin
  if Assigned(InplaceEditor) then
    SendMessageW(InplaceEditor.Handle,WM_IME_COMPOSITION,Msg.wParam,Msg.lParam);
  Msg.Result:=-1;
end;
{$endif}

end.

