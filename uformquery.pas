unit uFormQuery;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, StdCtrls;

type

  { TFormQuery }

  TFormQuery = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    TextQuery: TStaticText;
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  FormQuery: TFormQuery;

implementation

{$R *.lfm}

end.

