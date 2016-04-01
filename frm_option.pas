unit frm_option;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, Forms, Controls, Graphics, Dialogs, ValEdit,
  StdCtrls;

type

  { TFormOption }

  TFormOption = class(TForm)
    ButtonOk: TButton;
    ButtonCancel: TButton;
    ValueListEditor1: TValueListEditor;
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  FormOption: TFormOption;

implementation

{$R *.lfm}

end.

