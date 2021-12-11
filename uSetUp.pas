unit uSetUp;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls, zBase, zObjInspector, Vcl.Dialogs;

type
  TfmSetUp = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Panel1: TPanel;
    Button1: TButton;
    zObjSystem: TzObjectInspector;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmSetUp: TfmSetUp;





implementation
uses
  uCommonFunction, XSuperObject, uDataModule;

{$R *.dfm}

procedure TfmSetUp.Button1Click(Sender: TObject);
var
  SystemOptionFile, str : string;
  obj : ISuperObject;
begin
  SystemOptionFile := cDir + '\System.json';

  if OrgPassword <> SystemOption.AdminPassWord then
  begin
    SystemOption.AdminPassWord := EncryptString(SystemOption.AdminPassWord);
    OrgPassword := SystemOption.AdminPassWord;
  end;

  obj := TJSON.SuperObject<TSystemOption>(SystemOption);

  obj.SaveTo(SystemOptionFile);

  ShowMessage('System Option Saved');
end;

procedure TfmSetUp.FormShow(Sender: TObject);
begin
  zObjSystem.Component := nil;
  zObjSystem.Component := SystemOption;
end;

end.
