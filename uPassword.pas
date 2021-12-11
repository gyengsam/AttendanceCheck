unit uPassword;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TfmPassword = class(TForm)
    edPassword: TEdit;
    Label1: TLabel;
    Button1: TButton;
    Label18: TLabel;
    Label2: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure edPasswordKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
    Password : string;
  end;

var
  fmPassword: TfmPassword;

implementation

{$R *.dfm}

procedure TfmPassword.Button1Click(Sender: TObject);
begin
  Password := edPassword.text;
  close;
end;

procedure TfmPassword.edPasswordKeyPress(Sender: TObject; var Key: Char);
begin
  if Key=#13 then
  begin
    Password := edPassword.text;
    close;
  end;
end;

procedure TfmPassword.FormShow(Sender: TObject);
begin
  Password := '';
  edPassword.Text := '';
end;

end.
