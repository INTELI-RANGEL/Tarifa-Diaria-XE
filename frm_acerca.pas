unit frm_acerca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, GradienteM, jpeg;

type
  TfrmAcerca = class(TForm)
    GradienteM1: TGradienteM;
    Label1: TLabel;
    Label2: TLabel;
    Label5: TLabel;
    btnOk: TBitBtn;
    Image1: TImage;
    Version: TLabel;
    procedure btnOkClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAcerca: TfrmAcerca;

implementation

{$R *.dfm}

procedure TfrmAcerca.btnOkClick(Sender: TObject);
begin
  frmAcerca.Close ;
end;


procedure TfrmAcerca.FormShow(Sender: TObject);
var
   InfoSize, H, RsltLen: Cardinal;
   VersionBlock: Pointer;
   Rslt: PVSFixedFileInfo;
begin
   InfoSize := GetFileVersionInfoSize(PChar(Application.ExeName), H);
   VersionBlock := AllocMem(InfoSize);
   try
      GetFileVersionInfo(PChar(Application.ExeName), H, InfoSize, VersionBlock);
      VerQueryValue(VersionBlock, '\', Pointer(Rslt), RsltLen);
      Version.Caption := 'Inteligent 2011 VC 1.1';
         //+ Format('%d.%d.%d.%d', [
         //Rslt.dwProductVersionMS div 65536,
         //Rslt.dwProductVersionMS mod 65536,
         //Rslt.dwProductVersionLS div 65536,
         //Rslt.dwProductVersionLS mod 65536]);
   finally
      FreeMem(VersionBlock);
   end;
end;

end.
