unit Frm_PopUpImportacionPP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, JvExControls, JvDBLookup, NxCollection, DB,
  ZAbstractRODataset, ZDataset, NxPageControl, AdvGroupBox, AdvCombo, Lucombo;

type
  TFrmPopUpImportacionPP = class(TForm)
    QrContratos: TZReadOnlyQuery;
    QrFolios: TZReadOnlyQuery;
    dsContratos: TDataSource;
    dsFolios: TDataSource;
    NxPCImportacion: TNxPageControl;
    NxTabSheet1: TNxTabSheet;
    NxTabSheet2: TNxTabSheet;
    Panel1: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    jDblCmbFolio: TJvDBLookupCombo;
    btnCancelar: TNxButton;
    Label1: TLabel;
    jDblCmbContrato: TJvDBLookupCombo;
    btnAceptar: TNxButton;
    Label4: TLabel;
    AdvGroupBox1: TAdvGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    btnAceptar2: TNxButton;
    btnCancelar2: TNxButton;
    lCmbPartida: TLUCombo;
    lCmbPonderado: TLUCombo;
    procedure FormShow(Sender: TObject);
    procedure jDblCmbContratoChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnAceptarClick(Sender: TObject);
    procedure btnCancelar2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    Procedure CargarValores;
  public
    { Public declarations }
    
  end;

var
  FrmPopUpImportacionPP: TFrmPopUpImportacionPP;

implementation
      //Texto1...texto30
uses frm_connection, global;

{$R *.dfm}

procedure TFrmPopUpImportacionPP.btnAceptarClick(Sender: TObject);
begin
  NxPCImportacion.ActivePageIndex:=1;
  CargarValores;
end;

procedure TFrmPopUpImportacionPP.btnCancelar2Click(Sender: TObject);
begin
  NxPCImportacion.ActivePageIndex:=0;
end;

Procedure TFrmPopUpImportacionPP.CargarValores;
var
  i:Integer;
begin
  lCmbPartida.Items.Clear;
  lCmbPonderado.Items.Clear;
  lCmbPonderado.Items.Add('No Aplica');
  for I := 1 to 30 do
  begin
    lCmbPartida.Items.Add('Texto' + IntTostr(I));
    lCmbPonderado.Items.Add('Texto' + IntTostr(I));
  end;
  lCmbPartida.ItemIndex:=0;
  lCmbPonderado.ItemIndex:=0;
end;

procedure TFrmPopUpImportacionPP.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (QrFolios.RecordCount=0) and (ModalResult=MrOk) then
  begin
    Messagedlg('No Existen Folios en esta OT.',MtError,[MbOk],0);
    CanClose:=false;
  end;
end;

procedure TFrmPopUpImportacionPP.FormCreate(Sender: TObject);
begin
  NxPCImportacion.ShowTabs:=false;
  NxPCImportacion.ActivePageIndex:=0;
end;

procedure TFrmPopUpImportacionPP.FormShow(Sender: TObject);
begin
  QrContratos.Active:=false;
  QrContratos.Open;

  jDblCmbContrato.KeyValue:=global_Contrato;

  QrFolios.Active:=false;
  QrFolios.Open;

  jDblCmbContratoChange(Sender);
end;

procedure TFrmPopUpImportacionPP.jDblCmbContratoChange(Sender: TObject);
begin
  if (QrFolios.Active) and (QrFolios.RecordCount>0) then
    jDblCmbFolio.KeyValue:=QrFolios.FieldByName('sNumeroOrden').AsString;
end;

end.
