unit frm_ReporteDiario_Barco;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, 
  StdCtrls, DBCtrls, Buttons, 
  Menus, OleCtrls, ExtCtrls, 
  RXCtrls,
   rxSpeedbar, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint,
  dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
  dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
  dxSkinGlassOceans, dxSkinHighContrast, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMetropolis,
  dxSkinMetropolisDark, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Green, dxSkinOffice2007Pink,
  dxSkinOffice2007Silver, dxSkinOffice2010Black, dxSkinOffice2010Blue,
  dxSkinOffice2010Silver, dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray,
  dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic,
  dxSkinSharp, dxSkinSharpPlus, dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinTheAsphaltWorld, dxSkinsDefaultPainters,
  dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue,
  dxGDIPlusClasses, cxClasses, ImgList, cxImage, ComCtrls, cxTreeView, cxButtons;
type
  TfrmDiarioBarco = class(TForm)
    Panel2: TPanel;
    imgIconos: TcxImageList;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    cxButton3: TcxButton;
    cxButton4: TcxButton;
    cxButton5: TcxButton;
    cxButton6: TcxButton;
    cxButton7: TcxButton;

    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btTripulacionClick(Sender: TObject);
    procedure btMovimientosClick(Sender: TObject);
    procedure btProrrateoClick(Sender: TObject);
    procedure btnReportesDiariosClick(Sender: TObject);
    procedure btnGeneradoresClick(Sender: TObject);
    procedure btPernoctaClick(Sender: TObject);
    procedure btnPernoctaClick(Sender: TObject);
    procedure cxButton7Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDiarioBarco: TfrmDiarioBarco;
  sReporte : String ;
  iReporte : Integer ;
  lNuevoDia : Boolean ;
  dAvanceAnterior, dCantidadAnterior : Double ;
  dAvanceDiario, dAvanceAcumulado    : Double ;
  dPAnterior, dPDiario, dPAcumulado,
  dRAnterior, dRDiario, dRAcumulado : Real ;
  sArchivo     : String ;
  sOpcion      : String ;
  lIniciado    : Boolean ;
  SavePlace    : TBookmark;
implementation

uses 
  frm_tripulacion_diaria, frm_AdmonyTiempos, frm_ReporteDiarioTurno,
  frm_prorrateoPernocta, frm_PrintReportesDiarios,
  frm_Generadores_Barco, frm_cuadredepersonal, frm_lista_personal,
  frm_lista_personalV2, frm_tripulacion_pernoctas, frm_tripulacion_ajustes,
  frm_tripulacion_ajustesBarco, frm_tripulacion_generador;


{$R *.dfm}


procedure TfrmDiarioBarco.FormShow(Sender: TObject);
begin
  lIniciado := False ;

End;

procedure TfrmDiarioBarco.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   action := cafree ;
end;

procedure TfrmDiarioBarco.btTripulacionClick(Sender: TObject);
begin
     Application.CreateForm(TfrmTripulacionDiaria, frmTripulacionDiaria);
     frmTripulacionDiaria.ShowModal
end;

procedure TfrmDiarioBarco.cxButton7Click(Sender: TObject);
begin
    Application.CreateForm(TfrmTripulacionAjustesBarco,frmTripulacionAjustesBarco);
    frmTripulacionAjustesBarco.Show;
end;

procedure TfrmDiarioBarco.btMovimientosClick(Sender: TObject);
begin
    Application.CreateForm(TfrmAdmonyTiempos, frmAdmonyTiempos);
    frmAdmonyTiempos.ShowModal;
end;

procedure TfrmDiarioBarco.btPernoctaClick(Sender: TObject);
begin
    Application.CreateForm(TfrmTripulacionAjustes,frmTripulacionAjustes);
    frmTripulacionAjustes.Show;
end;

procedure TfrmDiarioBarco.btProrrateoClick(Sender: TObject);
begin
    Application.CreateForm(TfrmTripulacionPernoctas,frmTripulacionPernoctas);
    frmTripulacionPernoctas.Show;
end;

procedure TfrmDiarioBarco.btnGeneradoresClick(Sender: TObject);
begin
    //Application.CreateForm(TfrmGeneradoresBarco, frmGeneradoresBarco);
    //frmGeneradoresBarco.showmodal;
    Application.CreateForm(TfrmTripulacionGenerador,frmTripulacionGenerador);
    frmTripulacionGenerador.Show;
end;

procedure TfrmDiarioBarco.btnPernoctaClick(Sender: TObject);
begin
    Application.CreateForm(TfrmTripulacionDiaria,frmTripulacionDiaria);
    frmTripulacionDiaria.Show;
end;

procedure TfrmDiarioBarco.btnReportesDiariosClick(Sender: TObject);
begin
     Application.CreateForm(TfrmProrrateoPernocta, frmProrrateoPernocta);
    frmProrrateoPernocta.ShowModal
end;

end.


