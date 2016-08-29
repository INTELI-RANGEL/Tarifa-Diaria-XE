unit frm_OpcionesAvances;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, frm_connection,
  ExtCtrls, DBCtrls, db, Menus, OleCtrls, 
  Buttons, RxMemDS, RXCtrls, 
   ZAbstractRODataset, ZDataset, 
  UnitExcepciones;
type
  TfrmOpcionesAvances = class(TForm)
    pgOpciones: TPageControl;
    pgPartidas: TTabSheet;
    CmdOk: TButton;
    CmdCancel: TButton;
    GroupPageRange: TGroupBox;
    DescrL: TLabel;
    optTodas: TRadioButton;
    OptReportadas: TRadioButton;
    opcPartidas: TRadioButton;
    EditPartidas: TEdit;
    GroupQuality: TGroupBox;
    chkMayor: TCheckBox;
    chkMenor: TCheckBox;
    chkIgual: TCheckBox;
    Label1: TLabel;
    ComboPorciento: TComboBox;
    pgEstructura: TTabSheet;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    optTodosPaquetes: TRadioButton;
    optNingunPaquete: TRadioButton;
    optNumeroPaquetes: TRadioButton;
    EditPaquetes: TEdit;
    mdReporte: TRxMemoryData;
    mdReportesContrato: TStringField;
    mdReportesNumeroActividad: TStringField;
    mdReportemDescripcion: TMemoField;
    mdReportedCantidad: TFloatField;
    mdReportesMedida: TStringField;
    mdReporteDia1: TFloatField;
    mdReporteDia2: TFloatField;
    mdReporteDia3: TFloatField;
    mdReporteDia4: TFloatField;
    mdReporteDia5: TFloatField;
    mdReporteDia6: TFloatField;
    mdReporteDia7: TFloatField;
    mdReporteDia8: TFloatField;
    mdReporteDia9: TFloatField;
    mdReporteDia10: TFloatField;
    mdReporteDia11: TFloatField;
    mdReporteDia12: TFloatField;
    mdReporteDia13: TFloatField;
    mdReporteDia14: TFloatField;
    mdReporteDia15: TFloatField;
    mdReporteDia16: TFloatField;
    mdReporteDia17: TFloatField;
    mdReporteDia18: TFloatField;
    mdReporteDia19: TFloatField;
    mdReporteDia20: TFloatField;
    mdReporteDia21: TFloatField;
    mdReporteDia22: TFloatField;
    mdReporteDia23: TFloatField;
    mdReporteDia24: TFloatField;
    mdReporteDia25: TFloatField;
    mdReporteDia26: TFloatField;
    mdReporteDia27: TFloatField;
    mdReporteDia28: TFloatField;
    mdReporteDia29: TFloatField;
    mdReporteDia30: TFloatField;
    mdReporteDia31: TFloatField;
    mdReportesWbs: TStringField;
    mdReportedTotal: TFloatField;
    mdReporteMes: TStringField;
    mdReporteAnio: TStringField;
    mdReportedIdFecha: TDateField;
    mdReportesWbsAnterior: TStringField;
    mdDatosAux: TRxMemoryData;
    StringField10: TStringField;
    StringField11: TStringField;
    MemoField3: TMemoField;
    FloatField1: TFloatField;
    StringField12: TStringField;
    FloatField2: TFloatField;
    FloatField3: TFloatField;
    FloatField5: TFloatField;
    FloatField6: TFloatField;
    FloatField7: TFloatField;
    FloatField8: TFloatField;
    FloatField9: TFloatField;
    FloatField10: TFloatField;
    FloatField11: TFloatField;
    FloatField12: TFloatField;
    FloatField13: TFloatField;
    FloatField14: TFloatField;
    FloatField15: TFloatField;
    FloatField16: TFloatField;
    FloatField17: TFloatField;
    FloatField18: TFloatField;
    FloatField19: TFloatField;
    FloatField20: TFloatField;
    FloatField21: TFloatField;
    FloatField22: TFloatField;
    FloatField23: TFloatField;
    FloatField24: TFloatField;
    FloatField25: TFloatField;
    FloatField26: TFloatField;
    FloatField27: TFloatField;
    FloatField28: TFloatField;
    FloatField29: TFloatField;
    FloatField30: TFloatField;
    FloatField31: TFloatField;
    FloatField32: TFloatField;
    FloatField33: TFloatField;
    StringField13: TStringField;
    FloatField34: TFloatField;
    StringField14: TStringField;
    StringField15: TStringField;
    DateField1: TDateField;
    mdDatosAuxsWbsAnterior: TStringField;
    Q_Detalle: TZReadOnlyQuery;
    Detalle: TRxMemoryData;
    StringField1: TStringField;
    StringField2: TStringField;
    MemoField1: TMemoField;
    FloatField4: TFloatField;
    StringField3: TStringField;
    StringField5: TStringField;
    StringField6: TStringField;
    DateField2: TDateField;
    StringField7: TStringField;
    DetallesWbs: TStringField;
    DetalledAvanceActual: TFloatField;
    mdDatosAuxdFechaInicio: TDateField;
    mdDatosAuxdFechaFinal: TDateField;
    mdReportedFechaInicio: TDateField;
    mdReportedFechaFinal: TDateField;
    DetalledFechaInicio: TDateField;
    DetalledFechaFinal: TDateField;
    pgAlcances: TTabSheet;
    GroupBox2: TGroupBox;
    Label3: TLabel;
    optTodosAlcances: TRadioButton;
    optUltimoAlcance: TRadioButton;
    optNumeroAlcance: TRadioButton;
    txtNumeroAlcance: TEdit;
    optNingunAlcance: TRadioButton;
    DetalleFase: TIntegerField;
    DetallesDescripcion: TStringField;
    mdReporteFase: TIntegerField;
    mdReportesDescripcion: TStringField;
    mdDatosAuxFase: TIntegerField;
    mdDatosAuxsDescripcion: TStringField;
    GroupBox3: TGroupBox;
    chkAllPaquetes: TRadioButton;
    chkEspPaquete: TRadioButton;
    txtIsometrico: TEdit;
    DetalledCantidadAnterior: TFloatField;
    DetalledAvanceAnterior: TFloatField;
    mdReportedTotalAnt: TFloatField;
    procedure FormShow(Sender: TObject);
    procedure CmdOkClick(Sender: TObject);
    procedure EditPaquetesEnter(Sender: TObject);
    procedure opcPartidasExit(Sender: TObject);
    procedure EditPartidasEnter(Sender: TObject);
    procedure ComboPorcientoChange(Sender: TObject);
    procedure txtNumeroAlcanceEnter(Sender: TObject);
    procedure optNumeroAlcanceExit(Sender: TObject);
    procedure txtIsometricoEnter(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmOpcionesAvances: TfrmOpcionesAvances;

implementation

uses
    frm_ReportePeriodo;

{$R *.dfm}

procedure TfrmOpcionesAvances.CmdOkClick(Sender: TObject);
var
  sCadena, sDia, MiWbs : String;
  iPos, NumPaq, i,x      : Integer;
  Q_Paquetes,
  Q_Fases              : tzReadOnlyquery ;
  ArrayPaquetes : array [1..10, 1..2] of String;

  {Variable de Opciones..}
  Nivel  : integer;
  Condicion : string;
  Partidas,
  sSigno    : string;
  lContinua : boolean;
  sAlcanceVar,
  sAlcanceLine,
  sAlcanceGrupo,
  sAlcance       : string;

  {Para busqueda de Alcances}
  Alcance,
  AlcanceAux    : string;
  lMuestra      : boolean;

  {Para Tablas de ActividadesxAnexo o ActividadexOrden}
  sCampo,
  sTabla,
  sLineaWbs,
  sLineaJoin,
  sLineaCondicion,
  sLineaCondicion2,
  sIsometrico, MiMes : String;
  MiFechaI, MiFechaF, MiFecha: tDate;
  Total  : double;
  //*******************BRITO 27/01/11*********************
  myYear, myMonth, myDay : Word;
  //*******************BRITO 27/01/11*********************

begin
  try

  except
    on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Opciones de Reporte', 'Al aceptar', 0);
    end;
  end;
end;

procedure TfrmOpcionesAvances.ComboPorcientoChange(Sender: TObject);
var
   cad : string;
   num : integer;
begin
     if pos('%', ComboPorciento.Text) <> 0 then
     begin
         cad := copy(Comboporciento.Text, 0 , pos('%', ComboPorciento.Text) - 2);
         try
             num := StrToInt(cad);
         Except
              messageDLG('Porcentaje Incorrecto!', mtError, [mbOk], 0);
              ComboPorciento.SetFocus;
         end;
     end
     else
     begin
         try
             num := StrToInt(cad);
         Except
              messageDLG('Porcentaje Incorrecto!', mtError, [mbOk], 0);
              ComboPorciento.SetFocus;
         end;
     end;
end;

procedure TfrmOpcionesAvances.EditPaquetesEnter(Sender: TObject);
begin
      optNumeroPaquetes.Checked := True;
end;

procedure TfrmOpcionesAvances.EditPartidasEnter(Sender: TObject);
begin
     opcPartidas.Checked := True;
end;

procedure TfrmOpcionesAvances.FormShow(Sender: TObject);
var
   i : integer;
begin


     if frmReportePeriodo.chkFases.Checked then
        pgAlcances.Enabled := True
     else
        pgAlcances.Enabled := False;

     ComboPorciento.Clear;
     for i := 1 to 100 do
         ComboPorciento.Items.Add(IntToStr(i) + ' %');

     ComboPorciento.ItemIndex := 0;

end;

procedure TfrmOpcionesAvances.opcPartidasExit(Sender: TObject);
begin
    EditPartidas.Text := '';
end;

procedure TfrmOpcionesAvances.optNumeroAlcanceExit(Sender: TObject);
begin
     txtNumeroAlcance.Text := '';
end;

procedure TfrmOpcionesAvances.txtIsometricoEnter(Sender: TObject);
begin
     chkespPaquete.Checked := True;
end;

procedure TfrmOpcionesAvances.txtNumeroAlcanceEnter(Sender: TObject);
begin
    optNumeroAlcance.Checked := True;
end;

end.
