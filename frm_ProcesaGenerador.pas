unit frm_ProcesaGenerador;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, AdvCombo, DB, ZAbstractRODataset, ZDataset,
  ComCtrls, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
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
  cxGroupBox, CurvyControls, cxCheckBox, cxTextEdit, cxMemo, Menus, cxButtons;

type
  TfrmProcesaGenerador = class(TForm)
    ZFolios: TZReadOnlyQuery;
    ZCeros: TZReadOnlyQuery;
    cxGroupBox1: TcxGroupBox;
    CmbFolios: TCurvyCombo;
    Label1: TLabel;
    cbxEquipo: TcxCheckBox;
    cbxPersonal: TcxCheckBox;
    mFolios: TcxMemo;
    Panel2: TPanel;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    PanelProgress: TPanel;
    Label15: TLabel;
    BarraEstado: TProgressBar;
    procedure Button1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
  private
    procedure ActualiaFactorGeneradorEQ(sParamEmbarcacion, sParamOrden,
      sParamFolio: string);
    procedure ActualiaFactorGeneradorPER(sParamEmbarcacion, sParamOrden,
      sParamFolio: string);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProcesaGenerador: TfrmProcesaGenerador;

implementation
uses frm_Connection,global;
{$R *.dfm}

procedure TfrmProcesaGenerador.Button1Click(Sender: TObject);
var xz:Integer;
begin
  if CmbFolios.ItemIndex = 0 then
  begin
    for xz := 1 to CmbFolios.Items.Count-1 do
    begin
      if CbxPersonal.Checked then
        ActualiaFactorGeneradorPER(global_barco, global_contrato,CmbFolios.Items[xz]);
      if CbxEquipo.Checked then
        ActualiaFactorGeneradorEQ(global_barco, global_contrato,CmbFolios.Items[xz]);
      PanelProgress.Visible := False;
    end;
  end
  else
  begin
    if CbxPersonal.Checked then
      ActualiaFactorGeneradorPER(global_barco, global_contrato,cmbFolios.text);
    if CbxEquipo.Checked then
      ActualiaFactorGeneradorEQ(global_barco, global_contrato,cmbFolios.text);
    PanelProgress.Visible := False;
  end;
end;

procedure TfrmProcesaGenerador.Button2Click(Sender: TObject);
begin
  ZCeros.Active := False;
  ZCeros.ParamByName('contrato').AsString := global_contrato;
  ZCeros.Open;
  MFolios.Clear;
  ZCeros.First;
  while not ZCeros.eof do
  begin
    MFolios.Lines.Add( ZCeros.FieldByName('folio').AsString);
    ZCeros.Next;
  end;
end;

procedure TfrmProcesaGenerador.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmProcesaGenerador.FormShow(Sender: TObject);
begin
  ZFolios.Active := False;
  ZFolios.ParamByName('contrato').AsString := global_contrato;
  ZFolios.Open;
  CmbFolios.Items.Clear;
  CmbFolios.Items.Add('<TODOS>');
  while not ZFolios.Eof do
  begin
    CmbFolios.Items.Add(ZFolios.FieldByName('snumeroorden').AsString);
    ZFolios.Next;
  end;
  CmbFolios.ItemIndex := 0;
end;

procedure TfrmProcesaGenerador.ActualiaFactorGeneradorPER(sParamEmbarcacion: string; sParamOrden: string; sParamFolio: string);
var
 zqFactoresPersonal,
 zqFolios,
 zqMovFolios,
 zqActualizaFolios : TZQuery;
 TotalFolios,
 CantPersonalFolio : Double;
 sFecha : string;
 Progreso, TotalProgreso: real;
begin
  //Funcion par aactualizar los factores de los folios antes de imprimir generador de barco JJF by ivan 2 Nov 2013
  zqFolios:=TZQuery.Create (Self);
  zqFolios.connection:= connection.zConnection;

  zqActualizaFolios:=TZQuery.Create (Self);
  zqActualizaFolios.connection:= connection.zConnection;

  zqFactoresPersonal := TZQuery.Create(Self);
  zqFactoresPersonal.Connection := Connection.zConnection;
  zqFactoresPersonal.Active := False ;
  zqFactoresPersonal.SQL.Clear ;
  zqFactoresPersonal.SQL.Add('select bp.sContrato, bp.sNumeroOrden, bp.sIdPersonal as IdRecurso, '+
                              'bp.dIdFecha, ROUND(SUM(bp.dCanthh), 2) as sFactor, SUM(bp.dCanthh) as sFactorTotal, (ROUND(SUM(bp.dCanthh), 2) - SUM(bp.dCanthh)) as dDiferencia from bitacoradepersonal bp '+
                              'Inner Join personal p on (bp.sContrato = p.sContrato and  bp.sIdPersonal = p.sIdPersonal and p.lCobro ="Si") '+
                              'Inner Join contratos c on (bp.sContrato = c.sContrato) '+
                              'where bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden  and bp.sNumeroOrden =:Folio and bp.dCantidad > 0 and bp.sTipoObra = "PU" '+
                              'Group By bp.sContrato, bp.sNumeroOrden, p.sIdPersonal, bp.dIdFecha order By bp.sContrato, bp.dIdFecha  asc ');
  zqFactoresPersonal.Params.ParamByName('Embarcacion').DataType := ftString ;
  zqFactoresPersonal.Params.ParamByName('Embarcacion').Value    := sParamEmbarcacion ;
  zqFactoresPersonal.Params.ParamByName('Orden').DataType       := ftString ;
  zqFactoresPersonal.Params.ParamByName('Orden').Value          := sParamOrden ;
  zqFactoresPersonal.Params.ParamByName('Folio').DataType       := ftString ;
  zqFactoresPersonal.Params.ParamByName('folio').Value          := sParamFolio ;
  zqFactoresPersonal.Open;

  sFecha := '';
  PanelProgress.Visible := True;
  zqFactoresPersonal.First;
  while not zqFactoresPersonal.Eof do
  begin
    Progreso := (1 / (zqFactoresPersonal.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
    TotalProgreso := TotalProgreso + Progreso;
    Label15.Caption := 'Procesando Personal...'+sParamFolio;
    Label15.Refresh;
    BarraEstado.Position := Trunc(TotalProgreso);

    {Actualizamos los valores de dCantHH a dCantHHGenerador}
    zqFolios.Active := False;
    zqFolios.SQL.Clear;
    zqFolios.SQL.Add('Update bitacoradepersonal set dCantHHGenerador = dCantHH '+
                     'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                     'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdPersonal =:Id ');
    zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresPersonal.FieldByName('dIdFecha').AsDateTime;
    zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
    zqFolios.ParamByName('Orden').AsString       := zqFactoresPersonal.FieldByName('sContrato').AsString;
    zqFolios.ParamByName('Folio').AsString       := zqFactoresPersonal.FieldByName('sNumeroOrden').AsString;
    zqFolios.ParamByName('Id').AsString          := zqFactoresPersonal.FieldByName('idRecurso').AsString;
    zqFolios.ExecSQL;

    {Consultamos el mayor factor de personal para aplicar ajuste}
    zqFolios.Active := False;
    zqFolios.SQL.Clear;
    zqFolios.SQL.Add('select * from bitacoradepersonal bp '+
                     'where bp.dIdFecha =:Fecha and bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden '+
                     'and bp.sNumeroOrden =:Folio and bp.sTipoObra = "PU" and bp.sIdPersonal =:Id '+
                     'order By bp.sFactor DESC limit 1 ');
    zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresPersonal.FieldByName('dIdFecha').AsDateTime;
    zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
    zqFolios.ParamByName('Orden').AsString       := zqFactoresPersonal.FieldByName('sContrato').AsString;
    zqFolios.ParamByName('Folio').AsString       := zqFactoresPersonal.FieldByName('sNumeroOrden').AsString;
    zqFolios.ParamByName('Id').AsString          := zqFactoresPersonal.FieldByName('idRecurso').AsString;
    zqFolios.Open;

    if zqFolios.RecordCount > 0 then
    begin
      zqActualizaFolios.Active := False;
      zqActualizaFolios.SQL.Clear;
      zqActualizaFolios.SQL.Add('Update bitacoradepersonal set dCantHHGenerador = :Cantidad '+
                       'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                       'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdPersonal =:Id and sIdPlataforma =:Plataforma and sHoraInicio =:HoraI and sHoraFinal =:HoraF and sDescripcion =:Descripcion ');
      zqActualizaFolios.ParamByName('Fecha').AsDateTime     := zqFolios.FieldByName('dIdFecha').AsDateTime;
      zqActualizaFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
      zqActualizaFolios.ParamByName('Orden').AsString       := zqFolios.FieldByName('sContrato').AsString;
      zqActualizaFolios.ParamByName('Folio').AsString       := zqFolios.FieldByName('sNumeroOrden').AsString;
      zqActualizaFolios.ParamByName('Id').AsString          := zqFolios.FieldByName('sIdPersonal').AsString;
      zqActualizaFolios.ParamByName('Plataforma').AsString  := zqFolios.FieldByName('sIdPlataforma').AsString;
      zqActualizaFolios.ParamByName('HoraI').AsString       := zqFolios.FieldByName('sHoraInicio').AsString;
      zqActualizaFolios.ParamByName('HoraF').AsString       := zqFolios.FieldByName('sHoraFinal').AsString;
      zqActualizaFolios.ParamByName('Cantidad').AsFloat     := zqFolios.FieldByName('dCantHHGenerador').AsFloat + zqFactoresPersonal.FieldByName('dDiferencia').AsFloat;
      zqActualizaFolios.ParamByName('Descripcion').AsString := zqFolios.FieldByName('sDescripcion').AsString ;
      zqActualizaFolios.ExecSQL;
    end;

    zqFactoresPersonal.Next;
  end;

end;

procedure TfrmProcesaGenerador.ActualiaFactorGeneradorEQ(sParamEmbarcacion: string; sParamOrden: string; sParamFolio: string);
var
 zqFactoresEquipo,
 zqFolios,
 zqMovFolios,
 zqActualizaFolios : TZQuery;
 TotalFolios,
 CantPersonalFolio : Double;
 sFecha : string;
 Progreso, TotalProgreso: real;
begin
  //Funcion par actualizar los factores de los folios antes de imprimir generador de barco JJF by ivan 2 Nov 2013
  zqFolios:=TZQuery.Create (Self);
  zqFolios.connection:= connection.zConnection;

  zqActualizaFolios:=TZQuery.Create (Self);
  zqActualizaFolios.connection:= connection.zConnection;

  zqFactoresEquipo := TZQuery.Create(Self);
  zqFactoresEquipo.Connection := Connection.zConnection;
  zqFactoresEquipo.Active := False ;
  zqFactoresEquipo.SQL.Clear ;
  zqFactoresEquipo.SQL.Add('select  bp.sContrato, bp.sNumeroOrden, bp.sIdEquipo as IdRecurso, '+
                           'bp.dIdFecha, ROUNd(sum(bp.dCantHH),2) as sFactor, SUM(bp.dCanthh) as sFactorTotal, (ROUND(SUM(bp.dCanthh), 2) - SUM(bp.dCanthh)) as dDiferencia '+
                           'from bitacoradeequipos bp '+
                           'Inner Join equipos p on (bp.sContrato = p.sContrato and  bp.sIdEquipo = p.sIdEquipo and p.lCobro ="Si") '+
                           'Inner Join contratos c on (bp.sContrato = c.sContrato) '+
                           'where bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden and bp.sNumeroOrden =:Folio and bp.dCantidad > 0 and bp.sTipoObra = "PU" '+
                           'Group By bp.sContrato, bp.sNumeroOrden, p.sIdEquipo, bp.dIdFecha order By bp.sContrato, bp.dIdFecha  asc ');
  zqFactoresEquipo.Params.ParamByName('Embarcacion').DataType := ftString ;
  zqFactoresEquipo.Params.ParamByName('Embarcacion').Value    := sParamEmbarcacion ;
  zqFactoresEquipo.Params.ParamByName('Orden').DataType       := ftString ;
  zqFactoresEquipo.Params.ParamByName('Orden').Value          := sParamOrden ;
  zqFactoresEquipo.Params.ParamByName('Folio').DataType       := ftString ;
  zqFactoresEquipo.Params.ParamByName('folio').Value          := sParamFolio ;
  zqFactoresEquipo.Open;

  sFecha := '';
  PanelProgress.Visible := True;
  zqFactoresEquipo.First;
  while not zqFactoresEquipo.Eof do
  begin
    Progreso := (1 / (zqFactoresEquipo.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
    TotalProgreso := TotalProgreso + Progreso;
    Label15.Caption := 'Procesando Equipo...'+sParamFolio;
    Label15.Refresh;
    BarraEstado.Position := Trunc(TotalProgreso);

    {Actualizamos los valores de dCantHH a dCantHHGenerador}
    zqFolios.Active := False;
    zqFolios.SQL.Clear;
    zqFolios.SQL.Add('Update bitacoradeequipos set dCantHHGenerador = dCantHH '+
                     'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                     'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdEquipo =:Id ');
    zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresEquipo.FieldByName('dIdFecha').AsDateTime;
    zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
    zqFolios.ParamByName('Orden').AsString       := zqFactoresEquipo.FieldByName('sContrato').AsString;
    zqFolios.ParamByName('Folio').AsString       := zqFactoresEquipo.FieldByName('sNumeroOrden').AsString;
    zqFolios.ParamByName('Id').AsString          := zqFactoresEquipo.FieldByName('idRecurso').AsString;
    zqFolios.ExecSQL;

    {Consultamos el mayor factor de personal para aplicar ajuste}
    zqFolios.Active := False;
    zqFolios.SQL.Clear;
    zqFolios.SQL.Add('select * from bitacoradeequipos bp '+
                     'where bp.dIdFecha =:Fecha and bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden '+
                     'and bp.sNumeroOrden =:Folio and bp.sTipoObra = "PU" and bp.sIdEquipo =:Id '+
                     'order By bp.sFactor DESC limit 1 ');
    zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresEquipo.FieldByName('dIdFecha').AsDateTime;
    zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
    zqFolios.ParamByName('Orden').AsString       := zqFactoresEquipo.FieldByName('sContrato').AsString;
    zqFolios.ParamByName('Folio').AsString       := zqFactoresEquipo.FieldByName('sNumeroOrden').AsString;
    zqFolios.ParamByName('Id').AsString          := zqFactoresEquipo.FieldByName('idRecurso').AsString;
    zqFolios.Open;

    if zqFolios.RecordCount > 0 then
    begin
      zqActualizaFolios.Active := False;
      zqActualizaFolios.SQL.Clear;
      zqActualizaFolios.SQL.Add('Update bitacoradeequipos set dCantHHGenerador = :Cantidad '+
                       'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                       'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdEquipo =:Id and sHoraInicio =:HoraI and sHoraFinal =:HoraF and sDescripcion =:Descripcion ');
      zqActualizaFolios.ParamByName('Fecha').AsDateTime     := zqFolios.FieldByName('dIdFecha').AsDateTime;
      zqActualizaFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
      zqActualizaFolios.ParamByName('Orden').AsString       := zqFolios.FieldByName('sContrato').AsString;
      zqActualizaFolios.ParamByName('Folio').AsString       := zqFolios.FieldByName('sNumeroOrden').AsString;
      zqActualizaFolios.ParamByName('Id').AsString          := zqFolios.FieldByName('sIdEquipo').AsString;
      zqActualizaFolios.ParamByName('HoraI').AsString       := zqFolios.FieldByName('sHoraInicio').AsString;
      zqActualizaFolios.ParamByName('HoraF').AsString       := zqFolios.FieldByName('sHoraFinal').AsString;
      zqActualizaFolios.ParamByName('Cantidad').AsFloat     := zqFolios.FieldByName('dCantHHGenerador').AsFloat + zqFactoresEquipo.FieldByName('dDiferencia').AsFloat;
      zqActualizaFolios.ParamByName('Descripcion').AsString := zqFolios.FieldByName('sDescripcion').AsString ;
      zqActualizaFolios.ExecSQL;
    end;

    zqFactoresEquipo.Next;
  end;

end;


end.
