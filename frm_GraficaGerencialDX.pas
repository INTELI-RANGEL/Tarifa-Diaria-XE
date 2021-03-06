unit frm_GraficaGerencialDX;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, StdCtrls, Buttons, ComCtrls, DBCtrls, DB, ADODB,
  ShellApi, global, RxMemDS, ExtCtrls,Math, jpeg,  OleServer,
  ZAbstractRODataset, ZDataset, ComObj, Grids,
  DBGrids, DateUtils, UnitExcel, Mask, rxToolEdit, Excel2000, ExcelXP,Utilerias,
  AdvCircularProgress;
type
  TPeriodo = Class
    Inicio, Termino: TDate;
  End;

  TfrmGraficaGerencialDX = class(TForm)
    btnGraph: TBitBtn;
    btnExit: TBitBtn;
    Opciones: TRadioGroup;
    ds_ordenesdetrabajo: TDataSource;
    ordenesdetrabajo: TZReadOnlyQuery;
    ReporteDiario: TZReadOnlyQuery;
    ProgressBar1: TProgressBar;
    qAnterior: TZReadOnlyQuery;
    Label3: TLabel;
    AvPaq: TRadioGroup;
    RadioGroup1: TRadioGroup;
    GroupBox1: TGroupBox;
    CheckBox1: TCheckBox;
    RadioButton1: TRadioButton;
    Label7: TLabel;
    RadioButton2: TRadioButton;
    Button1: TButton;
    Panel2: TPanel;
    Panel3: TPanel;
    Label8: TLabel;
    cbPeriodo: TComboBox;
    dInicio: TDateEdit;
    dTermino: TDateEdit;
    Panel1: TPanel;
    Label1: TLabel;
    tdIdFecha: TDateTimePicker;
    tdFechaInicio: TDateTimePicker;
    tsNumeroOrden: TDBLookupComboBox;
    roqOrdenAntes: TZReadOnlyQuery;
    roqOrdenHoy: TZReadOnlyQuery;
    dsOrden: TDataSource;
    roqAnexo: TZReadOnlyQuery;
    dsAnexo: TDataSource;
    cbPartidas: TCheckBox;
    cbTipoGrafica: TComboBox;
    Label10: TLabel;
    cbTipo: TComboBox;
    Label9: TLabel;
    Label11: TLabel;
    cbPeriodoOrden: TComboBox;
    Label2: TLabel;
    cbCalcular: TCheckBox;
    SdgExcel: TSaveDialog;
    PnlProcesando: TPanel;
    AdvCircularProgress1: TAdvCircularProgress;
    procedure FormShow(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tsArchivoKeyPress(Sender: TObject; var Key: Char);
    procedure btnExitClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tdFechaInicioEnter(Sender: TObject);
    procedure tdFechaInicioExit(Sender: TObject);
    procedure tdFechaInicioKeyPress(Sender: TObject; var Key: Char);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure btnGraphClick(Sender: TObject);
    Procedure Encabezado(oGantt: Variant; oHoja: Variant);
    Function Ajuste(oGantt: Variant; Valor : Real) : Real;
    procedure OpcionesClick(Sender: TObject);
    Procedure SeleccionaPeriodoReal(Var RangoPer: TPeriodo);
    procedure Button1Click(Sender: TObject);
    procedure GenerarNuevaGrafica;
    function cfnCalculaAvances1(sParamContrato, sParamOrden, sParamConvenio, sParamTurno: string; lParamMultiple: Boolean; dParamFecha: tDate; sParamSeguridad: string; tOrigen: TComponent): Currency;
    function cfnCalculaAvances2(sParamContrato, sParamOrden, sParamConvenio, sParamTurno: string; lParamMultiple: Boolean; dParamFecha: tDate; sParamSeguridad: string; tOrigen: TComponent): Currency;
    procedure cbPeriodoChange(Sender: TObject);
    procedure roqOrdenAntesFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure roqOrdenHoyFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    procedure cbTipoGraficaChange(Sender: TObject);
    procedure ordenesdetrabajoAfterScroll(DataSet: TDataSet);
    procedure AjustaPeriodoInicio;
    procedure cbPeriodoOrdenChange(Sender: TObject);
    procedure SdgExcelTypeChange(Sender: TObject);
    procedure AjustarTexto(var rangoE:Variant;TotalR:Integer);
    Procedure PFormatosExcel_EstablecerAreaDeImpresion(Hoja: Variant; ColumnaIzq, FilaIzq, ColumnaDer, FilaDer: Integer);
  private
    { Private declarations }
    Procedure EncabezadoDetalle(oGantt: Variant; oHoja: Variant);
    procedure ConfigurarHoja(var excel: Variant; var Hoja: Variant);
  public
    { Public declarations }
  end;
Const
  xlCenter = -4108;
  xlLeft = -4131;
  xlTop = -4160;
  xlGeneral = 1;
  xlSolid = 1;
  xlContinuous = 1;
  xlEdgeLeft = 7;
  xlThin = 2;
  xlEdgeTop = 8;
  xlEdgeBottom = 9;
  xlEdgeRight = 10;
  xlInsideVertical = 11;
  xlInsideHorizontal = 12;
  xlNormal = -4143;
  xlJustify = -4130;
  xlCellValue = 1;
  xlEqual = 3;
  msoFalse = 0;
  msoTrue = 1;
  msoShapeRoundedRectangle = 5;
  msoGradientHorizontal = 1;
  msoShapeRightTriangle = 8;


  IncHora = 1;
  MaxBloques = 240;
  ColorExced = 2;
  AnchoA =200;//544.2;// 58;
  OrAnchoColumnasGrafica = 6;

var
  OpcAnexo: Boolean;
  frmGraficaGerencialDX: TfrmGraficaGerencialDX;
  Lista: Array[1..14] Of Real;
  nCol: Real;     // N�mero de columnas a graficar
  RangoInicio: Real;
  BarraRango: Real;
  wCol: Real;
  aCol,SaltoCol: Integer;
  pFactor, pInicioP, pTerminoP: Real;
  AnchoColumasGrafica: Integer;
  AnchoColumnasPixels: Extended;
  OrInicio, OrTermino: TDateTime;
  nMes: Array[1..12] Of String;
  Periodo: TZReadOnlyQuery;
  mdPeriodo: TrxMemoryData;
  ListaPeriodos, RealPeriodos, TotalPeriodos: TStringList;
  OldCheck: Boolean;
  InicioSemana: Byte = 1;
  ItemIndex: Byte;
  RealDates: TPeriodo;
  FirstShow: Boolean;
  tmpNombre,tmpNombreC: String;
  rAncho :Integer= 17;      // N�mero de columnas correspondientes a datos generales de partidas
  rAlto :Integer = 11;       // N�mero de renglones correspondientes al encabezado
implementation

uses frm_ReporteDiarioTurno;

{$R *.dfm}



procedure TfrmGraficaGerencialDX.ConfigurarHoja(var excel: Variant; var Hoja: Variant);
var
  pfHoja: Byte;
  SubCad,CadError: String;
  Difer, AcumDifer: Extended;
  sFirmante1,sFirmante2,sPuesto1,sPuesto2:string;
  QryBuscarFirmas: tzReadOnlyQuery;
  fs:TStream;
  //tmpNombreC:string;
  imgAux:TImage;
  Pic : TJpegImage;
  TempPath: array [0..MAX_PATH-1] of Char;
  FNombre1,FNombre2:TFileName;
  sCadT:string;
begin

  // Seleccionar el periodo de firmantes
  application.ProcessMessages;



   Application.ProcessMessages;

    // Poner las firmas en todas las hojas del libro generado
    try


        Excel.ActiveSheet.PageSetup.PaperSize := xlPaperLetter;
        Excel.ActiveWindow.View :=xlPageLayoutView;

        //Excel.PrintCommunication := true;
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;
        Excel.ActiveSheet.PageSetup.PrintTitleRows := '$1:$9';

        Application.ProcessMessages;
        Excel.ActiveSheet.PageSetUp.CenterFooter :='';
        Excel.ActiveSheet.PageSetUp.LeftFooter := '';
        Excel.ActiveSheet.PageSetUp.RightFooter := '';
        Excel.ActiveSheet.PageSetUp.LeftMargin     := 0;
        Excel.ActiveSheet.PageSetUp.RightMargin    := 0;
        Excel.ActiveSheet.PageSetUp.TopMargin      := 14;
        Excel.ActiveSheet.PageSetUp.BottomMargin   := 14;
        Excel.ActiveSheet.PageSetUp.HeaderMargin   := 0;
        Excel.ActiveSheet.PageSetUp.FooterMargin   :=Excel.InchesToPoints(0); //56;
        Excel.ActiveSheet.PageSetUp.Zoom := 32;
        application.ProcessMessages;



         Excel.ActiveSheet.PageSetUp.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetUp.AlignMarginsHeaderFooter := True;
        Excel.ActiveSheet.PageSetUp.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.FirstPage.RightFooter.Text := '';

        Excel.ActiveSheet.PageSetup.Orientation := xlLandscape;
        Excel.ActiveSheet.PageSetUp.Zoom           := False;
        Excel.ActiveSheet.PageSetUp.FitToPagesWide := 1;
        Excel.ActiveSheet.PageSetUp.FitToPagesTall := False;

        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := True;

         application.ProcessMessages;
        Excel.ActiveWindow.View := xlPageLayoutView;
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;
        Excel.ActiveSheet.PageSetup.LeftHeader := '';
        Excel.ActiveSheet.PageSetup.CenterHeader := '';
        Excel.ActiveSheet.PageSetup.RightHeader := '';
        Excel.ActiveSheet.PageSetup.LeftFooter := '';
        Excel.ActiveSheet.PageSetup.CenterFooter := '';
        Excel.ActiveSheet.PageSetup.RightFooter := '';
        Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0.5);
        Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.194444444444444);
        Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(0.194444444444444);
        Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.Zoom := 32;
        Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
        Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := true;
        Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
        if (Excel.Application.version >= 14) then
        begin
          Excel.PrintCommunication := True;
          Excel.PrintCommunication := False;
        end;
        application.ProcessMessages;
        Excel.ActiveSheet.PageSetup.LeftHeader := '';
        Excel.ActiveSheet.PageSetup.CenterHeader := '';
        Excel.ActiveSheet.PageSetup.RightHeader := '';
        Excel.ActiveSheet.PageSetup.LeftFooter := '';
        Excel.ActiveSheet.PageSetup.CenterFooter := '';
        Excel.ActiveSheet.PageSetup.RightFooter := '';
        Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0.5);
        Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.194444444444444);
        Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(0.194444444444444);
        Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.Zoom := 32;
        Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
        Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := true;
        Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := True;

        application.ProcessMessages;


        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;


        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := True;
        Excel.ActiveSheet.PageSetup.PrintArea := '';
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;

        application.ProcessMessages;
        Excel.ActiveSheet.PageSetup.LeftHeader := '';
        Excel.ActiveSheet.PageSetup.CenterHeader := '';
        Excel.ActiveSheet.PageSetup.RightHeader := '';
       // Excel.ActiveSheet.PageSetup.LeftFooter :='&G';
       // Excel.ActiveSheet.PageSetup.CenterFooter :='&"Arial,Normal"&'+inttostr(TamFont)+'&P de &#';//'&Z&G&P de &#&D&G'; //'&P de &N';
        //"&""Arial,Normal""&14&P de &#"
       // Excel.ActiveSheet.PageSetup.RightFooter :='&G';
        Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0.5);
        Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.196850393700787);
        Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(0.196850393700787);
        Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.PrintHeadings := False;
        Excel.ActiveSheet.PageSetup.PrintGridlines := False;
       // Excel.ActiveSheet.PageSetup.PrintComments := xlPrintNoComments;  //
        Excel.ActiveSheet.PageSetup.PrintQuality := 600;
        Excel.ActiveSheet.PageSetup.CenterHorizontally := False;
        Excel.ActiveSheet.PageSetup.CenterVertically := False;
      { if   pfHoja>(Excel.Sheets.Count-iTotalLandscape) then
        Excel.ActiveSheet.PageSetup.Orientation := xlLandscape
        else
         Excel.ActiveSheet.PageSetup.Orientation := xlPortrait; }
        Excel.ActiveSheet.PageSetup.Draft := False;
        try
          Excel.ActiveSheet.PageSetup.PaperSize := xlPaperLetter;
        except
          Excel.ActiveSheet.PageSetup.PaperSize :=119;
        end;
        Excel.ActiveSheet.PageSetup.FirstPageNumber := xlAutomatic;
        Excel.ActiveSheet.PageSetup.Order := xlDownThenOver;
        Excel.ActiveSheet.PageSetup.BlackAndWhite := False;
        Excel.ActiveSheet.PageSetup.Zoom := 32;
        Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
        Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
         Excel.ActiveSheet.PageSetUp.Zoom           := False;
        Excel.ActiveSheet.PageSetUp.FitToPagesWide := 1;
        Excel.ActiveSheet.PageSetUp.FitToPagesTall := False;
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := True;

        Application.ProcessMessages;
        Excel.ActiveWindow.View := xlPageBreakPreview;
        Excel.ActiveWindow.Zoom := 18;
    except
      on e:exception do
      begin
        ShowMessage(e.Message + ', ' + e.ClassName);
        CadError := 'No se puede generar el pie de p�gina';

      end;
    end;

end;




procedure TfrmGraficaGerencialDX.AjustarTexto(var rangoE: Variant;TotalR:Integer);
var
  sngAnchoTotal,sngAnchoCelda,sngAlto,sngResta:Extended;
  n:Integer;
begin
  sngAnchoTotal:=0;
  For n := 1 To TotalR do
    sngAnchoTotal := sngAnchoTotal + rangoE.columns.columns[n].ColumnWidth;

  sngAnchoCelda :=rangoE.columns.columns[1].ColumnWidth;
//  rangoE.HorizontalAlignment := xlJustify;
//  rangoE.VerticalAlignment := xlcenter;
  rangoE.MergeCells := False;


  if sngAnchoTotal>255 then
    rangoE.columns.columns[1].ColumnWidth :=255
  else
    rangoE.columns.columns[1].ColumnWidth := sngAnchoTotal;

  rangoE.parent.rows[rangoE.row].Autofit;
  sngAlto :=rangoE.RowHeight;

   if sngAlto>150 then
    sngResta:=100;

   if sngAlto>200 then
    sngResta:=150;

   if sngAlto>300 then
    sngResta:=200;

    if sngAlto>400 then
    sngResta:=220;

  rangoE.Merge;
  rangoE.Columns[1].EntireColumn.ColumnWidth := sngAnchoCelda;
  if sngAlto >409 then
    rangoE.Columns[1].RowHeight :=409 -sngResta
  else
  rangoE.Columns[1].RowHeight := sngAlto -sngResta;
  application.ProcessMessages;
end;



Procedure GeneraLineas(Objeto : Variant; Cadena : String);
Var
  Seccion, Rango : Variant;
Begin
    Rango := Objeto.Range[Cadena];
    Rango.Select;

    Seccion := Rango.Borders[xlEdgeLeft];
    Seccion.LineStyle := xlContinuous;
    Seccion.Weight := xlThin;

    Seccion := Rango.Borders[xlEdgeTop];
    Seccion.LineStyle := xlContinuous;
    Seccion.Weight := xlThin;

    Seccion := Rango.Borders[xlEdgeBottom];
    Seccion.LineStyle := xlContinuous;
    Seccion.Weight := xlThin;

    Seccion := Rango.Borders[xlEdgeRight];
    Seccion.LineStyle := xlContinuous;
    Seccion.Weight := xlThin;

    Seccion := Rango.Borders[xlInsideVertical];
    Seccion.LineStyle := xlContinuous;
    Seccion.Weight := xlThin;

    Seccion := Rango.Borders[xlInsideHorizontal];
    Seccion.LineStyle := xlContinuous;
    Seccion.Weight := xlThin;


End;

Procedure FormatoDatos(Objeto : Variant; Renglon : Integer);
Var
  Rango: Variant;
Begin
  Rango := Objeto.Columns['B:B'];
  Rango.WrapText := True;

  Rango := Objeto.Range['A9:A9'];
  Rango.WrapText := True;
  Rango.HorizontalAlignment := xlCenter;
  Rango := Objeto.Range['B9:B9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['C9:C9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['D9:D9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['E9:E9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['F9:F9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['G9:G9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['H9:H9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['I9:I9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['J9:J9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['K9:K9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['L9:L9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['M9:M9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['N9:N9'];
  Rango.WrapText := True;
  Rango := Objeto.Range['O9:O9'];
  Rango.WrapText := True;
End;

Procedure TfrmGraficaGerencialDX.PFormatosExcel_EstablecerAreaDeImpresion(Hoja: Variant; ColumnaIzq, FilaIzq, ColumnaDer, FilaDer: Integer);
begin
  if (ColumnaIzq <= 0) or (ColumnaDer <= 0)
  Then raise Exception.Create('');
  Hoja.PageSetup.PrintArea := '$' + ColumnaNombre(ColumnaIzq) + '$' + IntToStr(FilaIzq) + ':' +
                              '$' + ColumnaNombre(ColumnaDer) + '$' + IntToStr(FilaDer);
end;

procedure TfrmGraficaGerencialDX.AjustaPeriodoInicio;
begin
  case cbTipo.ItemIndex of
    1: begin
      // Semanal
      if DayOfWeek(RealDates.Termino) <> InicioSemana then
        while DayOfWeek(RealDates.Termino) <> InicioSemana do
          RealDates.Termino := IncDay(RealDates.Termino, -1);
      RealDates.Termino := IncDay(RealDates.Termino, 6);

      while DayOfWeek(RealDates.Inicio) <> InicioSemana do
        RealDates.Inicio := IncDay(RealDates.Inicio, -1);
    end;
    2: begin
      // Mensual
      while DayOfTheMonth(RealDates.Inicio) <> 1 do
        RealDates.Inicio := IncDay(RealDates.Inicio, -1);
    end;
    3: begin
      // Anual
      while (MonthOfTheYear(RealDates.Inicio) <> 1) or (DayOfTheMonth(RealDates.Inicio) <> 1) do
        RealDates.Inicio := IncDay(RealDates.Inicio, -1);
    end;
  end;
  pInicioP := RealDates.Inicio;
  pTerminoP := RealDates.Termino + 1;
end;

procedure TfrmGraficaGerencialDX.FormShow(Sender: TObject);
var
  Orig1, Orig2: TDate;
  vPeriodo,
  RangoPer: TPeriodo;
begin
  FirstShow := True;

  nMes[01] := 'ENERO';
  nMes[02] := 'FEBRERO';
  nMes[03] := 'MARZO';
  nMes[04] := 'ABRIL';
  nMes[05] := 'MAYO';
  nMes[06] := 'JUNIO';
  nMes[07] := 'JULIO';
  nMes[08] := 'AGOSTO';
  nMes[09] := 'SEPTIEMBRE';
  nMes[10] := 'OCTUBRE';
  nMes[11] := 'NOVIEMBRE';
  nMes[12] := 'DICIEMBRE';

  // Obtener la fecha de inicio y termino por anexo
  Connection.zCommand.Close;
  Connection.zCommand.SQL.Text := 'select min(dfechainicio) as inicio, max(dfechafinal) as termino from actividadesxanexo ' +
                                  'where scontrato = ' + quotedstr(global_contrato) + ' and sidconvenio = ' + quotedstr(global_convenio);
  Connection.zCommand.Open;
  if Connection.zCommand.RecordCount > 0 then
  Begin
    if Connection.zCommand.FieldValues['inicio'] = Null then
      OrInicio :=  date
    else
      OrInicio :=  Connection.zCommand.FieldValues['inicio'];

    if  Connection.zCommand.FieldValues['termino'] = Null then
      OrTermino := date
    else
      OrTermino := Connection.zCommand.FieldValues['termino'];
  End
  Else
  Begin
    OrInicio := date;
    OrTermino := date;
  End;

  OrdenesdeTrabajo.Active := False ;
  OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
  OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato ;
  ordenesdetrabajo.Params.ParamByName('status').DataType := ftString ;
  ordenesdetrabajo.Params.ParamByName('status').Value :=  connection.configuracion.FieldValues [ 'cStatusProceso' ];
  OrdenesdeTrabajo.Open ;
  tdidFecha.Date := Date ;
  if Panel1.Visible then
    If OrdenesdeTrabajo.RecordCount > 0 Then
    Begin
      tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ;
      tdIdFecha.SetFocus
    End
    Else
      tsNumeroOrden.SetFocus
  else
    cbPeriodo.SetFocus;

  // Obtener la fecha de inicio y termino por frentes de trabajo
  ListaPeriodos := TStringList.Create;
  ListaPeriodos.Clear;
  RealPeriodos := TStringList.Create;
  RealPeriodos.Clear;
  TotalPeriodos := TStringList.Create;
  TotalPeriodos.Clear;

  while not OrdenesdeTrabajo.Eof do
  begin
    RangoPer := TPeriodo.Create;
    RangoPer.Inicio := 0;
    RangoPer.Termino := 0;

    // Encontrar la duraci�n de este frente
    Connection.zCommand.Close;
    Connection.zCommand.SQL.Text := 'select min(dfechainicio) as inicio, max(dfechafinal) as termino from actividadesxorden ' +
                                    'where scontrato = ' + quotedstr(global_contrato) + ' and sidconvenio = ' + quotedstr(global_convenio) + ' and snumeroorden = ' + quotedstr(OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString);
    Connection.zCommand.Open;

    if (RangoPer.Inicio = 0) or (RangoPer.Inicio > Connection.zCommand.FieldByName('Inicio').AsDateTime) then
      RangoPer.Inicio := Connection.zCommand.FieldByName('Inicio').AsDateTime;

    if (RangoPer.Termino = 0) or (RangoPer.Termino < Connection.zCommand.FieldByName('termino').AsDateTime) then
      RangoPer.Termino := Connection.zCommand.FieldByName('termino').AsDateTime;

    vPeriodo := TPeriodo.Create;
    if Connection.zCommand.FieldValues['inicio'] = Null then
    begin
      vPeriodo.Inicio := date;
      vPeriodo.Termino := date;
    end
    else
    begin
      vPeriodo.Inicio := Connection.zCommand.FieldValues['inicio'];
      vPeriodo.Termino := Connection.zCommand.FieldValues['termino'];
    end;

    ListaPeriodos.AddObject(OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString, vPeriodo);

    SeleccionaPeriodoReal(RangoPer);  // Seleccionar el periodo real para esta orden de trabajo

    TotalPeriodos.AddObject(OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString, RangoPer);

    OrdenesdeTrabajo.Next;
  end;
  OrdenesdeTrabajo.First;

  if global_OptGrafica = 'Frente' then
    opciones.ItemIndex := 1;

  if ListaPeriodos.Count > 0 then
  begin
    tdFechaInicio.Date := TPeriodo(ListaPeriodos.Objects[0]).Inicio;
    tdIdFecha.Date := TPeriodo(ListaPeriodos.Objects[0]).Termino;
  end
  else
  begin
    tdFechaInicio.Date := Date;
    tdIdFecha.Date := Date;
  end;

  Periodo := TZReadOnlyQuery.Create(Nil);
  Periodo.Connection := Connection.ConnTrx;
  Periodo.Sql.Text := 'select "PROGRAMADO" as sTipo, ' +
                      'IF((select min(dFechaInicio) from actividadesxanexo ' +
                      'where scontrato = :contrato and sIdConvenio = :convenio) < (select min(dFechaInicio) from actividadesxorden where scontrato = :contrato and sidconvenio = :convenio), ' +
                      '(select min(dFechaInicio) from actividadesxanexo where scontrato = :contrato and sidconvenio = :convenio), ' +
                      '(select min(dFechaInicio) from actividadesxorden where scontrato = :contrato and sIdConvenio = :convenio)) as dInicio, ' +
                      'IF((select Max(dFechaFinal) from actividadesxanexo ' +
                      'where scontrato = :contrato and sIdConvenio = :convenio) < (select Max(dFechaFinal) from actividadesxorden where scontrato = :contrato and sidconvenio = :convenio), ' +
                      '(select Max(dFechaFinal) from actividadesxanexo where scontrato = :contrato and sidconvenio = :convenio), ' +
                      '(select Max(dFechaFinal) from actividadesxorden where scontrato = :contrato and sIdConvenio = :convenio)) as dTermino ' +
                      'union ' +
                      'select "REAL" as sTipo, min(a.dIdFecha) as dInicio, max(a.dIdFecha) as dTermino ' +
                      'from bitacoradeactividades a where a.sContrato = :Contrato';
  Periodo.ParamByName('contrato').AsString := global_contrato;
  Periodo.ParamByName('convenio').AsString := global_convenio;
  Periodo.Open;

  mdPeriodo := TrxMemoryData.Create(Self);
  mdPeriodo.FieldDefs.Add('sTipo', ftString, 30, True);
  mdPeriodo.FieldDefs.Add('dInicio', ftDate, 0, True);
  mdPeriodo.FieldDefs.Add('dTermino', ftDate, 0, True);
  mdPeriodo.Open;
  mdPeriodo.EmptyTable;

  // Agregar el periodo completo
  Periodo.RecNo := 1;
  Orig1 := Periodo.FieldByName('dInicio').AsDateTime;
  Orig2 := Periodo.FieldByName('dTermino').AsDateTime;
  Periodo.RecNo := 2;
  if Periodo.FieldByName('dInicio').AsDateTime < Orig1 then
    Orig1 := Periodo.FieldByName('dInicio').AsDateTime;
  if Periodo.FieldByName('dTermino').AsDateTime > Orig2 then
    Orig2 := Periodo.FieldByName('dTermino').AsDateTime;

  mdPeriodo.Append;
  mdPeriodo.FieldByName('sTipo').AsString := 'COMPLETO';
  mdPeriodo.FieldByName('dInicio').AsDateTime := Orig1;
  mdPeriodo.FieldByName('dTermino').AsDateTime := Orig2;
  mdPeriodo.Post;

  Periodo.First;
  while Not Periodo.Eof do
  begin
    mdPeriodo.Append;
    mdPeriodo.CopyFields(Periodo);
    mdPeriodo.Post;

    Periodo.Next;
  end;

  cbPeriodo.Items.Clear;
  mdPeriodo.First;
  while Not mdPeriodo.Eof do
  begin
    cbPeriodo.Items.Add(mdPeriodo.FieldByName('sTipo').AsString);
    mdPeriodo.Next;
  end;
  if cbPeriodo.Items.Count > 0 then
  begin
    cbPeriodo.ItemIndex := 0;
    cbPeriodo.OnChange(Self);
  end;
end;

procedure TfrmGraficaGerencialDX.OpcionesClick(Sender: TObject);
begin
  Panel3.Visible := Opciones.ItemIndex = 0;
  Panel1.Visible := Not Panel3.Visible;
  RadioGroup1.Enabled := Opciones.ItemIndex = 0;
  if Opciones.ItemIndex = 0 then
    Label3.Visible := False
  else
  begin
    // Graficar por frente de trabajo
    if tsNumeroOrden.KeyValue = Null then
      tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString;
    OrdenesdeTrabajo.CursorPosChanged;
    RadioGroup1.ItemIndex := 1;
    if FirstShow then
      cbPeriodoOrden.OnChange(Sender);
    FirstShow := False;
  end;
end;

procedure TfrmGraficaGerencialDX.ordenesdetrabajoAfterScroll(DataSet: TDataSet);
begin
  // Poner las fechas del periodo programado de acuerdo al frente seleccionado
  cbPeriodoOrden.ItemIndex := 0;
  cbPeriodoOrden.OnChange(Nil);
end;

procedure TfrmGraficaGerencialDX.SdgExcelTypeChange(Sender: TObject);
begin
  case SdgExcel.FilterIndex of
    1 : SdgExcel.DefaultExt := 'xls';
    2 : SdgExcel.DefaultExt := 'xlsx';
  else
    SdgExcel.DefaultExt := '';
  end;
end;

procedure TfrmGraficaGerencialDX.SeleccionaPeriodoReal(Var RangoPer: TPeriodo);
Var
  pasoInicio, pasoTermino: TDateTime;
  CadSql: String;
  vPeriodo: TPeriodo;
begin
{  if Global_OPtGrafica = 'Frente' then
    Opciones.ItemIndex := 1;}

  PasoInicio := 0;
  PasoTermino := 0;
  vPeriodo := TPeriodo.Create;

  CadSql := 'select MIN(didfecha) as inicio, MAX(didfecha) as termino from bitacoradeactividades where scontrato = ' + quotedstr(global_contrato) + ' and swbs <> ""';
  //if Opciones.ItemIndex = 1 then
    CadSql := CadSql + ' and snumeroorden = ' + quotedstr(ordenesdetrabajo.FieldValues['snumeroorden']);
  Connection.zCommand.Close;
  Connection.zCommand.SQL.Text := CadSql;
  Connection.zCommand.Open;

  if Connection.zCommand.FieldValues['inicio'] <> Null then
  Begin
    Label3.Visible := False;
    if (pasoInicio = 0) Or (Connection.zCommand.FieldValues['inicio'] < pasoInicio) then
      pasoInicio := Connection.zCommand.FieldValues['inicio'];

    if (pasoTermino = 0) Or (Connection.zCommand.FieldValues['termino'] > pasoTermino) then
      pasoTermino := Connection.zCommand.FieldValues['termino'];

    vPeriodo.Inicio := pasoinicio;
    vPeriodo.Termino := pasotermino + 1;
  End
  Else
  begin
    vPeriodo.Inicio := OrInicio;
    vPeriodo.Termino := OrInicio;
  end;

  if vPeriodo.Inicio < RangoPer.Inicio then
    RangoPer.Inicio := vPeriodo.Inicio;

  if vPeriodo.Termino > RangoPer.Termino then
    RangoPer.Termino := vPeriodo.Termino;

  RealPeriodos.AddObject(ordenesdetrabajo.FieldByName('snumeroorden').AsString, vPeriodo);
end;

procedure TfrmGraficaGerencialDX.roqOrdenAntesFilterRecord(DataSet: TDataSet;
  var Accept: Boolean);
var
  Cadena: String;
begin
  if OpcAnexo then
  begin
    Cadena := roqAnexo.FieldByName('sWbs').AsString + '.';
    Accept := (Copy(roqOrdenAntes.FieldByName('sWbsContrato').AsString, 1, Length(Cadena)) = Cadena);
              //and (roqOrdenAntes.FieldByName('dAvanceAnexoPondDirecto').AsFloat<>0);
  end
  else
    Accept := (roqOrdenAntes.FieldByName('sWbsContrato').AsString = roqAnexo.FieldByName('sWbs').AsString);
              //and (roqOrdenAntes.FieldByName('dAvanceAnexoPondDirecto').AsFloat<>0);
end;

procedure TfrmGraficaGerencialDX.roqOrdenHoyFilterRecord(DataSet: TDataSet;
  var Accept: Boolean);
var
  Cadena: String;
begin
  if OpcAnexo then
  begin
    Cadena := roqAnexo.FieldByName('sWbs').AsString + '.';
    Accept := (Copy(roqOrdenHoy.FieldByName('sWbsContrato').AsString,1,Length(Cadena)) = Cadena);
              //and (roqOrdenHoy.FieldByName('dAvanceAnexoPondDirecto').AsFloat<>0);
  end
  else
    Accept := (roqOrdenHoy.FieldByName('sWbsContrato').AsString = roqAnexo.FieldByName('sWbs').AsString);
              //and (roqOrdenHoy.FieldByName('dAvanceAnexoPondDirecto').AsFloat<>0);
end;

procedure TfrmGraficaGerencialDX.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 then
        tdIdFecha.SetFocus
end;

procedure TfrmGraficaGerencialDX.tsArchivoKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 then
        tsNumeroOrden.SetFocus
end;

procedure TfrmGraficaGerencialDX.btnExitClick(Sender: TObject);
begin
  close
end;

Function TfrmGraficaGerencialDX.Ajuste(oGantt: Variant; Valor: Real) : Real;
Var
  Resul : Real;
Begin
  if oGantt.Application.version >= 12 then
    Resul := Valor / 1.0070
  Else
    Resul := Valor;
  Ajuste := Resul
End;

Function Month(Fecha: TDateTime) : Word;
Var
  Anio, Mes, Dia : Word;
Begin
  DecodeDate(Fecha,Anio,Mes,Dia);
  Month := Mes;
End;

Function QuitarNoValidos(Cadena: String) : String;
Var
  X: Integer;
  Resultado: String;
Const
  NoValidos = '\/:*?"<>|';
Begin
  // Eliminar los caracteres no validos para un nombre de archivo
  X := 0;
  Resultado := '';

  while x < Length(Cadena) do
  Begin
    Inc(x);
    if Pos(Cadena[x],NoValidos) = 0 then
      Resultado := Resultado + Cadena[x];
  End;

  QuitarNoValidos := Resultado;
End;

Function RightStr
    (Const Str: String; Size: Word): String;
Begin
  If Size > Length(Str) Then Size := Length(Str);
  RightStr := Copy(Str, Length(Str)-Size+1, Size)
End;

Function FechaCadena(Fecha: TDateTime; Modo: Byte): String;
Var
  Hora, Min, Seg, Mili: Word;
  Resultado: String;
Begin
  // Convierte la fecha/hora en sus fracciones
  DecodeTime(Fecha,Hora,Min,Seg,Mili);
  if Modo = 1 then
  Begin
    if Hora + Min + Seg = 0 then
      Hora := 24;
  End;
  Resultado := Chr(39) + DateToStr(Fecha) + ' ' + RightStr('0' + IntToStr(Hora),2) + ':' + RightStr('0' + IntToStr(Min),2) + ':' + RightStr('0' + IntToStr(Seg),2);

  FechaCadena := Resultado
End;

Function QuitarEnter(Cadena : String) : String;
Var
  vPos: Integer;
  Resul: String;
Begin
  vPos := 1;
  while vPos <= Length(Cadena) do
  Begin
    if Cadena[vPos] <> chr(13) then
      Resul := Resul + Cadena[vPos];
    Inc(vPos);
  End;
  QuitarEnter := Resul;
End;

Procedure SetRowHeight(Objeto : Variant; Ren : Integer; Alto : Real);
Var
  CadRow : String;
  Renglon : Variant;
Begin
  CadRow := Trim(IntToStr(Ren)) + ':' + Trim(InttoStr(Ren));
  Renglon := Objeto.Rows[CadRow];
  Renglon.RowHeight := Alto;
End;

Function RowHeight(Objeto : Variant; Ren : Integer) : Real;
Var
  CadRow : String;
  Renglon : Variant;
Begin
  CadRow := Trim(IntToStr(Ren)) + ':' + Trim(IntToStr(Ren));
  Renglon := Objeto.Rows[CadRow];
  RowHeight := Renglon.RowHeight;
End;

Procedure SetSelect(Objeto : Variant; Cadena : String);
Var
  Rango : Variant;
Begin
  Objeto.Select;
  Try
    Rango := Objeto.Range[Cadena];
    Rango.Select;
  Except
    showmessage(Cadena);
  End;
End;

Function Letra(Numero:Integer) : String;
Var
  Paso1, Paso2 : Integer;
  Cadena : String;
Begin
  Dec(Numero);
  Paso1 := Trunc(Numero / 26);
  Paso2 := Numero - (paso1 * 26);
  if paso1 > 0 then Cadena := Chr(Paso1 + 64);
  Cadena := Cadena + chr(Paso2 + 65);
  Letra := Cadena;
End;

Procedure SetColumnWidth(Objeto : Variant; Col1 : Integer; Col2 : Integer; Ancho : Real;isPixel:Boolean=False);
Var
  CadCol : String;
  Columna : Variant;
  i:Integer;
Begin

  if isPixel then
  begin
    for I := Col1 to Col2 do
      Objeto.Columns[Letra(Col1) + ':' + Letra(Col1)].Width:=Ancho;
   // Columna.Width:= Ancho
  end
  else
  begin
    CadCol := Letra(Col1) + ':' + Letra(Col2);
    Columna := Objeto.Columns[CadCol];
    Columna.ColumnWidth := Ancho;
  end;


End;

Procedure TfrmGraficaGerencialDX.EncabezadoDetalle(oGantt: Variant; oHoja: Variant);
Var
  myCol: Variant;
  myRow: Variant;
  myTiempo: Real;
  Bloque: Extended;
  mFecha: TDateTime;
  OldFecha: Integer;
  OldCol,AuxCol,Hrs: Integer;
  sValor: String;
  Condision: Boolean;
  Cadena{, tmpNombre}: String;
  fs: TStream;
  TempPath: array [0..MAX_PATH-1] of Char;
  imgAux:TImage;
  Pic : TJpegImage;
  QrConfiguracion : TZReadOnlyQuery ;
  sMes:string;
  iIniciaMes:Integer;
function GetTempFile(const Extension: string): string;
var
  Buffer: array[0..MAX_PATH] of Char;
begin
  Repeat
    GetTempPath(SizeOf(Buffer) - 1, Buffer);
    GetTempFileName(Buffer, '~', 0, Buffer);
    Result := ChangeFileExt(Buffer, Extension);
  until not FileExists(Result);
end;

Begin
  QrConfiguracion := TZReadOnlyQuery.Create(Nil);
  QrConfiguracion.Connection := Connection.zConnection;
  QrConfiguracion.SQL.Text:='select c.sMostrarAvances,c.iFirmas, c.sOrdenPerEq, c.sTipoPartida, c.sImprimePEP,ot.sIdplataforma as localizacion, ' +
      ' (select sContrato from contratos where sContrato=:ContratoBarco and sTipoObra = "BARCO" ) as sContratoBarco, ' +
      ' (select mDescripcion from contratos where sContrato=:ContratoBarco and sTipoObra = "BARCO" ) as mDescripcionBarco, ' +
      'c.sClaveSeguridad, c.cStatusProceso, c.sOrdenExtraordinaria, c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3,' +
      'ot.bAvanceFrente, ot.bAvanceContrato, ot.bComentarios, ot.bPermisos, ot.lMostrarAvanceProgramado, ot.lImprimePersonalTM, ot.lPersonalxPartida, ' +
      'c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c2.sProrrateoBarco, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, c.lImprimeExtraordinario, ' +
      'c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, ot.lImprimeFases, cv.dFechaInicio, cv.dfechaFinal, ' +
      'ot.mdescripcion as DescFolio, ot.sNumeroOrden From contratos c2 INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
      'inner join ordenesdetrabajo ot on (ot.sContrato = c2.sContrato and ot.sNumeroOrden =:Orden ) ' +
      'inner join convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
      'Where c2.sContrato = :Contrato';
  QrConfiguracion.ParamByName('contrato').AsString:= global_contrato;
  QrConfiguracion.ParamByName('convenio').AsString:= global_convenio;
  QrConfiguracion.ParamByName('Orden').AsString:= tsNumeroOrden.KeyValue;
  QrConfiguracion.ParamByName('ContratoBarco').AsString:= global_Contrato_barco;
  //QrConfiguracion.ParamByName('Fecha').AsDate:= ReporteDiario.FieldByName('dIdFecha').AsDateTime;
  QrConfiguracion.Open;


  imgAux:=TImage.Create(nil);
  oGantt.ActiveWindow.Zoom := 10;
  SetColumnWidth(oGantt,2,2,231.78);  //Columns['B:B'].ColumnWidth := 50;
  SetColumnWidth(oGantt,rAncho,rAncho,0.5);

  // Dar formato general a las columnas
  myCol := oHoja.Columns['A:A'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '@';
  oGantt.Selection.Font.Name:='Arial';
  oGantt.Selection.Font.Size := 48;
  oGantt.Selection.ColumnWidth := 18.67;

  myCol := oHoja.Columns['B:B'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '@';
  oGantt.Selection.Font.Name:='Arial';
  oGantt.Selection.Font.Size := 48;


  myCol := oHoja.Columns['C:C'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlGeneral;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.Font.Name:='Arial';
  oGantt.Selection.Font.Size := 48;
  //oGantt.Selection.ColumnWidth := 10.71;

  myCol := oHoja.Columns['D:D'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.Font.Name:='Arial';
  oGantt.Selection.Font.Size := 48;
  oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';

  myCol := oHoja.Columns['E:E'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.WrapText := True;
  //oGantt.Selection.ColumnWidth := 12;
  oGantt.Selection.Font.Name:='Arial';
  oGantt.Selection.Font.Size := 48;
  SetColumnWidth(oGantt,3,6,83.78);

  myCol := oHoja.Columns['F:F'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;

  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Font.Name:='Arial';
  oGantt.Selection.Font.Size := 48;
  oGantt.Selection.ColumnWidth := 12;

  SetRowHeight(oGantt,10,13.8);
  myRow := oGantt.Rows['10:10'];
  myRow.WrapText := True;

  myRow := oGantt.Rows['10:10'];
  myRow.VerticalAlignment := xlTop;

  SetSelect(oHoja,'A10:A14');
  oGantt.Selection.Value := 'No.';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  {SetSelect(oHoja,'B09:B11');
  oGantt.Selection.Value := 'ANEXO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter; }

  SetSelect(oHoja,'B10:B14');
  oGantt.Selection.Value := 'ACTIVIDADES';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'C10:C14');
  oGantt.Selection.Value := 'DURACI�N';
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;


  SetSelect(oHoja,'D10:D14');
  oGantt.Selection.Value := 'FECHA INICIO';
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'E10:E14');
  oGantt.Selection.Value := 'FECHA FINAL';
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'F10:F12');
  oGantt.Selection.Value := 'M';
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.Font.size:=40;
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'F13:F13');
  oGantt.Selection.Value := 'D';
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.Font.size:=40;
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

    SetSelect(oHoja,'F14:F14');
  oGantt.Selection.Value := 'H';
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.Font.size:=40;
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  // Determinar la altura total de los titulos
  myRow := oHoja.Rows['1:1'];
  myRow.RowHeight := 20.3;
  myRow := oHoja.Rows['2:3'];
  myRow.RowHeight := 58.5;
  myRow := oHoja.Rows['4:4'];
  myRow.RowHeight := 61.5;
  myRow := oHoja.Rows['5:5'];
  myRow.RowHeight := 87;
  myRow := oHoja.Rows['6:6'];
  myRow.RowHeight := 217.5;
  myRow := oHoja.Rows['7:7'];
  myRow.RowHeight := 111;
  myRow := oHoja.Rows['8:8'];
  myRow.RowHeight := 114;
  myRow := oHoja.Rows['9:9'];
  myRow.RowHeight :=30; //114;

  // Agregar nota de tipo de grafica seleccionada
  Cadena := '';
  if AvPaq.ItemIndex = 1 then
  Begin
    Cadena := 'Nota: Los avances de los paquetes de actividades presentados corresponden exclusivamente a los avances de las actividades aqu� graficadas y no necesariamente corresponden a los avances generales.';
    SetColumnWidth(oGantt, 12, 12, 0.0);
    SetColumnWidth(oGantt, 14, 14, 0.0);
  End;

  // AGregar el mensaje correspondiente de acuerdo a los registros
  if Cadena <> '' then Cadena := Cadena + #10;

  case RadioGroup1.ItemIndex of
    0: Cadena := Cadena + '(Gr�fica generada incluyendo barras de vol�menes excedentes)';
    1: Cadena := Cadena + '(Gr�fica generada sin incluir barras de vol�menes excedentes)';
  end;
  {SetSelect(oHoja,'A8:U8');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.Font.Size := 10;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.NumberFormat := '0.00';
  oGantt.Selection.Value := Cadena;      }

  // Ajustar periodo de inicio al bloque completo
  myTiempo := 1440 / pFactor / 24 / 60;     // Convertir D�as a formato de tiempo de excel
  nCol := Trunc(pTerminoP - pInicioP) * pFactor;

  // Establecer las fechas en titulos de barras
  aCol := 0;
  mFecha := pInicioP;
  OldFecha := 0;
  OldCol := 0;
  SaltoCol:=0;
  sMes:='';
  iIniciaMes:=rAncho+1;
  // Rutina de corte para linea dos
  while mFecha < pTerminoP do
  begin
    if myTiempo > 1 then
      Condision := (Month(OldFecha) <> Month(mFecha)) Or ((mFecha + myTiempo) > pTerminoP)
    Else
      Condision := (OldFecha <> Trunc(mFecha)) or ((mFecha + myTiempo) > pTerminoP);

    if sMes<>FormatDateTime('mmmm',mFecha) then
    begin
      if sMes<>'' then
      begin
        sValor := Letra(iIniciaMes) + '10:' + Letra(rAncho + SaltoCol-1) + '11';
        SetSelect(oHoja,sValor);
        oGantt.Selection.Value := UpperCase(sMes);
        oGantt.Selection.Font.bold:=true;
        oGantt.Selection.VerticalAlignment := xlCenter;
        oGantt.Selection.HorizontalAlignment := xlCenter;
        oGantt.Selection.MergeCells := True;
        oGantt.Selection.Font.Name:='Arial';
        oGantt.Selection.Interior.Color:=10092543;
        oGantt.Selection.Font.Size := 35;
      end;
      sMes:=FormatDateTime('mmmm',mFecha);
      if SaltoCol=0 then
        iIniciaMes:=1 + rAncho
      else
        iIniciaMes:=SaltoCol + rAncho;
    end;

    if Condision then
    begin
      if OldFecha <> 0 then
      begin
        // Cerrar el bloque anterior
        sValor := Letra(rAncho + OldCol) + '12:' + Letra(rAncho + SaltoCol {aCol}) + '12';
        SetSelect(oHoja,sValor);
        oGantt.Selection.MergeCells := True;

        AuxCol:=OldCol;
        Hrs:=6;
        while AuxCol<=SaltoCol do
        begin
          sValor := Letra(rAncho + AuxCol) + '14:' + Letra(rAncho + AuxCol {aCol}) + '14';
          SetSelect(oHoja,sValor);
          //oGantt.Selection.NumberFormat :='@';
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.Font.Name:='Arial';
          oGantt.Selection.Font.bold:=true;
          oGantt.Selection.Font.Size := 32;
          oGantt.Selection.Interior.Color:= 10079487;
          oGantt.Selection.Value :=IntToStr(Hrs);
          oGantt.Selection.Orientation := 90;
          Inc(Hrs,6);
          inc(AuxCol);
        end;


        if {aCol}SaltoCol = OldCol then
        Begin
          // Se trata de una misma columna
          if AnchoColumasGrafica = OrAnchoColumnasGrafica then
          Begin
            oGantt.Selection.Font.bold:=true;
             oGantt.Selection.Font.Name:='Arial';
              oGantt.Selection.Font.Size := 35;
            if myTiempo <= 1 then
              oGantt.Selection.NumberFormat := 'dd/mm';
          End;
        End;
      end;
      OldCol :=SaltoCol + 1;// aCol + 1;     /// Aqui Puede ser el salto
      //SaltoCol:=SaltoCol + 1;
    end;

    sValor := Letra(rAncho + SaltoCol + 1) + '12:' + Letra(rAncho + SaltoCol + 4) + '12';
    SetSelect(oHoja,sValor);
    oGantt.Selection.Value := mFecha;
    oGantt.Selection.Font.bold:=true;
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.MergeCells := True;
     oGantt.Selection.Font.Name:='Arial';
    oGantt.Selection.Font.Size := 35;
    if myTiempo > 1 then
      oGantt.Selection.NumberFormat := 'mmmm/aa'
    Else
      oGantt.Selection.NumberFormat := 'dd/mm/aa';

    // Rutina de tercer valor
    sValor := Letra(rAncho + SaltoCol + 1) + '13:' + Letra(rAncho + SaltoCol + 4) + '13';
    SetSelect(oHoja,sValor);
    oGantt.Selection.Font.bold:=true;
    oGantt.Selection.Font.Name:='Arial';
    oGantt.Selection.Font.Size := 32;

    if myTiempo < 1 then
    Begin
      // Incremento de tiempo inferior a un d�a, colocar los valores correspondientes al horario en tercer rengl�n
{      if CheckBox2.Checked then
        oGantt.Selection.Value := ((mFecha - Trunc(mFecha)) + myTiempo) * 24;}
    End
    Else
    Begin
      if myTiempo > 1 then
      begin
        // Incremento de tiempo superior a un d�a, colocar los valores correspondientes a la fecha en tercer rengl�n
        oGantt.Selection.Font.bold:=true;
        oGantt.Selection.NumberFormat := 'dd/mm';
        oGantt.Selection.Value := mFecha;
      end
      else
      begin
        // Incremento de tiempo igual a un d�a, colocar contador de d�as en tercer rengl�n
        oGantt.Selection.HorizontalAlignment := xlCenter;
        oGantt.Selection.VerticalAlignment := xlCenter;
        oGantt.Selection.Font.bold:=true;
        oGantt.Selection.Value := aCol + 1;
        oGantt.Selection.Interior.Color:= 16777164;//colorIndex := 6;
        oGantt.Selection.Interior.Pattern := xlSolid;

      end;
    End;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.HorizontalAlignment := xlCenter;

    //oGantt.Selection.Font.Size := 8;
















    OldFecha := Trunc(mFecha);
    //OldCol := aCol;
    mFecha := mFecha + myTiempo;
    Inc(aCol);
    inc(SaltoCol,4);
  end;


  sValor := Letra(iIniciaMes) + '10:' + Letra(rAncho + SaltoCol) + '11';
  SetSelect(oHoja,sValor);
  oGantt.Selection.Value := UpperCase(sMes);
  oGantt.Selection.Font.bold:=true;
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.Interior.Color:=10092543;
  oGantt.Selection.Font.Name:='Arial';
  oGantt.Selection.Font.Size := 35;


  sValor := Letra(rAncho + OldCol) + '12:' + Letra(rAncho + SaltoCol{aCol}) + '12';
  SetSelect(oHoja,sValor);
  oGantt.Selection.MergeCells := True;

  AuxCol:=OldCol;
  Hrs:=6;
  while AuxCol<=SaltoCol do
  begin
    sValor := Letra(rAncho + AuxCol) + '14:' + Letra(rAncho + AuxCol {aCol}) + '14';
    SetSelect(oHoja,sValor);
    //oGantt.Selection.NumberFormat :='@';
    oGantt.Selection.Font.Name:='Arial';
    oGantt.Selection.Interior.Color:= 10079487;
    oGantt.Selection.Font.Size := 32;
    oGantt.Selection.Font.bold:=true;
    oGantt.Selection.VerticalAlignment := xlCenter;
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.Value :=IntToStr(Hrs);
    oGantt.Selection.Orientation := 90;
    Inc(Hrs,6);
    inc(AuxCol);
  end;


  // Asegurar la verificaci�n del fin de calendario
  oGantt.Range[Letra(rAncho + SaltoCol{aCol} + 1) + '12:' + Letra(rAncho + {aCol}SaltoCol + 4) + '14'].Select;
  oGantt.Selection.Value := mFecha;

  if OldCol - 1 = {aCol}SaltoCol then
  Begin
    // Se trata de la misma columna, cambiar el formato
    sValor := Letra(rAncho + {aCol}SaltoCol) + '12:' + Letra(rAncho + {aCol}SaltoCol+4) + '12';
    SetSelect(oHoja,sValor);
     oGantt.Selection.Font.Name:='Arial';
    oGantt.Selection.Font.Size := 35;
    oGantt.Selection.NumberFormat := 'dd/mm';
  End;

  // Establecer el ancho de columnas
  SetColumnWidth(oGantt,rAncho + 1,rAncho + {aCol}SaltoCol,28.5{AnchoColumasGrafica});
  myCol := oGantt.Columns[Letra(rAncho + 1)]; // + Letra(aCol + rAncho)];
  myCol.Select;
  AnchoColumnasPixels := myCol.Width;

  wCol :=AnchoColumnasPixels * (SaltoCol+1);   //myCol.Width * (SaltoCol+1);//aCol;
  //Dec(aCol);
  //dec(SaltoCol);

  RangoInicio := pInicioP;
  BarraRango := myTiempo;

  // Obtener los datos de la empresa
  {if Connection.configuracion.RecordCount > 0 then
  Begin
    if aColSaltoCol > 8 then
      SetSelect(oHoja,'C1:' + Letra(rAncho + SaltoColaCol - 6) + '2')
    else
      SetSelect(oHoja,'D1:J1');

    oGantt.Selection.HorizontalAlignment := xlLeft;
    oGantt.Selection.Font.Size := 24;
    oGantt.Selection.Font.Bold := True;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.WrapText := False;
    oGantt.Selection.Value := Connection.Configuracion.FieldByName('sNombreCorto').AsString;
  End;}

  // Generar los encabezados de la gr�fica
  if {aCol}SaltoCol > 8 then
    SetSelect(oHoja,'C2:' + Letra(rAncho + {aCol}SaltoCol) + '3')
  else
    SetSelect(oHoja,'D2:J6');

  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 62;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := false;

  Cadena := Connection.contrato.FieldByName('mCliente').AsString + chr(10) + chr(10);
  Cadena := Cadena + Connection.contrato.FieldByName('sContrato').AsString + '  -  ' + Connection.contrato.FieldByName('mDescripcion').AsString + chr(10) + chr(10);
  Cadena := QuitarEnter(Cadena);

  //SetSelect(oHoja,'D2');
  oGantt.Selection.Value := QrConfiguracion.FieldByName('mDescripcionBarco').AsString +' CONTRATO No. ' + QrConfiguracion.FieldByName('sContratoBarco').AsString;
  //oGantt.Columns['C']
  oHoja.Range['C2'].Select;
  oHoja.Shapes.AddShape(5,oGantt.Selection.left,oGantt.Selection.top,(oGantt.Columns['C'].Width*3) + oGantt.Columns['F'].Width + (oGantt.Columns['G'].Width*Saltocol)  , oHoja.Rows['2:2'].RowHeight + oHoja.Rows['3:3'].RowHeight).Select;
  oGantt.Selection.ShapeRange.Fill.Visible := msoFalse;

  oGantt.Selection.ShapeRange.Line.Visible := msotrue;
  oGantt.Selection.ShapeRange.Line.Weight := 1;

  { myRow := oHoja.Rows['2:3'];
  myRow.RowHeight := 58.5;}

 //Selection.ShapeRange.Fill.Visible = msoFalse







   if {aCol}SaltoCol > 8 then
    SetSelect(oHoja,'C5:C5')
  else
    SetSelect(oHoja,'D2:J6');

  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value :='LOCALIZACI�N:' ;

  SetSelect(oHoja,'D5:' + Letra(rAncho + {aCol}SaltoCol) + '5');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  //oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value := QrConfiguracion.FieldByName('localizacion').AsString;

  SetSelect(oHoja,'C6:C6');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value :='PROGRAMA:' ;

  SetSelect(oHoja,'D6:' + Letra(rAncho + {aCol}SaltoCol) + '6');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value := QrConfiguracion.FieldByName('mdescripcion').AsString;

    SetSelect(oHoja,'C7:C7');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value :='FOLIO:' ;

  SetSelect(oHoja,'D7:' + Letra(rAncho + {aCol}SaltoCol) + '7');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value := QrConfiguracion.FieldByName('sNumeroOrden').AsString;

  SetSelect(oHoja,'C8:C8');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value :='PERIODO:' ;

  SetSelect(oHoja,'D8:' + Letra(rAncho + {aCol}SaltoCol) + '8');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 50;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value := 'DEL ' + IntToStr(DayOfTheMonth(pInicioP)) + ' DE ' + nMes[MonthOfTheYear(pInicioP)] + ' DE ' + FormatDateTime('yyyy',pInicioP) + ' AL ' + IntToStr(DayOfTheMonth(pTerminoP - 1)) + ' DE ' + nMes[MonthOfTheYear(pTerminoP - 1)] + ' DE ' + FormatDateTime('yyyy',pTerminoP - 1);

    oHoja.Range['C5'].Select;
  oHoja.Shapes.AddShape(5,oGantt.Selection.left,oGantt.Selection.top,(oGantt.Columns['C'].Width*3) + oGantt.Columns['F'].Width + (oGantt.Columns['G'].Width*Saltocol)  , oHoja.Rows['5:5'].RowHeight + oHoja.Rows['6:6'].RowHeight+ oHoja.Rows['7:7'].RowHeight+ + oHoja.Rows['8:8'].RowHeight + (oHoja.Rows['9:9'].RowHeight/2)).Select;
  oGantt.Selection.ShapeRange.Fill.Visible := msoFalse;

  oGantt.Selection.ShapeRange.Line.Visible := msotrue;
  oGantt.Selection.ShapeRange.Line.Weight := 1;
  oGantt.Selection.ShapeRange.Incrementtop(((oHoja.Rows['4:4'].RowHeight/2) * -1));






  {Case Opciones.ItemIndex of
    0: Cadena := 'DIAGRAMA DE GANTT GERENCIAL PERIODO DE GRAFICACI�N DEL ' + IntToStr(DayOfTheMonth(pInicioP)) + ' DE ' + nMes[MonthOfTheYear(pInicioP)] + ' DE ' + FormatDateTime('yyyy',pInicioP) + ' AL ' + IntToStr(DayOfTheMonth(pTerminoP - 1)) + ' DE ' + nMes[MonthOfTheYear(pTerminoP - 1)] + ' DE ' + FormatDateTime('yyyy',pTerminoP - 1);
    1: Cadena := 'DIAGRAMA DE GANTT ORDEN DE TRABAJO: ' + Ordenesdetrabajo.FieldValues['sNumeroOrden'] + #13 + ' PERIODO DE GRAFICACI�N DEL ' + IntToStr(DayOfTheMonth(pInicioP)) + ' DE ' + nMes[MonthOfTheYear(pInicioP)] + ' DE ' + FormatDateTime('yyyy',pInicioP) + ' AL ' + IntToStr(DayOfTheMonth(pTerminoP - 1)) + ' DE ' + nMes[MonthOfTheYear(pTerminoP - 1)] + ' DE ' + FormatDateTime('yyyy',pTerminoP - 1);
  End;
  Cadena := QuitarEnter(Cadena);

  //Cadena := 'DIAGRAMA DE GANTT DE SEGUIMIENTO';

  SetSelect(oHoja,'B7:' + ColumnaNombre(rAncho+1)+'7');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 12;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value := Cadena; }



  // Obtener la imagen del cliente desde la base de datos
  if tmpNombre='' then
  begin

    //tmpNombreC := GetTempFile('.~im');
    GetTempPath(SizeOf(TempPath), TempPath);
    tmpNombre:=TempPath +'imgtempMar'+formatdatetime('dddddd hhnnss',now)+'.jpg';

    fs := Connection.Contrato.CreateBlobStream(Connection.Contrato.FieldByName('bImagen'), bmRead);
    If fs.Size > 1 Then
    Begin
      try
        Pic:=TJpegImage.Create;
        try
           Pic.LoadFromStream(fs);
           imgAux.Picture.Graphic := Pic;
        finally
           Pic.Free;
        end;
      finally
          fs.Free;
      End;
      imgAux.Picture.SaveToFile(tmpNombre);
    End;


  end;

  if FileExists(tmpNombre) then
  begin

    // Agregar Imagen de la empresa
    oHoja.Cells[1,1].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombre).Select;
    // Determinar el tama�o real de la imagen
    //Bloque := (oGantt.Rows[1].Height + oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height) + 20;   // * 0.7;
    //oGantt.Selection.ShapeRange.ScaleWidth(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromBottomRight);
    //oGantt.Selection.ShapeRange.ScaleHeight(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);
    oGantt.Selection.ShapeRange.Height:=oHoja.Rows['1:1'].RowHeight+ oHoja.Rows['2:2'].RowHeight+oHoja.Rows['3:3'].RowHeight+oHoja.Rows['4:4'].RowHeight+ oHoja.Rows['5:5'].RowHeight + oHoja.Rows['6:6'].RowHeight+ oHoja.Rows['7:7'].RowHeight+ + oHoja.Rows['8:8'].RowHeight + oHoja.Rows['9:9'].RowHeight
    + oHoja.Rows['10:10'].RowHeight+ oHoja.Rows['11:11'].RowHeight;
    oGantt.Selection.ShapeRange.Width:=oGantt.Columns['A'].Width + oGantt.Columns['B'].Width -30;
    oGantt.Selection.ShapeRange.Left := 24;  //Margen;    //Excel.Columns['A:A'].Width + Margen;
    oGantt.Selection.ShapeRange.Top := 0;

    oGantt.ActiveSheet.Shapes.AddPicture(tmpNombre,false,True,oGantt.Selection.ShapeRange.Left,oGantt.Selection.ShapeRange.Top,oGantt.Selection.ShapeRange.Width,oGantt.Selection.ShapeRange.Height);
    oGantt.Selection.ShapeRange.delete;


  end;

   if tmpNombreC='' then
  begin
    GetTempPath(SizeOf(TempPath), TempPath);
    tmpNombreC:=TempPath +'imgtempMar2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
    fs :=Connection.configuracion.CreateBlobStream(Connection.configuracion.FieldByName('bImagen'), bmRead);

    If fs.Size > 1 Then
    Begin
      try
        Pic:=TJpegImage.Create;
        try
           Pic.LoadFromStream(fs);
           imgAux.Picture.Graphic := Pic;
        finally
           Pic.Free;
        end;
      finally
          fs.Free;
      End;
      imgAux.Picture.SaveToFile(tmpNombreC);
    End;

  end;


   if FileExists(tmpNombreC) then
  begin


    // Obtener la imagen de la empresa desde la base de datos

    // Agregar Imagen Cliente a la hoja de excel
   // oHoja.Cells[1,1].Select;
    oHoja.Cells[1,rAncho + {aCol}saltocol + 1].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombreC).Select;


    // Determinar el tama�o real de la imagen
    //Bloque := (oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height);   // * 0.7;
    //oGantt.Selection.ShapeRange.ScaleWidth(oGantt.Columns[ColumnaNombre(rAncho +saltocol+1)].Width + oGantt.Columns[ColumnaNombre(rAncho +saltocol+2)].Width + oGantt.Columns[ColumnaNombre(rAncho +saltocol+3)].Width, msoFalse, msoScaleFromTopLeft); //msoScaleFromBottomRight);
    //oGantt.Selection.ShapeRange.ScaleHeight(oHoja.Rows['1:1'].RowHeight+ oHoja.Rows['2:2'].RowHeight+oHoja.Rows['3:3'].RowHeight+oHoja.Rows['4:4'].RowHeight+ oHoja.Rows['5:5'].RowHeight + oHoja.Rows['6:6'].RowHeight+ oHoja.Rows['7:7'].RowHeight+ + oHoja.Rows['8:8'].RowHeight + oHoja.Rows['9:9'].RowHeight
   // + oHoja.Rows['10:10'].RowHeight+ oHoja.Rows['11:11'].RowHeight, msoFalse, msoScaleFromTopLeft);

   // oGantt.Selection.ShapeRange.Width:=(40);
  //  oGantt.Selection.ShapeRange.Height:=oHoja.Rows['1:1'].RowHeight+ oHoja.Rows['2:2'].RowHeight+oHoja.Rows['3:3'].RowHeight+oHoja.Rows['4:4'].RowHeight+ oHoja.Rows['5:5'].RowHeight + oHoja.Rows['6:6'].RowHeight+ oHoja.Rows['7:7'].RowHeight+ + oHoja.Rows['8:8'].RowHeight + oHoja.Rows['9:9'].RowHeight
  //  + oHoja.Rows['10:10'].RowHeight;
   oGantt.Selection.ShapeRange.Width:=(oHoja.Rows['1:1'].RowHeight+ oHoja.Rows['2:2'].RowHeight+oHoja.Rows['3:3'].RowHeight+oHoja.Rows['4:4'].RowHeight+ oHoja.Rows['5:5'].RowHeight + oHoja.Rows['6:6'].RowHeight+ oHoja.Rows['7:7'].RowHeight+ + oHoja.Rows['8:8'].RowHeight + oHoja.Rows['9:9'].RowHeight
   + oHoja.Rows['10:10'].RowHeight + oHoja.Rows['11:11'].RowHeight);

    //oGantt.Selection.ShapeRange.IncrementLeft((oGantt.Selection.ShapeRange.Width + 0.5) * -1);
    oGantt.Selection.ShapeRange.Top := 40;
    oGantt.Selection.ShapeRange.Left:=oGantt.Selection.ShapeRange.Left + 10;
    oGantt.ActiveSheet.Shapes.AddPicture(tmpNombreC,false,True,oGantt.Selection.ShapeRange.Left,oGantt.Selection.ShapeRange.Top,oGantt.Selection.ShapeRange.Width,oGantt.Selection.ShapeRange.Height);
    oGantt.Selection.ShapeRange.delete;
  end;

  // Eliminar el archivo temporal
 {Try
    DeleteFile(tmpNombre);
  Finally
    tmpNombre := '';
  End; } 
End;


Procedure TfrmGraficaGerencialDX.Encabezado(oGantt: Variant; oHoja: Variant);
Var
  myCol: Variant;
  myRow: Variant;
  myTiempo: Real;
  Bloque: Extended;
  mFecha: TDateTime;
  OldFecha: Integer;
  OldCol: Integer;
  sValor: String;
  Condision: Boolean;
  Cadena, tmpNombre: String;
  fs: TStream;

function GetTempFile(const Extension: string): string;
var
  Buffer: array[0..MAX_PATH] of Char;
begin
  Repeat
    GetTempPath(SizeOf(Buffer) - 1, Buffer);
    GetTempFileName(Buffer, '~', 0, Buffer);
    Result := ChangeFileExt(Buffer, Extension);
  until not FileExists(Result);
end;

Begin
  oGantt.ActiveWindow.Zoom := 75;
  SetColumnWidth(oGantt,3,3,50);  //Columns['B:B'].ColumnWidth := 50;
  SetColumnWidth(oGantt,rAncho,rAncho,0.5);

  // Dar formato general a las columnas
  myCol := oHoja.Columns['A:A'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '@';
  oGantt.Selection.Font.Size := 8;
  oGantt.Selection.ColumnWidth := 10.71;

  myCol := oHoja.Columns['B:B'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '@';
  oGantt.Selection.Font.Size := 8;
  oGantt.Selection.ColumnWidth := 10.71;

  myCol := oHoja.Columns['C:C'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlGeneral;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.Font.Size := 8;

  myCol := oHoja.Columns['D:J'];
  myCol.Select;
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.VerticalAlignment := xlTop;

  myCol := oHoja.Columns['F:F'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.0000%';

  myCol := oHoja.Columns['G:G'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.0000';
  oGantt.Selection.ColumnWidth := 0;

  myCol := oHoja.Columns['H:H'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.0000';
  oGantt.Selection.ColumnWidth := 0;

  myCol := oHoja.Columns['I:I'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.0000%';

  myCol := oHoja.Columns['J:J'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.00';
  oGantt.Selection.ColumnWidth := 0;

  myCol := oHoja.Columns['K:K'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.00';
  oGantt.Selection.ColumnWidth := 0;

  myCol := oHoja.Columns['L:L'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.0000%';

  myCol := oHoja.Columns['M:M'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.0000%';

  myCol := oHoja.Columns['N:N'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := '0.0000%';

  myCol := oHoja.Columns['O:O'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.ColumnWidth := 12;

  myCol := oHoja.Columns['P:P'];
  myCol.Select;
  oGantt.Selection.VerticalAlignment := xlTop;
  oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';
  oGantt.Selection.HorizontalAlignment := xlCenter;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.ColumnWidth := 12;

  SetRowHeight(oGantt,10,13.8);
  myRow := oGantt.Rows['10:10'];
  myRow.WrapText := True;

  myRow := oGantt.Rows['10:10'];
  myRow.VerticalAlignment := xlTop;

  SetSelect(oHoja,'A09:A11');
  oGantt.Selection.Value := 'PARTIDA';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'B09:B11');
  oGantt.Selection.Value := 'ANEXO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'C09:C11');
  oGantt.Selection.Value := 'DESCRIPCI�N';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'D09:D11');
  oGantt.Selection.Value := 'TIEMPO PROGRAMADO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'E09:E11');
  oGantt.Selection.Value := 'DIAS';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'F09:F11');
  oGantt.Selection.Value := 'POND.';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'G09:G11');
  oGantt.Selection.Value := 'COSTO TOTAL MN';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'H09:H11');
  oGantt.Selection.Value := 'COSTO TOTAL USD';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'I09:I11');
  oGantt.Selection.Value := 'AVANCE PONDERADO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'J09:J11');
  oGantt.Selection.Value := 'COSTO EJECUTADO MN';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'K09:K11');
  oGantt.Selection.Value := 'COSTO EJECUTADO USD';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'L09:L11');
  oGantt.Selection.Value := 'AVANCE ANTERIOR AL PERIODO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'M09:M11');
  oGantt.Selection.Value := 'AVANCE PERIODO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'N09:N11');
  oGantt.Selection.Value := 'AVANCE TOTAL AL PERIODO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'O09:O11');
  oGantt.Selection.Value := 'FECHA PROG. DE INICIO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  SetSelect(oHoja,'P09:P11');
  oGantt.Selection.Value := 'FECHA PROG. DE TERMINO';
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.HorizontalAlignment := xlCenter;

  // Determinar la altura total de los titulos
  myRow := oHoja.Rows['1:6'];
  myRow.RowHeight := 30;
  myRow := oHoja.Rows['7:7'];
  myRow.RowHeight := 40;
  myRow := oHoja.Rows['8:8'];
  myRow.RowHeight := 32;

  // Agregar nota de tipo de grafica seleccionada
  Cadena := '';
  if AvPaq.ItemIndex = 1 then
  Begin
    Cadena := 'Nota: Los avances de los paquetes de actividades presentados corresponden exclusivamente a los avances de las actividades aqu� graficadas y no necesariamente corresponden a los avances generales.';
    SetColumnWidth(oGantt, 12, 12, 0.0);
    SetColumnWidth(oGantt, 14, 14, 0.0);
  End;

  // AGregar el mensaje correspondiente de acuerdo a los registros
  if Cadena <> '' then Cadena := Cadena + #10;

  case RadioGroup1.ItemIndex of
    0: Cadena := Cadena + '(Gr�fica generada incluyendo barras de vol�menes excedentes)';
    1: Cadena := Cadena + '(Gr�fica generada sin incluir barras de vol�menes excedentes)';
  end;
  SetSelect(oHoja,'A8:U8');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.Font.Size := 10;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.NumberFormat := '0.00';
  oGantt.Selection.Value := Cadena;

  // Ajustar periodo de inicio al bloque completo
  myTiempo := 1440 / pFactor / 24 / 60;     // Convertir D�as a formato de tiempo de excel
  nCol := Trunc(pTerminoP - pInicioP) * pFactor;

  // Establecer las fechas en titulos de barras
  aCol := 0;
  mFecha := pInicioP;
  OldFecha := 0;
  OldCol := 0;

  // Rutina de corte para linea dos
  while mFecha < pTerminoP do
  begin
    if myTiempo > 1 then
      Condision := (Month(OldFecha) <> Month(mFecha)) Or ((mFecha + myTiempo) > pTerminoP)
    Else
      Condision := (OldFecha <> Trunc(mFecha)) or ((mFecha + myTiempo) > pTerminoP);

    if Condision then
    begin
      if OldFecha <> 0 then
      begin
        // Cerrar el bloque anterior
        sValor := Letra(rAncho + OldCol) + '10:' + Letra(rAncho + aCol) + '10';
        SetSelect(oHoja,sValor);
        oGantt.Selection.MergeCells := True;
        if aCol = OldCol then
        Begin
          // Se trata de una misma columna
          if AnchoColumasGrafica = OrAnchoColumnasGrafica then
          Begin
            oGantt.Selection.Font.Size := 6;
            if myTiempo <= 1 then
              oGantt.Selection.NumberFormat := 'dd/mm';
          End;
        End;
      end;
      OldCol := aCol + 1;
    end;

    sValor := Letra(rAncho + aCol + 1) + '10:' + Letra(rAncho + aCol + 1) + '10';
    SetSelect(oHoja,sValor);
    oGantt.Selection.Value := mFecha;
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.Font.Size := 8;
    if myTiempo > 1 then
      oGantt.Selection.NumberFormat := 'mmmm/aa'
    Else
      oGantt.Selection.NumberFormat := 'dd/mm/aa';

    // Rutina de tercer valor
    sValor := Letra(rAncho + aCol + 1) + '11:' + Letra(rAncho + aCol + 1) + '11';
    SetSelect(oHoja,sValor);

    if myTiempo < 1 then
    Begin
      // Incremento de tiempo inferior a un d�a, colocar los valores correspondientes al horario en tercer rengl�n
{      if CheckBox2.Checked then
        oGantt.Selection.Value := ((mFecha - Trunc(mFecha)) + myTiempo) * 24;}
    End
    Else
    Begin
      if myTiempo > 1 then
      begin
        // Incremento de tiempo superior a un d�a, colocar los valores correspondientes a la fecha en tercer rengl�n
        oGantt.Selection.NumberFormat := 'dd/mm';
        oGantt.Selection.Value := mFecha;
      end
      else
      begin
        // Incremento de tiempo igual a un d�a, colocar contador de d�as en tercer rengl�n
        oGantt.Selection.Value := aCol + 1;
      end;
    End;

    oGantt.Selection.HorizontalAlignment := xlCenter;

    oGantt.Selection.Font.Size := 8;

    OldFecha := Trunc(mFecha);
    //OldCol := aCol;
    mFecha := mFecha + myTiempo;
    Inc(aCol);
  end;

  sValor := Letra(rAncho + OldCol) + '10:' + Letra(rAncho + aCol) + '10';
  SetSelect(oHoja,sValor);
  oGantt.Selection.MergeCells := True;

  // Asegurar la verificaci�n del fin de calendario
  oGantt.Range[Letra(rAncho + aCol + 1) + '9:' + Letra(rAncho + aCol + 1) + '11'].Select;
  oGantt.Selection.Value := mFecha;

  if OldCol - 1 = aCol then
  Begin
    // Se trata de la misma columna, cambiar el formato
    sValor := Letra(rAncho + aCol) + '10:' + Letra(rAncho + aCol) + '10';
    SetSelect(oHoja,sValor);
    oGantt.Selection.NumberFormat := 'dd/mm';
  End;

  // Establecer el ancho de columnas
  SetColumnWidth(oGantt,rAncho + 1,rAncho + aCol, AnchoColumasGrafica);
  myCol := oGantt.Columns[Letra(rAncho + 1)]; // + Letra(aCol + rAncho)];
  myCol.Select;
  AnchoColumnasPixels := myCol.Width;

 // ColumnWith;


  wCol := myCol.Width * aCol;
  Dec(aCol);
  RangoInicio := pInicioP;
  BarraRango := myTiempo;

  // Obtener los datos de la empresa
  if Connection.configuracion.RecordCount > 0 then
  Begin
    if aCol > 8 then
      SetSelect(oHoja,'D1:' + Letra(rAncho + aCol - 9) + '1')
    else
      SetSelect(oHoja,'D1:J1');

    oGantt.Selection.HorizontalAlignment := xlLeft;
    oGantt.Selection.Font.Size := 24;
    oGantt.Selection.Font.Bold := True;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.WrapText := False;
    oGantt.Selection.Value := Connection.Configuracion.FieldByName('sNombreCorto').AsString;
  End;

  // Generar los encabezados de la gr�fica
  if aCol > 8 then
    SetSelect(oHoja,'D2:' + Letra(rAncho + aCol - 9) + '6')
  else
    SetSelect(oHoja,'D2:J6');

  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.Font.Size := 14;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;

  Cadena := Connection.contrato.FieldByName('mCliente').AsString + chr(10) + chr(10);
  Cadena := Cadena + Connection.contrato.FieldByName('sContrato').AsString + '  -  ' + Connection.contrato.FieldByName('mDescripcion').AsString + chr(10) + chr(10);
  Cadena := QuitarEnter(Cadena);

  SetSelect(oHoja,'D2');
  oGantt.Selection.Value := Cadena;

  Case Opciones.ItemIndex of
    0: Cadena := 'DIAGRAMA DE GANTT GERENCIAL PERIODO DE GRAFICACI�N DEL ' + IntToStr(DayOfTheMonth(pInicioP)) + ' DE ' + nMes[MonthOfTheYear(pInicioP)] + ' DE ' + FormatDateTime('yyyy',pInicioP) + ' AL ' + IntToStr(DayOfTheMonth(pTerminoP - 1)) + ' DE ' + nMes[MonthOfTheYear(pTerminoP - 1)] + ' DE ' + FormatDateTime('yyyy',pTerminoP - 1);
    1: Cadena := 'DIAGRAMA DE GANTT ORDEN DE TRABAJO: ' + Ordenesdetrabajo.FieldValues['sNumeroOrden'] + #13 + ' PERIODO DE GRAFICACI�N DEL ' + IntToStr(DayOfTheMonth(pInicioP)) + ' DE ' + nMes[MonthOfTheYear(pInicioP)] + ' DE ' + FormatDateTime('yyyy',pInicioP) + ' AL ' + IntToStr(DayOfTheMonth(pTerminoP - 1)) + ' DE ' + nMes[MonthOfTheYear(pTerminoP - 1)] + ' DE ' + FormatDateTime('yyyy',pTerminoP - 1);
  End;
  Cadena := QuitarEnter(Cadena);

  //Cadena := 'DIAGRAMA DE GANTT DE SEGUIMIENTO';

  SetSelect(oHoja,'B7:R7');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 12;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := True;
  oGantt.Selection.Value := Cadena;

  // Obtener la imagen del cliente desde la base de datos
  tmpNombre := GetTempFile('.~im');

  if Not Connection.configuracion.FieldByName('bImagen').IsNull then
  begin
    fs := Connection.configuracion.CreateBlobStream(Connection.configuracion.FieldByName('bImagen'), bmRead);
    try
      fs.Seek(0, soFromBeginning);
      with TFileStream.Create(tmpNombre, fmCreate) do
        try
          CopyFrom(fs, fs.Size)
        finally
          Free
        end;
    finally
      fs.Free
    end;

    // Agregar Imagen de la empresa
    oHoja.Cells[1,1].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombre).Select;
    // Determinar el tama�o real de la imagen
    Bloque := (oGantt.Rows[1].Height + oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height) + 20;   // * 0.7;
    oGantt.Selection.ShapeRange.ScaleWidth(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromBottomRight);
    oGantt.Selection.ShapeRange.ScaleHeight(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);

    oGantt.Selection.ShapeRange.Left := 24;  //Margen;    //Excel.Columns['A:A'].Width + Margen;
    oGantt.Selection.ShapeRange.Top := 0;
  end;

  if Not Connection.Contrato.FieldByName('bImagen').IsNull then
  begin
    // Obtener la imagen de la empresa desde la base de datos
    fs := Connection.Contrato.CreateBlobStream(Connection.Contrato.FieldByName('bImagen'), bmRead);
    try
      fs.Seek(0, soFromBeginning);
      with TFileStream.Create(tmpNombre, fmCreate) do
        try
          CopyFrom(fs, fs.Size)
        finally
          Free
        end;
    finally
      fs.Free
    end;
    // Agregar Imagen Cliente a la hoja de excel
    oHoja.Cells[1,1].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombre).Select;

    oGantt.Selection.Cut;
    oHoja.Cells[2,rAncho + aCol + 2].Select;
    oGantt.ActiveSheet.Paste;

    // Determinar el tama�o real de la imagen
    Bloque := (oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height);   // * 0.7;
    oGantt.Selection.ShapeRange.ScaleWidth(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft); //msoScaleFromBottomRight);
    oGantt.Selection.ShapeRange.ScaleHeight(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);

    oGantt.Selection.ShapeRange.IncrementLeft((oGantt.Selection.ShapeRange.Width + 0.5) * -1);
    oGantt.Selection.ShapeRange.Top := oGantt.Rows[1].Height;
  end;

  // Eliminar el archivo temporal
  Try
    DeleteFile(tmpNombre);
  Finally
    tmpNombre := '';
  End;
End;

procedure TfrmGraficaGerencialDX.btnGraphClick(Sender: TObject);
Var
  zrqPrograma, qPaquete0, qBusca,
  qLoc, qAvances, qAnteriorPda,
  zrqponderado, qPaquetes         : TZReadOnlyQuery;
  OpenDialog: TOpenDialog;
  mContrato, Llave, fName, Ext,
  Cadena, mCol, nombretabla,
  nombrecampo, ncantidadanexo,
  nwbs, numeroorden, paquete      : String;
  Resul, frente                   : Boolean;
  nDias, AcumPonde, Alto, hCent,
  ColBarra, Val1, Val2, DurBarra,
  fAlto, tFecha, iFecha, tAvance,
  i2Fecha, t2Fecha, Avance,
  pPonde, myInicio, myTermino,
  AcumAnterior, Original, hCol,
  AvancePond                      : Extended;
  Continua, Pendiente: Boolean;
  Cuenta, mCta, Ren, cNivel,
  cBarras, Cuantos                : Integer;
  oGantt, oLibro, oHoja, myCol    : Variant;
  Marca, sCadQuery0, sCadQuery1,
  sCadQuery2, sCadQuery3,
  sxCadQuery, OldActiv,
  NumberFormat, OldNumero         : String;
  OldIniBar, OldFinBar, gVolumen,
  VolExced,
  Col2Barra, SumCant, Resto,
  FactorExcedente,
  PorcentajeExcedente       : Real;
  OldExced                        : Variant;
  Reintentar, Graficar, Mostrar   : Boolean;
  LastDate                        : TDate;
  LastColor                       : Integer;
  PeriodoFrente                   : TPeriodo;
  LocRen: Byte;
  MryAvances,MryDetallePdas:TRxMemoryData;
  OriFecha:TDate;
  OriClasif,OriHora:string;
  i:Byte;
  iColHr,iSaltos,iAux:Integer;
  fHoraF,fHoraI,fRestaI:Extended;
  QrMovimientos,QrPaquete:TZReadOnlyQuery;
  PosCol:Integer;
  ColENd:Integer;
  dAVance,eAvance:Extended;
  tmpFechaI,TmpFechaT:TDateTime;
  tmpFechaIP,TmpFechaTP:TDateTime;
  sSumaTotalHrs,sSumaTotalHrsP,sTotalP,sTotalR,sAux:string;
  iRenFinaly:Integer;
  QrDetallePdas:TZReadOnlyQuery;
  EAux:Variant;
  iColHrF:Integer;
  AnchoPage,BloquePage:Extended;
    RangoE:Variant;


  //myCol: Variant;
begin
  tmpNombre:='';
  tmpNombreC:='';

  if Opciones.ItemIndex = 0 then
  begin
    if (OrInicio <> dInicio.Date) and (orTermino <> dTermino.Date) and (cbTipoGrafica.ItemIndex = 0) and (messagedlg('El periodo seleccionado para la grafica de GANTT es diferente del periodo programado.' + #13 + #13 +
                 '�Desea seleccionar automaticamente el periodo programado?' + #13 +
                 '(Si elige "No" la gr�fica se generar� con el periodo seleccionado)', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
    begin
      dInicio.Date := OrInicio;
      dTermino.Date := OrTermino;
    end;
  end;

  zrqPrograma := TZReadOnlyQuery.Create(Nil);
  zrqPrograma.Connection := Connection.zConnection;

  QrDetallePdas:=TZReadOnlyQuery.Create(nil);
  QrDetallePdas.Connection:=connection.zConnection;

  QrPaquete:=TZReadOnlyQuery.Create(nil);
  QrPaquete.Connection:=connection.zConnection;

  MryAvances:=TRxMemoryData.Create(nil);
  MryAvances.Active:=true;

  MryDetallePdas:=TRxMemoryData.Create(nil);
  MryDetallePdas.Active:=true;

  if Opciones.ItemIndex = 0 then
  begin
    // Seleccionar los datos de anexo
    //zrqPrograma.Sql.Text := 'select sWbs, sAnexo, sTipoActividad, sNumeroActividad, mDescripcion, dFechaInicio, dFechaFinal, dDuracion, dPonderado, iNivel ' +
    zrqPrograma.Sql.Text := 'select sWbs, sTipoActividad, sNumeroActividad, mDescripcion, dFechaInicio, dFechaFinal, dDuracion, dPonderado, iNivel ' +
                            'from actividadesxanexo where scontrato = :contrato and sidconvenio = :convenio order by iItemOrden';
    zrqPrograma.ParamByName('contrato').AsString := global_contrato;
    zrqPrograma.ParamByName('convenio').AsString := global_convenio;
  end;

  if Opciones.ItemIndex = 1 then
  begin
    // Seleccionar los datos de frente de trabajo
    //zrqPrograma.Sql.Text := 'select sWbs, sAnexo, sTipoActividad, sNumeroActividad, mDescripcion, dFechaInicio, dFechaFinal, dDuracion, dPonderado, iNivel ' +
    zrqPrograma.Sql.Text := 'select sWbs, sTipoActividad, sNumeroActividad, mDescripcion, dFechaInicio, dFechaFinal, dDuracion, dPonderado, iNivel ' +
                            'from actividadesxorden where scontrato = :contrato and sidconvenio = :convenio and snumeroorden = :orden order by iItemOrden';
    zrqPrograma.ParamByName('contrato').AsString := global_contrato;
    zrqPrograma.ParamByName('convenio').AsString := global_convenio;
    zrqPrograma.ParamByName('orden').AsString := OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString;
  end;
  zrqPrograma.Open;

  // Determinar la fecha real de inicio
  RealDates := TPeriodo.Create;
  if Opciones.ItemIndex = 0 then
  begin
    RealDates.Inicio := dInicio.Date;
    RealDates.Termino := dTermino.Date;
  end
  else
  begin
    RealDates.Inicio := tdFechaInicio.Date;
    RealDates.Termino := tdIdFecha.Date;
  end;

  if cbTipoGrafica.ItemIndex<>2 then
    AjustaPeriodoInicio;    // Ajustar la fecha de inicio en base a los parametros indicados por el usuario

  // Preparar repositorio para avances del periodo
  qAvances := TZReadOnlyQuery.Create(Nil);
  qAvances.Connection := Connection.zConnection;

  if Opciones.ItemIndex = 0 then
  begin
    qAvances.SQL.Text := 'select a.dIdFecha, c.dCantidadAnexo as dCantidadTotal, sum(a.dCantidad) as dCantidad, ' +
                         '(ifNull((select sum(dCantidad) from bitacoradeactividades where sContrato = a.sContrato and sNumeroOrden = a.sNumeroOrden and sWbs = a.sWbs and dIdFecha <= a.dIdFecha), 0) - c.dCantidadAnexo) as dExcedente, ' +
                         'sum(a.dAvance * c.dPonderado / 100) * (b.dCantidad / c.dCantidadAnexo) as dAvancePond, SUM(a.dAvance) * (b.dCantidad / c.dCantidadAnexo) as dAvance ' +
                         'from bitacoradeactividades a inner join actividadesxorden b on (b.sContrato = a.sContrato and b.sIdConvenio = :Convenio and b.sNumeroOrden = a.sNumeroOrden and b.sWbs = a.sWbs) ' +
                         'inner join actividadesxanexo c on (c.sContrato = b.sContrato and c.sIdConvenio = b.sIdConvenio and c.sWbs = b.sWbsContrato) ' +
                         'where a.scontrato = :Contrato and b.sWbsContrato = :Wbs and a.dIdFecha >= :Inicio and a.dIdFecha <= :Final group by a.dIdFecha';
  end
  else
  begin
    if cbTipoGrafica.ItemIndex=2 then
    begin                                                                                  //correccion ed para q no sea excedente
      qAvances.SQL.Text :='select a.dIdFecha,a.sIdClasificacion,cast(a.sHoraInicio as TIME) as sHoraInicio,cast(a.sHoraFinal as TIME) as sHoraFinal, b.dCantidad as dCantidadTotal, '+
                          'ifnull(sum(a.dCantidad),0) as dCantidad,(ifNull((select sum(dCantidad) '+
                          'from bitacoradeactividades where sContrato = a.sContrato and sNumeroOrden = a.sNumeroOrden ' +
                          'and sWbs = a.sWbs and dIdFecha <= a.dIdFecha  and sIdTipoMovimiento = "ED"), 0) - b.dCantidad) as dExcedente, '+
                          'sum(a.dAvance * b.dPonderado / 100) as dAvancePond, SUM(a.dAvance) as dAvance,tm2.iColor, '+
                          'a.sHoraInicio as HoraI,a.sHoraFinal as HoraF ' +
                          'from bitacoradeactividades a inner join actividadesxorden b on (b.sContrato = a.sContrato '+
                          'and b.sIdConvenio =:Convenio and b.sNumeroOrden = a.sNumeroOrden and b.sWbs = a.sWbs) '+
                          'inner join tiposdemovimiento tm on (tm.sContrato=:ContratoB and a.sIdTipoMovimiento=tm.sIdTipoMovimiento )'+
                          'inner join tiposdemovimiento tm2 on (tm2.sContrato=:ContratoB and a.sIdClasificacion=tm2.sIdTipoMovimiento )'+
                          'where a.scontrato =:Contrato and a.snumeroorden =:Orden and tm.sClasificacion="Tarifa Diaria" '+
                          'and a.swbs =:wbs and a.dIdFecha between :Inicio and :Final and tm2.lGrafica="Si" ' +
                          'group by a.dIdFecha,a.sHoraInicio,a.sIdClasificacion,a.sHoraFinal';

      qAvances.ParamByName('ContratoB').AsString := global_Contrato_Barco;

    end
    else
    begin
      qAvances.SQL.Text := 'select a.dIdFecha, b.dCantidad as dCantidadTotal, sum(a.dCantidad) as dCantidad, ' +
                           '(ifNull((select sum(dCantidad) from bitacoradeactividades where sContrato = a.sContrato and sNumeroOrden = a.sNumeroOrden and sWbs = a.sWbs and dIdFecha <= a.dIdFecha), 0) - b.dCantidad) as dExcedente, ' +
                           'sum(a.dAvance * b.dPonderado / 100) as dAvancePond, SUM(a.dAvance) as dAvance ' +
                           'from bitacoradeactividades a inner join actividadesxorden b on (b.sContrato = a.sContrato and b.sIdConvenio = :Convenio and b.sNumeroOrden = a.sNumeroOrden and b.sWbs = a.sWbs) ' +
                           'where a.scontrato = :Contrato and a.snumeroorden = :Orden and a.swbs = :Wbs and a.dIdFecha >= :Inicio and a.dIdFecha <= :Final ' +
                           'group by a.dIdFecha';
    end;
    qAvances.ParamByName('orden').AsString := OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString;
  end;

  qAvances.ParamByName('Contrato').AsString := Global_Contrato;
  qAvances.ParamByName('Convenio').AsString := Global_Convenio;
  qAvances.ParamByName('Inicio').AsDate := RealDates.Inicio;
  qAvances.ParamByName('Final').AsDate := RealDates.Termino;

  qAnteriorPda := TZReadOnlyQuery.Create(Nil);
  qAnteriorPda.Connection := Connection.zConnection;
  if Opciones.ItemIndex = 0 then
  begin
    qAnteriorPda.Sql.Text := 'select SUM(a.dAvance) * (b.dCantidad / c.dCantidadAnexo) as dAvanceAnterior ' +
                             'from bitacoradeactividades a ' +
                             'inner join actividadesxorden b on (b.sContrato = a.sContrato and b.sIdConvenio = :Convenio and b.sNumeroOrden = a.sNumeroOrden and b.sWbs = a.sWbs) ' +
                             'inner join actividadesxanexo c on (c.sContrato = b.sContrato and c.sIdConvenio = b.sIdConvenio and c.sWbs = b.sWbsContrato) ' +
                             'where a.scontrato = :Contrato and b.sWbsContrato = :Wbs and a.dIdFecha < :Fecha ' +
                             'group by a.dIdFecha';
  end
  else
  begin
    qAnteriorPda.Sql.Text := 'select sum(a1.dAvance) as dAvanceAnterior from bitacoradeactividades a1 ' +
                             'inner join actividadesxorden b1 on (b1.sContrato = a1.sContrato and b1.sIdConvenio = :Convenio and b1.sNumeroOrden = a1.sNumeroOrden and b1.sWbs = a1.sWbs) ' +
                             'where a1.sContrato = :Contrato and a1.sNumeroOrden = :Orden and a1.sWbs = :Wbs and a1.dIdFecha < :Fecha';
    qAnteriorPda.ParamByName('Orden').AsString := OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString;;
  end;

  qAnteriorPda.ParamByName('Contrato').AsString := global_contrato;
  qAnteriorPda.ParamByName('Convenio').AsString := global_convenio;
  qAnteriorPda.ParamByName('Fecha').AsDate := RealDates.Inicio;

  // Totalizar los avances de los paquetes
  qPaquetes := TZReadOnlyQuery.Create(Nil);
  qPaquetes.Connection := Connection.zConnection;

  if Opciones.ItemIndex = 0 then
  begin
    qPaquetes.Sql.Text := 'select Min(a.dIdFecha) as dInicio, Max(a.dIdFecha) as dFinal, ' +
                          '(sum(if(a.dIdFecha < :Inicio, a.dAvance * c.dPonderado / 100, 0) * (c.dCantidadAnexo / b.dCantidad))) as dAvanceAnterior, ' +
                          '(sum(if(a.dIdFecha between :Inicio and :Final, a.dAvance * c.dPonderado / 100, 0) * (b.dCantidad / c.dCantidadAnexo))) as dAvance ' +
                          'from bitacoradeactividades a ' +
                          'inner join actividadesxorden b on (b.sContrato = a.sContrato and b.sIdConvenio = :Convenio and b.sNumeroOrden = a.sNumeroOrden and b.sWbs = a.sWbs) ' +
                          'inner join actividadesxanexo c on (c.sContrato = b.sContrato and c.sIdConvenio = b.sIdConvenio and c.sWbs = b.sWbsContrato) ' +
                          'where a.scontrato = :Contrato and b.swbscontrato like concat(:Wbs, ".%") and ((:Par=1 and a.dAvance > 0) or :Par=0)';
  end
  else
  begin
    qPaquetes.Sql.Text := 'select Min(a.dIdFecha) as dInicio, Max(a.dIdFecha) as dFinal, ' +
                          'sum(if(a.dIdFecha < :Inicio, a.dAvance * b.dPonderado / 100, 0)) as dAvanceAnterior, ' +
                          'sum(if(a.dIdFecha between :Inicio and :Final, a.dAvance * b.dPonderado / 100, 0)) as dAvance ' +
                          'from bitacoradeactividades a ' +
                          'inner join actividadesxorden b on (b.sContrato = a.sContrato and b.sIdConvenio = :Convenio and b.sNumeroOrden = a.sNumeroOrden and b.sWbs = a.sWbs) ' +
                          'where a.scontrato = :Contrato and a.snumeroorden = :Orden and a.swbs like concat(:Wbs, ".%") and ((:Par=1 and a.dAvance > 0) or :Par=0)';
    qPaquetes.ParamByName('Orden').AsString := OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString;
  end;
  qPaquetes.ParamByName('Contrato').AsString := Global_Contrato;
  qPaquetes.ParamByName('Convenio').AsString := Global_Convenio;
  qPaquetes.ParamByName('Inicio').AsDate := RealDates.Inicio;
  qPaquetes.ParamByName('Final').AsDate := RealDates.Termino;

  // Obtener el avance del paquete de nivel 0 (para hacer coinsidir los avances con los reportados en los reportes diarios)
  qPaquete0 := TZReadOnlyQuery.Create(Nil);
  qPaquete0.Connection := Connection.zConnection;
  qPaquete0.SQL.Text := 'select sum(if(a.dIdFecha < :Inicio, a.dAvance, 0)) as dAvanceAnterior, sum(if(a.dIdFecha >= :Inicio, a.dAvance, 0)) as dAvance ' +
                        'from avancesglobalesxorden a where a.sContrato = :Contrato And a.sIdConvenio = :Convenio And a.sNumeroOrden = :Orden And a.dIdFecha <= :Final';
  qPaquete0.ParamByName('Contrato').AsString := global_contrato;
  qPaquete0.ParamByName('Convenio').AsString := global_convenio;
  if Opciones.ItemIndex = 0 then
    qPaquete0.ParamByName('Orden').AsString := ''
  else
    qPaquete0.ParamByName('Orden').AsString := OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString;
  qPaquete0.ParamByName('Inicio').AsDate := RealDates.Inicio;
  qPaquete0.ParamByName('Final').AsDate := RealDates.Termino;
  qPaquete0.Open;

  // Eliminar los caracteres incorrectos de la candea sContrato
  mContrato := QuitarNovalidos(Global_Contrato);

  // Localizar periodo de generaci�n (Programado)
  zrqPrograma.First;
  Llave := zrqPrograma.FieldByName('sWbs').AsString;

  // Localizar periodo de generacion (Real)
  nDias := 0;
  LastDate := pInicioP;
  while LastDate <= pTerminoP do
  begin
    nDias := nDias + 1;

    case cbTipo.ItemIndex of
      0: LastDate := IncDay(LastDate);
      1: LastDate := IncDay(LastDate, 7);
      2: LastDate := IncMonth(LastDate);
      3: LastDate := IncYear(LastDate);
    end;
  end;

  Lista[01] := 1;             // Cada 24 horas
  Lista[02] := (1 / 7);       // Cada Semana
  Lista[03] := (1 / 31);      // Cada Mes
  Lista[04] := (1 / 366);     // Cada A�o

  // Verificar el formato del calendario
  ItemIndex := cbTipo.ItemIndex;
  Continua := True;
  pFactor := 0;
  while (ItemIndex < 4) And (Continua) do
  Begin
    Inc(ItemIndex);
    if nDias * Lista[ItemIndex] <= MaxBloques then
    begin
      pFactor := Lista[ItemIndex];
      Continua := False;
    end;
  End;

  if Continua then
  Begin
    showmessage('Existe un problema en las fechas de inicio y terminaci�n de actividades, verifique esto e intente de nuevo.');
    Close;
    Exit;
  End;

  cbTipo.ItemIndex := ItemIndex - 1;
  AjustaPeriodoInicio;    // Volver a ajustar el calendario

  // Ajustar el ancho de columnas
  AnchoColumasGrafica := OrAnchoColumnasGrafica;
  case ItemIndex of
    1: AnchoColumasGrafica := OrAnchoColumnasGrafica * 2;
    2: AnchoColumasGrafica := OrAnchoColumnasGrafica * 3;
    3: AnchoColumasGrafica := OrAnchoColumnasGrafica * 4;
    4: AnchoColumasGrafica := OrAnchoColumnasGrafica * 5;
  end;

  //  Establecer por medio del uso de un usuario el nombre y localizaci�n del archivo que se ha de generar
  if cbTipoGrafica.ItemIndex<>2 then
  begin

    OpenDialog := TOpenDialog.Create(self);
    OpenDialog.Filter := 'Archivos de MS EXCEL|*.xls';
    OpenDialog.FilterIndex := 1;

    // Determinar el nombre propuesto para la grafica de gantt
    OpenDialog.FileName := 'Gr�fica de GANTT de seguimiento - ' + mContrato + '.xls';

  // Verificar la extenci�n en el nombre de archivo

    If openDialog.Execute Then
      fName := OpenDialog.FileName
    Else
    Begin
      Close;
      Exit;
    End;

    if Length(fName) > 3 then
    Begin
      Ext := Upcase(fName[Length(fName)- 3]) + Upcase(fName[Length(fName)- 2]) + Upcase(fName[Length(fName)- 1]) + Upcase(fName[Length(fName)]);
      if Ext <> '.XLS' then
        fName := fName + '.xls';
    End
    Else
      fName := fName + '.xls';
  end
  else
  begin
    rAncho:=6;
    rAlto:=14;
     SdgExcel.FileName := 'Gr�fica de GANTT de seguimiento - ' + mContrato + '.xlsx';
    if SdgExcel.Execute then
    begin
      fName := SdgExcel.FileName;
    end
    Else
    Begin
      Close;
      Exit;
    End;

  end;

  // Generar el ambiente de excel para vaciar los datos de la gr�fica
  Try
    oGantt := CreateOleObject('Excel.Application');
  except
    oGantt.quit;
    showmessage('No se pudo crear el objeto EXCEL o se produjo alg�n error.');
    raise;
  End;

  Mostrar := True;
  try
    PnlProcesando.Visible := True;
    try
      oGantt.Visible := False;
      oGantt.DisplayAlerts := False;
      //oGantt.ScreenUpdating := False;
      oLibro := oGantt.Workbooks.Add;    // Crear el libro sobre el que se ha de trabajar
      oHoja := oLibro.Sheets[1];        // Determinar la hoja de trabajo

      // Verificar si la grafica cabe dentro del rango que tenemos en esta version de excel
      //Mostrar := False;
      if cbTipoGrafica.ItemIndex<>2 then
        Encabezado(oGantt,oHoja)   // Colocar el encabezado del diagrama de GANTT
      else
      begin
        EncabezadoDetalle(oGantt,oHoja);
        ColEnd:= SaltoCol+rAncho;
      end;

      oGantt.Visible := CheckBox1.Checked;
      oGantt.ActiveWindow.Zoom := 18;
      // Mostrar los datos del encabezado para que el usuario aprecie la generaci�n
      if CheckBox1.Checked then
      Begin
        oGantt.ScreenUpdating := True;
        //oGantt.ScreenUpdating := False;
      End;

      // Barrer el todos los registros
      ProgressBar1.Min := 0;
      ProgressBar1.Max := zrqPrograma.RecordCount;
      ProgressBar1.Visible := True;
      ProgressBar1.Step := 1;

      zrqPrograma.First;
      Ren := rAlto + 1;       // Aqu� inicia el proceso de datos
      Cuenta := 0;
      cNivel := -1;
      AcumPonde := 0;
      Alto := 0;
      fAlto := 150;//10;
      // Verificar el tipo de grafica que se va a generar para determinar las columnas que se deben incluir
      if cbTipoGrafica.ItemIndex = 0 then
      begin
        // Ocultar las columnas que no se usan para una grafica de gantt
        SetColumnWidth(oGantt,8,8,0);
        SetColumnWidth(oGantt,11,13,0);
      end;

      // Localizar la columna en donde se debe generar la barra
      case ItemIndex of
        1: LocRen := 10;
        2: LocRen := 11;
        3: LocRen := 11;
        4: LocRen := 11;
      end;

      if cbTipoGrafica.ItemIndex=2 then 
        LocRen := 12;

      sTotalR:='00:00';
      sTotalp:='00:00';
      While Cuenta < zrqPrograma.RecordCount do
      Begin
        dAVance:=0;
        // Solo verificar esto si se trata de un gantt de seguimiento
        if (cbTipoGrafica.ItemIndex = 1) or (cbTipoGrafica.ItemIndex=2) then
        begin
          // Abrir datos para verificar si se debe graficar este registro
          if zrqPrograma.FieldByName('sTipoActividad').AsString = 'Paquete' then
          Begin
            oHoja.Rows[IntToStr(ren)+ ':'+IntToStr(ren+3)].RowHeight := 141.75;
            qPaquetes.Close;
            qPaquetes.ParamByName('wbs').AsString := zrqPrograma.FieldByName('sWbs').AsString;
            if RadioGroup1.ItemIndex = 0 then
              qPaquetes.ParamByName('Par').AsInteger := 0
            else
              qPaquetes.ParamByName('Par').AsInteger := 1;
            qPaquetes.Open;

            QrDetallePdas.Active:=False;
            QrDetallePdas.SQL.Text:='select sNumeroOrden,swbs,sNumeroActividad,dFechaInicio,cast(sHoraInicio as Time) as sHoraInicio,dFechaFinal,'+
                                    'cast(sHoraFinal as Time) as sHoraFinal,sHoraInicio as HoraI,sHoraFinal as HoraF ' +
                                    'from actividadesxorden_detalle where sContrato=:Contrato and sIdConvenio=:Convenio' + #10 +
                                    'and sNumeroOrden=:Orden and swbs like :wbs and dFechaInicio Between :Inicio and :Final ' +
                                    'order by dFechaInicio,Time(sHoraInicio)';
            QrDetallePdas.ParamByName('Contrato').AsString := Global_Contrato;
            QrDetallePdas.ParamByName('Convenio').AsString := Global_Convenio;
            QrDetallePdas.ParamByName('Orden').AsString:= tsNumeroOrden.KeyValue;
            QrDetallePdas.ParamByName('wbs').AsString:=zrqPrograma.FieldByName('sWbs').AsString+'%';
            QrDetallePdas.ParamByName('Inicio').AsDate := RealDates.Inicio;
            QrDetallePdas.ParamByName('Final').AsDate := RealDates.Termino;
            QrDetallePdas.Open;

            MryDetallePdas.Active:=False;
            MryDetallePdas.EmptyTable;
            MryDetallePdas.CopyStructure(QrDetallePdas);
            MryDetallePdas.Active:=True;

            OriFecha:=StrToDate('23/02/1984');
            //OriClasif:='ABY';
            OriHora:='SAHL';
            sSumaTotalHrsP:='00:00';
            QrDetallePdas.First;
            while not QrDetallePdas.Eof do
            begin
              if QrDetallePdas.RecNo=1 then
                tmpFechaIP:=StrToDateTime(QrDetallePdas.FieldByName('dFechaInicio').AsString + ' ' + QrDetallePdas.FieldByName('sHoraInicio').AsString)
              else
                if QrDetallePdas.RecNo=QrDetallePdas.RecordCount then
                  TmpFechaTP:=StrToDateTime(QrDetallePdas.FieldByName('dFechaFinal').AsString + ' ' + QrDetallePdas.FieldByName('sHoraFinal').AsString);

              if (QrDetallePdas.FieldByName('dFechaInicio').AsDateTime<>OriFecha) then
              begin
                MryDetallePdas.Append;
                for I := 0 to QrDetallePdas.FieldCount-1 do
                begin
                  MryDetallePdas.Fields[i].ReadOnly:=False;
                  if MryDetallePdas.Fields[i].DataType=ftFloat then
                    MryDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).AsFloat:= QrDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).AsFloat
                  else
                    MryDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).Value:= QrDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).Value;

                end;
                MryDetallePdas.Post;
                OriFecha:=QrDetallePdas.FieldByName('dFechaFinal').AsDateTime;
                //OriClasif:=QrDetallePdas.FieldByName('sIdClasificacion').AsString;
                OriHora:=QrDetallePdas.FieldByName('sHoraFinal').AsString;

                sSumaTotalHrsP:=sfnSumaHoras(sSumaTotalHrsP,sfnRestaHoras(QrDetallePdas.FieldByName('HoraF').AsString,QrDetallePdas.FieldByName('HoraI').AsString));
              end
              else
              begin
                if OriHora=QrDetallePdas.FieldByName('sHoraInicio').AsString then
                begin
                  MryDetallePdas.Edit;
                  MryDetallePdas.FieldByName('sHoraFinal').AsString:=QrDetallePdas.FieldByName('sHoraFinal').AsString;
                  //MryDetallePdas.FieldByName('dAvance').AsFloat:=MryDetallePdas.FieldByName('dAvance').AsFloat + QrDetallePdas.FieldByName('dAvance').AsFloat;
                  MryDetallePdas.Post;
                  OriHora:=QrDetallePdas.FieldByName('sHoraFinal').AsString;
                  sSumaTotalHrsP:=sfnSumaHoras(sSumaTotalHrsP,sfnRestaHoras(QrDetallePdas.FieldByName('HoraF').AsString,QrDetallePdas.FieldByName('HoraI').AsString));

                end
                else
                begin
                  if (QrDetallePdas.FieldByName('sHoraInicio').AsString>=MryDetallePdas.FieldByName('sHoraInicio').AsString)
                  or (QrDetallePdas.FieldByName('sHoraFinal').AsString<=MryDetallePdas.FieldByName('sHoraFinal').AsString) then
                  begin
                     //sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(QrDetallePdas.FieldByName('sHoraFinal').AsString,QrDetallePdas.FieldByName('sHoraInicio').AsString));



                  end
                  else
                  begin
                    MryDetallePdas.Append;
                    for I := 0 to QrDetallePdas.FieldCount-1 do
                    begin
                      MryDetallePdas.Fields[i].ReadOnly:=False;
                      MryDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name     ).Value:= QrDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).Value;
                    end;

                    MryDetallePdas.Post;
                    OriHora:=QrDetallePdas.FieldByName('sHoraFinal').AsString;
                    sSumaTotalHrsP:=sfnSumaHoras(sSumaTotalHrsP,sfnRestaHoras(QrDetallePdas.FieldByName('HoraF').AsString,QrDetallePdas.FieldByName('HoraI').AsString));

                  end;
                end;
              end;
              QrDetallePdas.Next;
            end;


            QrPaquete.Active:=False;
            QrPaquete.SQL.Text:='select a.dIdFecha,a.sIdClasificacion,cast(a.sHoraInicio as TIME) as sHoraInicio,cast(a.sHoraFinal as TIME) as sHoraFinal, b.dCantidad as dCantidadTotal, '+
                                'ifnull(sum(a.dCantidad),0) as dCantidad,(ifNull((select sum(dCantidad) '+
                                'from bitacoradeactividades where sContrato = a.sContrato and sNumeroOrden = a.sNumeroOrden ' +
                                'and sWbs = a.sWbs and dIdFecha <= a.dIdFecha), 0) - b.dCantidad) as dExcedente, '+
                                'sum(a.dAvance * b.dPonderado / 100) as dAvancePond, SUM(a.dAvance) as dAvance,tm2.iColor, '+
                                'a.sHoraInicio as HoraI,a.sHoraFinal as HoraF ' +
                                'from bitacoradeactividades a inner join actividadesxorden b on (b.sContrato = a.sContrato '+
                                'and b.sIdConvenio =:Convenio and b.sNumeroOrden = a.sNumeroOrden and b.sWbs = a.sWbs) '+
                                'inner join tiposdemovimiento tm on (tm.sContrato=:ContratoB and a.sIdTipoMovimiento=tm.sIdTipoMovimiento )'+
                                'inner join tiposdemovimiento tm2 on (tm2.sContrato=:ContratoB and a.sIdClasificacion=tm2.sIdTipoMovimiento )'+
                                'where a.scontrato =:Contrato and a.snumeroorden =:Orden and tm.sClasificacion="Tarifa Diaria" '+
                                'and a.swbs like :wbs and a.dIdFecha between :Inicio and :Final and tm2.lGrafica="Si" ' +
                                'group by a.swbs,a.dIdFecha,a.sHoraInicio,a.sIdClasificacion,a.sHoraFinal '+
                                'Order by a.dIdFecha,a.sHoraInicio,a.sIdClasificacion,a.sHoraFinal';

            QrPaquete.ParamByName('ContratoB').AsString := global_Contrato_Barco;
            QrPaquete.ParamByName('Contrato').AsString := Global_Contrato;
            QrPaquete.ParamByName('Convenio').AsString := Global_Convenio;
            QrPaquete.ParamByName('Orden').AsString := tsNumeroOrden.KeyValue;
            QrPaquete.ParamByName('Inicio').AsDate := RealDates.Inicio;
            QrPaquete.ParamByName('Final').AsDate := RealDates.Termino;
            QrPaquete.ParamByName('wbs').AsString:= zrqPrograma.FieldByName('sWbs').AsString + '%';
            QrPaquete.Open;

            MryAvances.Active:=False;
            MryAvances.EmptyTable;
            MryAvances.CopyStructure(QrPaquete);
            MryAvances.Active:=True;

            OriFecha:=StrToDate('23/02/1984');
            OriClasif:='ABY';
            OriHora:='SAHL';
            oHoja.Rows[IntToStr(ren)+ ':'+IntToStr(ren+3)].RowHeight := 141.75;

            //  tmpFechaI,TmpFechaT:TDateTime;
            //sSumaTotalHrs:string;
            sSumaTotalHrs:='00:00';
            QrPaquete.First;
            while not QrPaquete.Eof do
            begin
              if QrPaquete.RecNo=1 then
                tmpFechaI:=StrToDateTime(QrPaquete.FieldByName('dIdFecha').AsString + ' ' + QrPaquete.FieldByName('sHoraInicio').AsString)
              else
                if QrPaquete.RecNo=QrPaquete.RecordCount then
                  TmpFechaT:=StrToDateTime(QrPaquete.FieldByName('dIdFecha').AsString + ' ' + QrPaquete.FieldByName('sHoraFinal').AsString);

              if (QrPaquete.FieldByName('dIdFecha').AsDateTime<>OriFecha) or (QrPaquete.FieldByName('sIdClasificacion').AsString<>OriClasif) then
              begin
                MryAvances.Append;
                for I := 0 to QrPaquete.FieldCount-1 do
                begin
                  MryAvances.Fields[i].ReadOnly:=False;
                  if MryAvances.Fields[i].DataType=ftFloat then
                    MryAvances.FieldByName(QrPaquete.FieldDefs.Items[i].Name).AsFloat:= QrPaquete.FieldByName(QrPaquete.FieldDefs.Items[i].Name).AsFloat
                  else
                    MryAvances.FieldByName(QrPaquete.FieldDefs.Items[i].Name).Value:= QrPaquete.FieldByName(QrPaquete.FieldDefs.Items[i].Name).Value;

                end;
                MryAvances.Post;
                OriFecha:=QrPaquete.FieldByName('dIdFecha').AsDateTime;
                OriClasif:=QrPaquete.FieldByName('sIdClasificacion').AsString;
                OriHora:=QrPaquete.FieldByName('sHoraFinal').AsString;

                sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(QrPaquete.FieldByName('HoraF').AsString,QrPaquete.FieldByName('HoraI').AsString));
              end
              else
              begin
                if OriHora=QrPaquete.FieldByName('sHoraInicio').AsString then
                begin
                  MryAvances.Edit;
                  MryAvances.FieldByName('sHoraFinal').AsString:=QrPaquete.FieldByName('sHoraFinal').AsString;
                  MryAvances.FieldByName('dAvance').AsFloat:=MryAvances.FieldByName('dAvance').AsFloat + QrPaquete.FieldByName('dAvance').AsFloat;
                  MryAvances.Post;
                  OriHora:=QrPaquete.FieldByName('sHoraFinal').AsString;
                  //sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(qAvances.FieldByName('sHoraFinal').AsString,qAvances.FieldByName('sHoraInicio').AsString));
                  sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(QrPaquete.FieldByName('HoraF').AsString,QrPaquete.FieldByName('HoraI').AsString));

                end
                else
                begin
                  MryAvances.Append;
                  for I := 0 to QrPaquete.FieldCount-1 do
                  begin
                    MryAvances.Fields[i].ReadOnly:=False;
                    MryAvances.FieldByName(QrPaquete.FieldDefs.Items[i].Name     ).Value:= QrPaquete.FieldByName(QrPaquete.FieldDefs.Items[i].Name).Value;
                  end;

                  MryAvances.Post;
                  OriHora:=QrPaquete.FieldByName('sHoraFinal').AsString;
                  //sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(qAvances.FieldByName('sHoraFinal').AsString,qAvances.FieldByName('sHoraInicio').AsString));
                  sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(QrPaquete.FieldByName('HoraF').AsString,QrPaquete.FieldByName('HoraI').AsString));

                end;
              end;
              QrPaquete.Next;
            end;

            Graficar := ((MryDetallePdas.FieldByName('dFechaInicio').AsFloat > 0) and (MryDetallePdas.FieldByName('dFechaInicio').AsFloat < pTerminoP)) or cbPartidas.Checked;

            //Graficar := ((qPaquetes.FieldByName('dInicio').AsFloat > 0) and (qPaquetes.FieldByName('dInicio').AsFloat < pTerminoP) and (qPaquetes.FieldByName('dAvance').AsFloat > 0)) or cbPartidas.Checked;
          End
          else
          begin
            qAvances.Close;
            qAvances.ParamByName('wbs').AsString := zrqPrograma.FieldByName('sWbs').AsString;
            qAvances.Open;

            qAnteriorPda.Close;
            qAnteriorPda.ParamByName('wbs').AsString := zrqPrograma.FieldByName('sWbs').AsString;
            qAnteriorPda.Open;

            if cbTipoGrafica.ItemIndex=2 then
            begin

              QrDetallePdas.Active:=False;
              QrDetallePdas.SQL.Text:='select sNumeroOrden,swbs,sNumeroActividad,dFechaInicio,cast(sHoraInicio as Time) as sHoraInicio,dFechaFinal,'+
                                      'cast(sHoraFinal as Time) as sHoraFinal,sHoraInicio as HoraI,sHoraFinal as HoraF '+
                                      'from actividadesxorden_detalle where sContrato=:Contrato and sIdConvenio=:Convenio' + #10 +
                                      'and sNumeroOrden=:Orden and swbs=:wbs and dFechaInicio Between :Inicio and :Final ' +
                                      'order by dFechaInicio,Time(sHoraInicio)';
              QrDetallePdas.ParamByName('Contrato').AsString := Global_Contrato;
              QrDetallePdas.ParamByName('Convenio').AsString := Global_Convenio;
              QrDetallePdas.ParamByName('Orden').AsString:= tsNumeroOrden.KeyValue;
              QrDetallePdas.ParamByName('wbs').AsString:= zrqPrograma.FieldByName('sWbs').AsString;
              QrDetallePdas.ParamByName('Inicio').AsDate := RealDates.Inicio;
              QrDetallePdas.ParamByName('Final').AsDate := RealDates.Termino;
              QrDetallePdas.Open;

              MryDetallePdas.Active:=False;
              MryDetallePdas.EmptyTable;
              MryDetallePdas.CopyStructure(QrDetallePdas);
              MryDetallePdas.Active:=True;
              sSumaTotalHrsP:='00:00';
              QrDetallePdas.First;
              while not QrDetallePdas.Eof do
              begin
                if QrDetallePdas.RecNo=1 then
                  tmpFechaIP:=StrToDateTime(QrDetallePdas.FieldByName('dFechaInicio').AsString + ' ' + QrDetallePdas.FieldByName('sHoraInicio').AsString)
                else
                  if QrDetallePdas.RecNo=QrDetallePdas.RecordCount then
                    TmpFechaTP:=StrToDateTime(QrDetallePdas.FieldByName('dFechaFinal').AsString + ' ' + QrDetallePdas.FieldByName('sHoraFinal').AsString);


                MryDetallePdas.Append;
                for I := 0 to QrDetallePdas.FieldCount-1 do
                begin
                  MryDetallePdas.Fields[i].ReadOnly:=False;
                  if MryDetallePdas.Fields[i].DataType=ftFloat then
                    MryDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).AsFloat:= QrDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).AsFloat
                  else
                    MryDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).Value:= QrDetallePdas.FieldByName(QrDetallePdas.FieldDefs.Items[i].Name).Value;

                end;
                MryDetallePdas.Post;
                sSumaTotalHrsP:=sfnSumaHoras(sSumaTotalHrsP,sfnRestaHoras(QrDetallePdas.FieldByName('HoraF').AsString,QrDetallePdas.FieldByName('HoraI').AsString));

                QrDetallePdas.Next;
              end;


              MryAvances.Active:=False;
              MryAvances.EmptyTable;
              MryAvances.CopyStructure(qAvances);
              MryAvances.Active:=True;

              OriFecha:=StrToDate('23/02/1984');
              OriClasif:='ABY';
              OriHora:='SAHL';
              oHoja.Rows[IntToStr(ren)+ ':'+IntToStr(ren+3)].RowHeight := 141.75;

              //  tmpFechaI,TmpFechaT:TDateTime;
              //sSumaTotalHrs:string;
              sSumaTotalHrs:='00:00';
              qAvances.First;
              while not qAvances.Eof do
              begin
                if qAvances.RecNo=1 then
                  tmpFechaI:=StrToDateTime(qAvances.FieldByName('dIdFecha').AsString + ' ' + qAvances.FieldByName('sHoraInicio').AsString)
                else
                  if qAvances.RecNo=qAvances.RecordCount then
                    TmpFechaT:=StrToDateTime(qAvances.FieldByName('dIdFecha').AsString + ' ' + qAvances.FieldByName('sHoraFinal').AsString);

                if (qAvances.FieldByName('dIdFecha').AsDateTime<OriFecha) or (qAvances.FieldByName('sIdClasificacion').AsString<>OriClasif) then
                begin
                  MryAvances.Append;
                  for I := 0 to qAvances.FieldCount-1 do
                  begin
                    MryAvances.Fields[i].ReadOnly:=False;
                    if MryAvances.Fields[i].DataType=ftFloat then
                      MryAvances.FieldByName(qAvances.FieldDefs.Items[i].Name).AsFloat:= qAvances.FieldByName(qAvances.FieldDefs.Items[i].Name).AsFloat
                    else
                      MryAvances.FieldByName(qAvances.FieldDefs.Items[i].Name).Value:= qAvances.FieldByName(qAvances.FieldDefs.Items[i].Name).Value;

                  end;
                  MryAvances.Post;
                  OriFecha:=qAvances.FieldByName('dIdFecha').AsDateTime;
                  OriClasif:=qAvances.FieldByName('sIdClasificacion').AsString;
                  OriHora:=qAvances.FieldByName('sHoraFinal').AsString;

                  sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(qAvances.FieldByName('HoraF').AsString,qAvances.FieldByName('HoraI').AsString));
                end
                else
                begin
                  if OriHora=qAvances.FieldByName('sHoraInicio').AsString then
                  begin
                    MryAvances.Edit;
                    MryAvances.FieldByName('sHoraFinal').AsString:=qAvances.FieldByName('sHoraFinal').AsString;
                    MryAvances.FieldByName('dAvance').AsFloat:=MryAvances.FieldByName('dAvance').AsFloat + qAvances.FieldByName('dAvance').AsFloat;
                    MryAvances.Post;
                    OriHora:=qAvances.FieldByName('sHoraFinal').AsString;
                    //sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(qAvances.FieldByName('sHoraFinal').AsString,qAvances.FieldByName('sHoraInicio').AsString));
                    sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(qAvances.FieldByName('HoraF').AsString,qAvances.FieldByName('HoraI').AsString));

                  end
                  else
                  begin
                    MryAvances.Append;
                    for I := 0 to qAvances.FieldCount-1 do
                    begin
                      MryAvances.Fields[i].ReadOnly:=False;
                      MryAvances.FieldByName(qAvances.FieldDefs.Items[i].Name     ).Value:= qAvances.FieldByName(qAvances.FieldDefs.Items[i].Name).Value;
                    end;

                    MryAvances.Post;
                    OriHora:=qAvances.FieldByName('sHoraFinal').AsString;
                    //sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(qAvances.FieldByName('sHoraFinal').AsString,qAvances.FieldByName('sHoraInicio').AsString));
                    sSumaTotalHrs:=sfnSumaHoras(sSumaTotalHrs,sfnRestaHoras(qAvances.FieldByName('HoraF').AsString,qAvances.FieldByName('HoraI').AsString));

                  end;
                end;
                qAvances.Next;
              end;

              Graficar := (MryAvances.RecordCount > 0) or cbPartidas.Checked;
            end
            else
              Graficar := (qAvances.RecordCount > 0) or cbPartidas.Checked;
          end;
        end
        else
          Graficar := True;

        if Graficar then
        begin
          //*****************************************************************************************************************************
          //* Inicia proceso de avance programado
          MryDetallePdas.First;
          while not MryDetallePdas.Eof do
          Begin
            if MryDetallePdas.FieldByName('dFechaInicio').AsDateTime < pTerminoP then
            begin
              myInicio := trunc(MryDetallePdas.FieldByName('dFechaInicio').AsFloat);
              myTermino := Trunc(MryDetallePdas.FieldByName('dFechaFinal').AsFloat);// + 1;

              ColBarra := rAncho + 1;

              if cbTipoGrafica.ItemIndex=2 then
              begin
                while oGantt.Cells[LocRen, ColBarra].Value < myInicio do
                begin

                  ColBarra := ColBarra + 4{1};

                end;


                if MryDetallePdas.FieldByName('sHoraInicio').AsFloat=0 then
                begin
                  if MryDetallePdas.FieldByName('sHoraInicio').AsString='00:00' then
                    fHoraI:=0
                  else
                    fHoraI:=1;
                end
                else
                  fHoraI:=MryDetallePdas.FieldByName('sHoraInicio').AsFloat;

                Cuantos:=Trunc(ColBarra);

              end
              else
              begin
                while oGantt.Cells[LocRen, ColBarra + 1].Value <= myInicio do
                  ColBarra := ColBarra + 1;
              end;

              if cbTipoGrafica.ItemIndex=2 then
              begin
                {fHoraF:=Frac(zrqPrograma.FieldByName('dFechaFinal').AsFloat);
                if fHoraF=0 then
                  fHoraF:=1;}
                if MryDetallePdas.FieldByName('sHoraFinal').AsFloat=0 then
                begin
                  if MryDetallePdas.FieldByName('sHoraFinal').AsString='00:00' then
                    fHoraF:=0
                  else
                    fHoraF:=1;
                end
                else
                  fHoraF:=MryDetallePdas.FieldByName('sHoraFinal').AsFloat;


                //Cuantos:=0;
                while (oGantt.Cells[LocRen,Cuantos + 1].Value <> 0) and (oGantt.Cells[LocRen, Cuantos + 1].Value < myTermino) do
                begin
                  Cuantos := Cuantos + 4;//1;
                end;
                Cuantos := Cuantos-(Trunc(ColBarra));
                //Cuantos:=Cuantos - 4; }

              end
              else
              begin
                Cuantos := 0;



                while (oGantt.Cells[LocRen, rAncho + Cuantos + 2].Value <> 0) and (oGantt.Cells[LocRen, rAncho + Cuantos + 2].Value <= myTermino) do
                  Cuantos := Cuantos + 1;
              end;

              if cbTipoGrafica.ItemIndex=2 then
              begin
                 //if myTermino > pTerminoP then

                //  DurBarra := (((pTerminoP - myInicio) / (pTerminoP - pInicioP)   ) ) * wCol
                // else
                 begin
                  //DurBarra := (((myTermino) - myInicio) / (( trunc(pTerminoP) - trunc(pInicioP)) * 4)      )  * wCol;

                  iColHr:=trunc(fHoraI * 4);
                  ColBarra:=ColBarra+iColHr;

                 { iColHr:=trunc(fHoraF * 4);
                  if iColHr>0 then
                    Cuantos:=Cuantos+(iColHr-1);  }
                  Cuantos:=CUANTOS +  (trunc(fHoraF * 4) - trunc(fHoraI * 4));

                  EAux:=fHoraI * 24;
                  iAux:=Trunc(EAux);
                  fRestaI:=(iAux MOD 6) + Frac(EAux) ;
                  fRestaI:=fRestaI/6;

                  EAux:=fHoraF * 24;
                  iColHrF:=trunc(EAux);
                  pFactor:= (iColHrF MOD 6) + Frac(EAux) ;
                  pFactor:=pFactor/6;

                  //iColHr:=trunc(fHoraF * 4);
                  DurBarra:=(Cuantos * AnchoColumnasPixels)-(AnchoColumnasPixels * fRestaI)    ;
                  DurBarra:=DurBarra +(pFactor * AnchoColumnasPixels);

                 end;



              end
              else
              begin
                if myTermino > pTerminoP then
                  DurBarra := (((pTerminoP) - myInicio) / (pTerminoP - pInicioP)) * wCol
                else
                  DurBarra := (((myTermino) - myInicio) / (pTerminoP - pInicioP)) * wCol;

              end;
              if DurBarra < 0 then DurBarra := 0;

              // Verificar si es una partida concentradora
              if zrqPrograma.FieldByName('sTipoActividad').AsString = 'Paquete' then
              Begin
                // Color amarillo para las concentradoras
                SetSelect(oHoja,'A' + Trim(inttostr(Ren)) + ':' + Letra(ColENd+ 3) + Trim(Inttostr(Ren + 3)));
                oGantt.Selection.Interior.Color:= 13434828;//colorIndex := 6;
                oGantt.Selection.Interior.Pattern := xlSolid;

                (*ColBarra := rAncho + 1;

                while oGantt.Cells[LocRen, ColBarra].Value < myInicio do
                begin

                  ColBarra := ColBarra + 4{1};

                end;


                if MryDetallePdas.FieldByName('sHoraInicio').AsFloat=0 then
                begin
                  if MryDetallePdas.FieldByName('sHoraInicio').AsString='00:00' then
                    fHoraI:=0
                  else
                    fHoraI:=1;
                end
                else
                  fHoraI:=MryDetallePdas.FieldByName('sHoraInicio').AsFloat;

                Cuantos:=Trunc(ColBarra);

                if MryDetallePdas.FieldByName('sHoraFinal').AsFloat=0 then
                begin
                  if MryDetallePdas.FieldByName('sHoraFinal').AsString='00:00' then
                    fHoraF:=0
                  else
                    fHoraF:=1;
                end
                else
                  fHoraF:=MryDetallePdas.FieldByName('sHoraFinal').AsFloat;


                //Cuantos:=0;
                while (oGantt.Cells[LocRen,Cuantos + 1].Value <> 0) and (oGantt.Cells[LocRen, Cuantos + 1].Value < myTermino) do
                begin
                  Cuantos := Cuantos + 4;//1;
                end;
                Cuantos := Cuantos-(Trunc(ColBarra)); *)

                // Colocar barra concentradora
                if (trunc(zrqPrograma.FieldByName('dFechaInicio').AsFloat) < pTerminoP) and (trunc(zrqPrograma.FieldByName('dFechaFinal').AsFloat) >= pInicioP) then
                begin
                   // oGantt.Selection.Cut;
                  oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren)].Select;
                  oHoja.Shapes.AddShape(1{61},
                  oGantt.Selection.left,oGantt.Selection.top, DurBarra, (fAlto)).Select;

                 // oGantt.ActiveSheet.Paste;
                  oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);
                  oGantt.Selection.ShapeRange.Fill.Solid;
                  oGantt.Selection.ShapeRange.Fill.ForeColor.RGB:=RGB(7, 30, 169);//SchemeColor := 0;
                  oGantt.Selection.ShapeRange.Line.Visible := msotrue;
                  oGantt.Selection.ShapeRange.Line.Weight:=0.75;
                  {Selection.ShapeRange.Line
          .Visible = msoTrue
          .Weight = 0.75}
                  //oGantt.Selection.ShapeRange.Width := DurBarra;
                  // Verificar si es necesario mover la barra hacia delante

                  if cbTipoGrafica.ItemIndex=2 then
                  begin
                    {iColHr:=trunc(zrqPrograma.FieldByName('dFechaInicio').AsFloat * 24);
                    pFactor:= (iColHr MOD 6) ;
                    pFactor:=pFactor/6; }
                    EAux:=fHoraI * 24;
                    iAux:=Trunc(EAux);
                    fRestaI:=(iAux MOD 6) + Frac(EAux) ; ;
                    fRestaI:=fRestaI/6;

                    oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * fRestaI) );

                  end
                  else
                  begin
                    if oGantt.Cells[LocRen, ColBarra].Value < myInicio then
                      oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(myInicio) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));
                  end;




                End;


              end
              Else
              Begin
                // Colocar barra de actividad programada
                if trunc(MryDetallePdas.FieldByName('dFechaInicio').AsFloat) < pTerminoP then
                begin

                  oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren)].Select;
                  oHoja.Shapes.AddShape(1{61},
                  oGantt.Selection.left,oGantt.Selection.top, DurBarra, (fAlto)).Select;
                  oGantt.Selection.ShapeRange.Fill.Solid;
                  oGantt.Selection.ShapeRange.Fill.ForeColor.RGB:=RGB(7, 30, 169);//.SchemeColor := 48;
                  //oGantt.Selection.ShapeRange.Fill.OneColorGradient(msoGradientHorizontal, 4, 0.23);
                  oGantt.Selection.ShapeRange.Line.Visible := msotrue;
                  oGantt.Selection.ShapeRange.Line.Weight:=0.75;

                  oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);
                  if cbTipoGrafica.ItemIndex=2 then
                  begin
                   { iColHr:=trunc(zrqPrograma.FieldByName('dFechaInicio').AsFloat * 24);
                    pFactor:= (iColHr MOD 6) ;
                    pFactor:=pFactor/6;}

                    EAux:=fHoraI * 24;
                    iAux:=Trunc(EAux);
                    fRestaI:=(iAux MOD 6) + Frac(EAux) ;;
                    fRestaI:=fRestaI/6;
                    oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * fRestaI) );
                  end
                  else
                  begin
                    // Verificar si es necesario mover la barra hacia delante
                    if oGantt.Cells[LocRen, ColBarra].Value < myInicio then
                      oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(myInicio) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));

                  end;
                end;
              end;
            end;
            MryDetallePdas.next;
          end;
            //* Terminar proceso de avances programados
            //*****************************************************************************************************************************























          if (cbTipoGrafica.ItemIndex = 1) or (cbTipoGrafica.ItemIndex=2) then
          begin
            //*****************************************************************************************************************************
            //* Inicia proceso de avances reales
            if (cbTipoGrafica.ItemIndex = 1) then
            begin
              {$REGION 'Graficas Normales'}
              if zrqPrograma.FieldByName('sTipoActividad').AsString = 'Paquete' then
              Begin
                SetSelect(oHoja,'I' + Trim(IntToStr(Ren)) + ':I' + Trim(IntToStr(Ren + 3)));
                oGantt.Selection.Value := qPaquetes.FieldByName('dAvance').AsFloat / 100;
                oGantt.Selection.MergeCells := True;

                // Verificar si se debe mostrar el avance global calculado o el totalizado de reportes diarios
                if (zrqPrograma.FieldByName('iNivel').AsInteger = 0) and Not cbCalcular.Checked then   //and (Opciones.ItemIndex <> 0) then
                begin
                  SetSelect(oHoja,'L' + Trim(IntToStr(Ren)) + ':L' + Trim(IntToStr(Ren + 3)));
                  oGantt.Selection.Value := qPaquete0.FieldByName('dAvanceAnterior').AsFloat / 100;
                  oGantt.Selection.MergeCells := True;

                  SetSelect(oHoja,'M' + Trim(IntToStr(Ren)) + ':M' + Trim(IntToStr(Ren + 3)));
                  oGantt.Selection.Value := qPaquete0.FieldByName('dAvance').AsFloat / 100;
                  oGantt.Selection.MergeCells := True;

                  SetSelect(oHoja,'N' + Trim(IntToStr(Ren)) + ':N' + Trim(IntToStr(Ren + 3)));
                  oGantt.Selection.Value := (qPaquete0.FieldByName('dAvanceAnterior').AsFloat + qPaquete0.FieldByName('dAvance').AsFloat) / 100;
                  oGantt.Selection.MergeCells := True;
                end
                else
                begin
                  SetSelect(oHoja,'L' + Trim(IntToStr(Ren)) + ':L' + Trim(IntToStr(Ren + 3)));
                  oGantt.Selection.Value := qPaquetes.FieldByName('dAvanceAnterior').AsFloat / 100;
                  oGantt.Selection.MergeCells := True;

                  SetSelect(oHoja,'M' + Trim(IntToStr(Ren)) + ':M' + Trim(IntToStr(Ren + 3)));
                  oGantt.Selection.Value := qPaquetes.FieldByName('dAvance').AsFloat / 100;
                  oGantt.Selection.MergeCells := True;

                  SetSelect(oHoja,'N' + Trim(IntToStr(Ren)) + ':N' + Trim(IntToStr(Ren + 3)));
                  oGantt.Selection.Value := (qPaquetes.FieldByName('dAvanceAnterior').AsFloat + qPaquetes.FieldByName('dAvance').AsFloat) / 100;
                  oGantt.Selection.MergeCells := True;
                end;

                if (qPaquetes.FieldByName('dInicio').AsFloat > 0) and (qPaquetes.FieldByName('dInicio').AsFloat < pTerminoP) then
                begin
                  myInicio := qPaquetes.FieldByName('dInicio').AsFloat;
                  myTermino := qPaquetes.FieldByName('dFinal').AsFloat + 1;

                  // Localizar la columna en donde se debe generar la barra
                  ColBarra := rAncho + 1;
                  while oGantt.Cells[LocRen, ColBarra + 1].Value <= myInicio do
                    ColBarra := ColBarra + 1;

                  Cuantos := 0;
                  while (oGantt.Cells[LocRen, rAncho + Cuantos + 2].Value <> 0) and (oGantt.Cells[LocRen, rAncho + Cuantos + 2].Value <= myTermino) do
                    Cuantos := Cuantos + 1;

                  if myTermino > pTerminoP then
                    DurBarra := (((pTerminoP) - myInicio) / (pTerminoP - pInicioP)) * wCol
                  else
                    DurBarra := (((myTermino) - myInicio) / (pTerminoP - pInicioP)) * wCol;

                  if DurBarra < 0 then DurBarra := 0;

                  if DurBarra > 0 then
                  begin
                    oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto / 2)).Select;
                    oGantt.Selection.Cut;
                    oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                    oGantt.ActiveSheet.Paste;
                    oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);
                    // Verificar si es necesario mover la barra hacia delante
                    if oGantt.Cells[LocRen, ColBarra].Value < myInicio then
                      oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(myInicio) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));

                    oGantt.Selection.ShapeRange.Fill.Solid;
                    oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 16;
                    oGantt.Selection.ShapeRange.Line.Visible := msotrue;
                  oGantt.Selection.ShapeRange.Line.Weight:=0.75;
                    Val1 := oGantt.Selection.Left + oGantt.Selection.Width;

                    if qPaquetes.FieldByName('dInicio').AsFloat >= pInicioP then
                    Begin
                      // Colocar indicador de inicio de concentradora
                      oHoja.Shapes.AddShape(msoShapeRightTriangle, 0, 0, fAlto / 2, fAlto / 2).Select;
                      oGantt.Selection.ShapeRange.IncrementRotation(90);
                      oGantt.Selection.ShapeRange.Fill.Solid;
                      oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 16;
                      oGantt.Selection.ShapeRange.Line.Visible := msoFalse;
                      oGantt.Selection.Cut;
                      oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                      oGantt.ActiveSheet.Paste;
                      oGantt.Selection.ShapeRange.IncrementTop(fAlto / 1.5);
                      // Verificar si es necesario mover la barra hacia delante
                      if oGantt.Cells[LocRen, ColBarra].Value < myInicio then
                        oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(myInicio) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));
                    End;

                    if (qPaquetes.FieldByName('dFinal').AsFloat <= pTerminoP) and (qPaquetes.FieldByName('dFinal').AsFloat > pInicioP) then
                    Begin
                      // Colocar indicador de termino de concentradora
                      oHoja.Shapes.AddShape(msoShapeRightTriangle, 0, 0, fAlto / 2, fAlto / 2).Select;
                      oGantt.Selection.ShapeRange.IncrementRotation(180);
                      oGantt.Selection.ShapeRange.Fill.Solid;
                      oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 16;
                      oGantt.Selection.ShapeRange.Line.Visible := msoFalse;
                      oGantt.Selection.Cut;
                      oHoja.Range['A' + IntToStr(Ren + 2)].Select;
                      oGantt.ActiveSheet.Paste;
                      oGantt.Selection.ShapeRange.IncrementTop(fAlto / 1.5);
                      oGantt.Selection.ShapeRange.IncrementLeft(Val1 - oGantt.Selection.Width);
                    End;
                  end;
                End
              End
              else
              begin
                AvancePond := 0;
                Avance := 0;
                LastDate := 0;
                LastColor := 5;


                while Not qAvances.Eof do
                begin
                  AvancePond := AvancePond + qAvances.FieldByName('dAvancePond').AsFloat;
                  Avance := Avance + qAvances.FieldByName('dAvance').AsFloat;

                  // Verificar si existen cantidades exedentes
                  if qAvances.FieldByName('dExcedente').AsFloat > 0 then
                  begin
                    // Existen cantidades excedentes, verificar si la cantidad es total o parcialmente excedente
                    if qAvances.FieldByName('dCantidad').AsFloat > qAvances.FieldByName('dExcedente').AsFloat then
                    begin
                      // Partir la barra en dos partes
                        // Poner la barra de la cantidad normal
                        FactorExcedente := qAvances.FieldByName('dExcedente').AsFloat / (qAvances.FieldByName('dExcedente').AsFloat + qAvances.FieldByName('dCantidadTotal').AsFloat);
                        PorcentajeExcedente := qAvances.FieldByName('dExcedente').AsFloat / (qAvances.FieldByName('dExcedente').AsFloat + qAvances.FieldByName('dCantidad').AsFloat);

                        ColBarra := rAncho + 1;
                        while oGantt.Cells[LocRen, ColBarra + 1].Value <= qAvances.FieldByName('dIdFecha').AsDateTime do
                          ColBarra := ColBarra + 1;

                        if RadioButton1.Checked then
                        begin
                          // Verificar si se debe generar la cantidad excedente
                          if RadioGroup1.ItemIndex = 0 then
                            oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := FloatToStrF(qAvances.FieldByName('dCantidad').AsFloat - qAvances.FieldByName('dExcedente').AsFloat, ffNumber, 16, 4) + ' / ' + FloatToStrF(qAvances.FieldByName('dExcedente').AsFloat, ffNumber, 16, 4)
                          else
                            oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := qAvances.FieldByName('dCantidad').AsFloat - qAvances.FieldByName('dExcedente').AsFloat;
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].HorizontalAlignment := xlRight;
                        end
                        else
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := FactorExcedente;

                        if ((IncDay(LastDate) = qAvances.FieldByName('dIdFecha').AsDateTime) and (LastDate > 0)) then
                        begin
                          DurBarra := oGantt.Selection.ShapeRange.Width;
                          DurBarra := DurBarra + (((1 / (pTerminoP - pInicioP)) * wCol) * (1 - PorcentajeExcedente));
                          oGantt.Selection.ShapeRange.Width := DurBarra;
                        end
                        else
                        begin
                          DurBarra := (1 / (pTerminoP - pInicioP)) * wCol;
                          DurBarra := DurBarra * (1 - FactorExcedente);

                          oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto)).Select;
                          oGantt.Selection.Cut;
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                          oGantt.ActiveSheet.Paste;
                          oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);

                          oGantt.Selection.ShapeRange.Fill.Solid;
                          oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 5;
                          oGantt.Selection.ShapeRange.Line.Visible := msoFalse;

                          // Verificar si es necesario mover la barra hacia delante
                          if oGantt.Cells[LocRen, ColBarra].Value < qAvances.FieldByName('dIdFecha').AsDateTime then
                            oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(qAvances.FieldByName('dIdFecha').AsDateTime) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));
                        end;

                        // Poner la barra de la fracci�n excedida solo si el usuario lo ha indicado
                        if RadioGroup1.ItemIndex = 0 then
                        begin
                          DurBarra := ((1 / (pTerminoP - pInicioP)) * wCol) * FactorExcedente;

                          oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto)).Select;
                          oGantt.Selection.Cut;
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                          oGantt.ActiveSheet.Paste;
                          oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);
                          oGantt.Selection.ShapeRange.IncrementLeft(((1 / (pTerminoP - pInicioP)) * wCol) * (1 - FactorExcedente));

                          oGantt.Selection.ShapeRange.Fill.Solid;
                          oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 2;
                          oGantt.Selection.ShapeRange.Line.Visible := msoFalse;

                          LastColor := 2;
                        end;
                      //*****************************************************************
                    end
                    else
                    begin
                      // Toda la cantidad es excedente
                      if RadioGroup1.ItemIndex = 0 then
                      begin
                        FactorExcedente := qAvances.FieldByName('dExcedente').AsFloat / qAvances.FieldByName('dCantidadTotal').AsFloat;
                        // Localizar la columna en donde se debe generar la barra
                        ColBarra := rAncho + 1;
                        while oGantt.Cells[LocRen, ColBarra + 1].Value <= myInicio do
                          ColBarra := ColBarra + 1;
                        //ColBarra := (qAvances.FieldByName('dIdFecha').AsDateTime - pInicioP) + rAncho + 1;
                        if RadioButton1.Checked then
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value + qAvances.FieldByName('dExcedente').AsFloat
                        else
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value + FactorExcedente;

                        if ((IncDay(LastDate) = qAvances.FieldByName('dIdFecha').AsDateTime) and (LastDate > 0)) And (LastColor = 2) then
                        begin
                          DurBarra := oGantt.Selection.ShapeRange.Width;
                          DurBarra := DurBarra + ((1 / (pTerminoP - pInicioP)) * wCol);
                          oGantt.Selection.ShapeRange.Width := DurBarra;
                        end
                        else
                        begin
                          DurBarra := (1 / (pTerminoP - pInicioP)) * wCol;

                          oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto)).Select;
                          oGantt.Selection.Cut;
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                          oGantt.ActiveSheet.Paste;
                          oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);

                          oGantt.Selection.ShapeRange.Fill.Solid;
                          oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 2;
                          oGantt.Selection.ShapeRange.Line.Visible := msoFalse;

                          if oGantt.Cells[LocRen, ColBarra].Value < qAvances.FieldByName('dIdFecha').AsDateTime then
                            oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(qAvances.FieldByName('dIdFecha').AsDateTime) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));
                        end;
                        LastColor := 2;
                      end;
                      //*******************************************************************************
                    end;
                  end
                  else
                  begin
                    // Toda la cantidad es normal
                    // Poner avances diarios (Sin seleccionar la celda)
                    ColBarra := rAncho + 1;
                    while oGantt.Cells[LocRen, ColBarra + 1].Value <= qAvances.FieldByName('dIdFecha').AsDateTime do
                      ColBarra := ColBarra + 1;
                    //ColBarra := (qAvances.FieldByName('dIdFecha').AsDateTime - pInicioP) + rAncho + 1;

                    Cuantos := 0;
                    while (oGantt.Cells[LocRen, ColBarra + Cuantos + 2].Value <> 0) and (oGantt.Cells[LocRen, ColBarra + Cuantos + 2].Value < qAvances.FieldByName('dIdFecha').AsFloat) do
                      Cuantos := Cuantos + 1;

                    if RadioButton1.Checked then
                      oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value + qAvances.FieldByName('dCantidad').AsFloat
                    else
                      oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value + (qAvances.FieldByName('dAvance').AsFloat / 100);

                    if (IncDay(LastDate) = qAvances.FieldByName('dIdFecha').AsDateTime) and (LastDate > 0) then
                    begin
                      DurBarra := oGantt.Selection.ShapeRange.Width;
                      DurBarra := DurBarra + ((1 / (pTerminoP - pInicioP)) * wCol);
                      oGantt.Selection.ShapeRange.Width := DurBarra;
                    end
                    else
                    begin
                      DurBarra := (1 / (pTerminoP - pInicioP)) * wCol;

                      oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto)).Select;
                      oGantt.Selection.Cut;
                      oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                      oGantt.ActiveSheet.Paste;
                      oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);

                      oGantt.Selection.ShapeRange.Fill.Solid;
                      oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 5;
                      oGantt.Selection.ShapeRange.Line.Visible := msoFalse;
                      // Verificar si es necesario mover la barra hacia delante
                      Val1 := oGantt.Cells[LocRen, ColBarra + Cuantos].Value;
                      if Val1 <= qAvances.FieldByName('dIdFecha').AsDateTime then
                      begin
                        Val1 := (Trunc(qAvances.FieldByName('dIdFecha').AsDateTime) - Trunc(oGantt.Cells[LocRen, ColBarra + Cuantos].Value));
                        oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * Val1);
                      end;
                    end;
                    LastColor := 5;
                  end;

                  LastDate := qAvances.FieldByName('dIdFecha').AsDateTime;
                  qAvances.Next;
                end;

                SetSelect(oHoja,'I' + Trim(IntToStr(Ren)) + ':I' + Trim(IntToStr(Ren + 3)));
                oGantt.Selection.Value := Avancepond / 100;
                oGantt.Selection.MergeCells := True;

                SetSelect(oHoja,'L' + Trim(IntToStr(Ren)) + ':L' + Trim(IntToStr(Ren + 3)));
                oGantt.Selection.Value := qAnteriorPda.FieldByName('dAvanceAnterior').AsFloat / 100;
                oGantt.Selection.MergeCells := True;

                SetSelect(oHoja,'M' + Trim(IntToStr(Ren)) + ':M' + Trim(IntToStr(Ren + 3)));
                oGantt.Selection.Value := Avance / 100;
                oGantt.Selection.MergeCells := True;

                SetSelect(oHoja,'N' + Trim(IntToStr(Ren)) + ':N' + Trim(IntToStr(Ren + 3)));
                oGantt.Selection.Value := (qAnteriorPda.FieldByName('dAvanceAnterior').AsFloat + Avance) / 100;
                oGantt.Selection.MergeCells := True;
              end;
              {$ENDREGION}

            end
            else
            begin

              if zrqPrograma.FieldByName('sTipoActividad').AsString = 'Paquete' then
              Begin

                SetSelect(oHoja,ColumnaNombre(ColEnd+1) + Trim(Inttostr(Ren)) + ':' +ColumnaNombre(ColEnd+1) + Trim(Inttostr(Ren+3)));
                oGantt.Selection.MergeCells := True;
                oGantt.Selection.VerticalAlignment := xlCenter;
                oGantt.Selection.HorizontalAlignment := xlCenter;
                oGantt.Selection.Font.Name:='Arial';
                oGantt.Selection.Font.Size := 48;
                oGantt.Selection.Value := zrqPrograma.FieldByName('dPonderado').AsFloat/100;

                //Este es avance Real
                MryAvances.first;
                while Not MryAvances.Eof do
                begin


                  if (MryAvances.FieldByName('dIdFecha').AsFloat > 0) and (MryAvances.FieldByName('dIdFecha').AsFloat < pTerminoP) then
                  begin
                    myInicio := MryAvances.FieldByName('dIdFecha').AsFloat;
                    myTermino := MryAvances.FieldByName('dIdFecha').AsFloat;// + 1;

                    ColBarra := rAncho + 1;

                    while oGantt.Cells[LocRen, ColBarra].Value < myInicio do
                    begin

                      ColBarra := ColBarra + 4{1};

                    end;


                    if MryAvances.FieldByName('sHoraInicio').AsFloat=0 then
                    begin
                      if MryAvances.FieldByName('HoraI').AsString='00:00' then
                        fHoraI:=0
                      else
                        fHoraI:=1;
                    end
                    else
                      fHoraI:=MryAvances.FieldByName('sHoraInicio').AsFloat;

                    Cuantos:=Trunc(ColBarra);

                    if MryAvances.FieldByName('sHoraFinal').AsFloat=0 then
                    begin
                      if MryAvances.FieldByName('HoraF').AsString='00:00' then
                        fHoraF:=0
                      else
                        fHoraF:=1;
                    end
                    else
                      fHoraF:=MryAvances.FieldByName('sHoraFinal').AsFloat;


                    //Cuantos:=0;
                    while (oGantt.Cells[LocRen,Cuantos + 1].Value <> 0) and (oGantt.Cells[LocRen, Cuantos + 1].Value < myTermino) do
                    begin
                      Cuantos := Cuantos + 4;//1;
                    end;
                    Cuantos := Cuantos-(Trunc(ColBarra));

                    iColHr:=trunc(fHoraI * 4);
                    ColBarra:=ColBarra+iColHr;

                   { iColHr:=trunc(fHoraF * 4);
                    if iColHr>0 then
                      Cuantos:=Cuantos+(iColHr-1);  }
                    Cuantos:=CUANTOS +  (trunc(fHoraF * 4) - trunc(fHoraI * 4));

                    EAux:=fHoraI * 24;
                    iAux:=Trunc(EAux);
                    fRestaI:=(iAux MOD 6) ;
                    fRestaI:=fRestaI/6;

                    EAux:=fHoraF * 24;
                    iColHrF:=trunc(EAux);
                    pFactor:= (iColHrF MOD 6) ;
                    pFactor:=pFactor/6;

                    //iColHr:=trunc(fHoraF * 4);
                    DurBarra:=(Cuantos * AnchoColumnasPixels)-(AnchoColumnasPixels * fRestaI)    ;
                    DurBarra:=DurBarra +(pFactor * AnchoColumnasPixels);

                    oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren+ 2)].Select;
                    oHoja.Shapes.AddShape(1{61},
                    oGantt.Selection.left,oGantt.Selection.top, DurBarra, (fAlto)).Select;

                   // oGantt.ActiveSheet.Paste;
                    oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);
                    oGantt.Selection.ShapeRange.Fill.Solid;
                    oGantt.Selection.ShapeRange.Fill.ForeColor.RGB{.SchemeColor} := ColorToRgb(tColor(MryAvances.FieldByName('iColor').AsInteger));
                    oGantt.Selection.ShapeRange.Line.Visible := msotrue;
                  oGantt.Selection.ShapeRange.Line.Weight:=0.75;

                    //oGantt.Selection.ShapeRange.Width := DurBarra;
                    // Verificar si es necesario mover la barra hacia delante


                    EAux:=fHoraI * 24;
                    iAux:=Trunc(EAux);
                    fRestaI:=(iAux MOD 6) ;
                    fRestaI:=fRestaI/6;

                    oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * fRestaI) );



                  End ;
                  MryAvances.next;

                end;
              End
              else
              begin




                MryAvances.first;
                while Not MryAvances.Eof do
                begin
                  AvancePond := AvancePond + MryAvances.FieldByName('dAvancePond').AsFloat;
                  Avance := Avance + MryAvances.FieldByName('dAvance').AsFloat;

                  // Verificar si existen cantidades exedentes
                  if MryAvances.FieldByName('dExcedente').AsFloat > 0 then
                  begin
                    // Existen cantidades excedentes, verificar si la cantidad es total o parcialmente excedente
                    if MryAvances.FieldByName('dCantidad').AsFloat > MryAvances.FieldByName('dExcedente').AsFloat then
                    begin
                      // Partir la barra en dos partes
                        // Poner la barra de la cantidad normal
                        FactorExcedente := MryAvances.FieldByName('dExcedente').AsFloat / (MryAvances.FieldByName('dExcedente').AsFloat + MryAvances.FieldByName('dCantidadTotal').AsFloat);
                        PorcentajeExcedente := MryAvances.FieldByName('dExcedente').AsFloat / (MryAvances.FieldByName('dExcedente').AsFloat + MryAvances.FieldByName('dCantidad').AsFloat);

                        ColBarra := rAncho + 1;
                        while oGantt.Cells[LocRen, ColBarra].Value <= MryAvances.FieldByName('dIdFecha').AsDateTime do
                          ColBarra := ColBarra + 4;


                        iColHr:=trunc(Frac(MryAvances.FieldByName('sHoraInicio').AsFloat) * 4);

                        ColBarra:=ColBarra+iColHr;

                        if ((IncDay(LastDate) = MryAvances.FieldByName('dIdFecha').AsDateTime) and (LastDate > 0)) then
                        begin
                          DurBarra := oGantt.Selection.ShapeRange.Width;
                          DurBarra := DurBarra + (((1 / (pTerminoP - pInicioP)) * wCol) * (1 - PorcentajeExcedente));
                          oGantt.Selection.ShapeRange.Width := DurBarra;
                        end
                        else
                        begin
                          DurBarra := (1 / (pTerminoP - pInicioP)) * wCol;
                          DurBarra := DurBarra * (1 - FactorExcedente);

                          oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto)).Select;
                          oGantt.Selection.Cut;
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                          oGantt.ActiveSheet.Paste;
                          oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);

                          oGantt.Selection.ShapeRange.Fill.Solid;
                          oGantt.Selection.ShapeRange.Fill.ForeColor.RGB{.SchemeColor} :=ColorToRgb(tColor(MryAvances.FieldByName('iColor').AsInteger));
                          oGantt.Selection.ShapeRange.Line.Visible := msoFalse;

                          // Verificar si es necesario mover la barra hacia delante
                          if oGantt.Cells[LocRen, ColBarra].Value < MryAvances.FieldByName('dIdFecha').AsDateTime then
                            oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(MryAvances.FieldByName('dIdFecha').AsDateTime) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));
                        end;


                    end
                    else
                    begin
                      // Toda la cantidad es excedente
                      if RadioGroup1.ItemIndex = 0 then
                      begin
                        FactorExcedente := MryAvances.FieldByName('dExcedente').AsFloat / MryAvances.FieldByName('dCantidadTotal').AsFloat;
                        // Localizar la columna en donde se debe generar la barra
                        ColBarra := rAncho + 1;
                        while oGantt.Cells[LocRen, ColBarra + 1].Value <= myInicio do
                          ColBarra := ColBarra + 1;
                        //ColBarra := (qAvances.FieldByName('dIdFecha').AsDateTime - pInicioP) + rAncho + 1;
                        if RadioButton1.Checked then
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value + MryAvances.FieldByName('dExcedente').AsFloat
                        else
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value := oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 3)].Value + FactorExcedente;

                        if ((IncDay(LastDate) = MryAvances.FieldByName('dIdFecha').AsDateTime) and (LastDate > 0)) And (LastColor = 2) then
                        begin
                          DurBarra := oGantt.Selection.ShapeRange.Width;
                          DurBarra := DurBarra + ((1 / (pTerminoP - pInicioP)) * wCol);
                          oGantt.Selection.ShapeRange.Width := DurBarra;
                        end
                        else
                        begin
                          DurBarra := (1 / (pTerminoP - pInicioP)) * wCol;

                          oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto)).Select;
                          oGantt.Selection.Cut;
                          oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                          oGantt.ActiveSheet.Paste;
                          oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);

                          oGantt.Selection.ShapeRange.Fill.Solid;
                          oGantt.Selection.ShapeRange.Fill.ForeColor.RGB{.SchemeColor} := ColorToRgb(tColor(MryAvances.FieldByName('iColor').AsInteger));
                          oGantt.Selection.ShapeRange.Line.Visible := msoFalse;

                          if oGantt.Cells[LocRen, ColBarra].Value < MryAvances.FieldByName('dIdFecha').AsDateTime then
                            oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) * (Trunc(MryAvances.FieldByName('dIdFecha').AsDateTime) - Trunc(oGantt.Cells[LocRen, ColBarra].Value)));
                        end;
                        LastColor := MryAvances.FieldByName('iColor').AsInteger;
                      end;
                      //*******************************************************************************
                    end;
                  end
                  else
                  begin
                    // Toda la cantidad es normal
                    // Poner avances diarios (Sin seleccionar la celda)
                    ColBarra := rAncho + 1;
                    while oGantt.Cells[LocRen, ColBarra ].Value < MryAvances.FieldByName('dIdFecha').AsDateTime do
                      ColBarra := ColBarra + 4;
                    //ColBarra := (qAvances.FieldByName('dIdFecha').AsDateTime - pInicioP) + rAncho + 1;

                    //iColHr:=trunc(Frac(MryAvances.FieldByName('sHoraInicio').AsFloat) * 4);

                    //ColBarra:=ColBarra+iColHr;

                    if MryAvances.FieldByName('sHoraInicio').AsFloat=0 then
                    begin
                      if MryAvances.FieldByName('HoraI').AsString='00:00' then
                        fHoraI:=0
                      else
                        fHoraI:=1;
                    end
                    else
                      fHoraI:=MryAvances.FieldByName('sHoraInicio').AsFloat;

                    Cuantos:=Trunc(ColBarra);

                    if MryAvances.FieldByName('sHoraFinal').AsFloat=0 then
                    begin
                      if MryAvances.FieldByName('HoraF').AsString='00:00' then
                        fHoraF:=0
                      else
                        fHoraF:=1;
                    end
                    else
                      fHoraF:=MryAvances.FieldByName('sHoraFinal').AsFloat;

                    //Cuantos := 0;
                    while (oGantt.Cells[LocRen, Cuantos + 1].Value <> 0) and (oGantt.Cells[LocRen, Cuantos + 1].Value < MryAvances.FieldByName('dIdFecha').AsFloat) do
                      Cuantos := Cuantos + 4;
                    Cuantos := Cuantos-(Trunc(ColBarra));

                    {fHoraF:=MryAvances.FieldByName('sHoraFinal').AsFloat;

                    if MryAvances.FieldByName('sHoraFinal').AsFloat=0 then
                      fHoraF:=1;



                    iColHr:=trunc(Frac(fHoraF - MryAvances.FieldByName('sHoraInicio').AsFloat) * 4);

                    Cuantos:=Cuantos +  iColHr;   }

                    begin
                      iColHr:=trunc(fHoraI * 4);
                      ColBarra:=ColBarra+iColHr;

                      {DurBarra := (1 / ((pTerminoP - pInicioP)*4) ) * wCol;
                      //AnchoColumnasPixels:= DurBarra;
                      DurBarra:=DurBarra *  ((fHoraF - MryAvances.FieldByName('sHoraInicio').AsFloat) * 4) ;}

                      Cuantos:=CUANTOS +  (trunc(fHoraF * 4) - trunc(fHoraI * 4));

                      EAux:=fHoraI * 24;
                      iAux:=Trunc(EAux);
                      fRestaI:=(iAux MOD 6) + Frac(EAux) ;
                      fRestaI:=fRestaI/6;

                      EAux:=fHoraF * 24;
                      iColHrF:=trunc(EAux);
                      pFactor:= (iColHrF MOD 6) + Frac(EAux);
                      pFactor:=pFactor/6;

                      //iColHr:=trunc(fHoraF * 4);
                      DurBarra:=(Cuantos * AnchoColumnasPixels)-(AnchoColumnasPixels * fRestaI)    ;
                      DurBarra:=DurBarra +(pFactor * AnchoColumnasPixels);

                      oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren+ 2)].Select;
                      oHoja.Shapes.AddShape(1{61},
                      oGantt.Selection.left,oGantt.Selection.top, DurBarra, (fAlto)).Select;


                      {oHoja.Shapes.AddShape(61, 0, 0, DurBarra, (fAlto)).Select;
                      oGantt.Selection.Cut;
                      oHoja.Range[ColumnaNombre(Trunc(ColBarra)) + IntToStr(Ren + 2)].Select;
                      oGantt.ActiveSheet.Paste;}
                      oGantt.Selection.ShapeRange.IncrementTop(fAlto / 4);

                      oGantt.Selection.ShapeRange.Fill.Solid;
                      oGantt.Selection.ShapeRange.Fill.ForeColor.RGB{.SchemeColor} := ColorToRgb(tColor(MryAvances.FieldByName('iColor').AsInteger));
                      oGantt.Selection.ShapeRange.Line.Visible := msotrue;
                      oGantt.Selection.ShapeRange.Line.Weight:=0.75;


                      {iColHr:=trunc(MryAvances.FieldByName('sHoraInicio').AsFloat * 24);
                      pFactor:= (iColHr MOD 6) ;
                      pFactor:=pFactor/6; }
                      EAux:=fHoraI * 24;
                      iAux:=Trunc(EAux);
                      fRestaI:=(iAux MOD 6) + Frac(EAux) ;
                      fRestaI:=fRestaI/6;

                      oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * fRestaI) );



                     // oGantt.Selection.ShapeRange.IncrementLeft((AnchoColumnasPixels * pFactor) );

                    end;
                    LastColor := MryAvances.FieldByName('iColor').AsInteger;
                  end;

                  LastDate := MryAvances.FieldByName('dIdFecha').AsDateTime;
                  dAVance:=Davance+ MryAvances.FieldByName('dAvance').AsFloat;
                  MryAvances.Next;
                end;




              end;

            end;

          end
          else
          begin
            // Darle formato a las columnas vacias cuando se trate de una grafica de gantt, dichas columnas van ocultas pero se les da formato por si el usuario los quiere usutiliar para otra cosa
            SetSelect(oHoja,'I' + Trim(IntToStr(Ren)) + ':I' + Trim(IntToStr(Ren + 3)));
            oGantt.Selection.MergeCells := True;

            SetSelect(oHoja,'L' + Trim(IntToStr(Ren)) + ':L' + Trim(IntToStr(Ren + 3)));
            oGantt.Selection.MergeCells := True;

            SetSelect(oHoja,'M' + Trim(IntToStr(Ren)) + ':M' + Trim(IntToStr(Ren + 3)));
            oGantt.Selection.MergeCells := True;

            SetSelect(oHoja,'N' + Trim(IntToStr(Ren)) + ':N' + Trim(IntToStr(Ren + 3)));
            oGantt.Selection.MergeCells := True;
          end;
          //* Terminar proceso de avances reales
          //*****************************************************************************************************************************

          //*****************************************************************************************************************************
          //* Inicia proceso de datos generales de partidas
          SetSelect(oHoja,'A' + Trim(Inttostr(Ren)) + ':A' + Trim(Inttostr(Ren + 3)));
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.Value := zrqPrograma.FieldByName('sNumeroActividad').asString;// + #10 + '(' + zrqPrograma.FieldByName('sWbs').asString + ')';
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.Font.Bold := True;
          {SetSelect(oHoja,'B' + Trim(Inttostr(Ren)) + ':B' + Trim(Inttostr(Ren + 3)));
          //oGantt.Selection.Value := zrqPrograma.FieldByName('sAnexo').asString;
          oGantt.Selection.MergeCells := True;  }

          SetSelect(oHoja,'B' + Trim(Inttostr(Ren)) + ':B' + Trim(Inttostr(Ren + 3)));
          oGantt.Selection.WrapText:=True;
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.VerticalAlignment := xlCenter ;
          oGantt.Selection.HorizontalAlignment :=XlJustify ;
          oGantt.Selection.NumberFormat:='@';
          oGantt.Selection.Value := zrqPrograma.FieldByName('mDescripcion').asString;
          if zrqPrograma.FieldByName('sTipoActividad').AsString = 'Paquete' then
            oGantt.Selection.Font.Bold := True;

          SetSelect(oHoja,'C' + Trim(Inttostr(Ren)) + ':C' + Trim(Inttostr(Ren + 1)));
          oGantt.Selection.NumberFormat:='@';
          oGantt.Selection.Value :=sSumaTotalHrsP; //(( )  - ) * 24;
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.Font.Bold := True;


          SetSelect(oHoja,'C' + Trim(Inttostr(Ren+2)) + ':C' + Trim(Inttostr(Ren + 3)));
          oGantt.Selection.NumberFormat:='@';
          oGantt.Selection.Value :=sSumaTotalHrs;//((zrqPrograma.FieldByName('dFechaFinal').AsFloat + 1)  - zrqPrograma.FieldByName('dFechaInicio').AsFloat) * 24;
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.Font.Bold := True;

          {sTotalP,sTotalR,sAux}
          //HoursBetween(zrqPrograma.FieldByName('dFechaInicio').AsFloat+ 1,zrqPrograma.FieldByName('dFechaFinal').AsFloat);

          {sAux:=IntToStr(HoursBetween(zrqPrograma.FieldByName('dFechaInicio').AsDateTime,zrqPrograma.FieldByName('dFechaFinal').AsDateTime))+':00';
          sTotalP:=sfnSumaHoras(sTotalP,sAux);  }




          if zrqPrograma.FieldByName('sTipoActividad').AsString <> 'Paquete' then
          Begin
            // Colocar formato condicional para actividades
            oGantt.Selection.FormatConditions.Delete;
            oGantt.Selection.FormatConditions.Add(xlCellValue, xlEqual, '1');
            oGantt.Selection.FormatConditions[1].Font.Bold := True;
            oGantt.Selection.FormatConditions[1].Interior.ColorIndex := 8;






          End;

          SetSelect(oHoja,'D' + Trim(Inttostr(Ren)) + ':D' + Trim(Inttostr(Ren + 1)));
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.Value :=tmpFechaIP; //FechaCadena(zrqPrograma.FieldByName('dFechaInicio').AsDateTime,0);
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.WrapText:=True;
          oGantt.Selection.Font.Bold := True;

          SetSelect(oHoja,'D' + Trim(Inttostr(Ren+2)) + ':D' + Trim(Inttostr(Ren + 3)));
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.Value :=tmpFechaI;//oGantt.Selection.Value := FechaCadena(zrqPrograma.FieldByName('dFechaInicio').AsDateTime,0);
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.WrapText:=True;
          oGantt.Selection.Font.Bold := True;

          SetSelect(oHoja,'E' + Trim(Inttostr(Ren)) + ':E' + Trim(Inttostr(Ren + 1)));
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.WrapText:=True;
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.Value :=TmpFechaTP; //FechaCadena(zrqPrograma.FieldByName('dFechaFinal').AsDateTime,1);
          oGantt.Selection.Font.Bold := True;

           SetSelect(oHoja,'E' + Trim(Inttostr(Ren+2)) + ':E' + Trim(Inttostr(Ren + 3)));
          oGantt.Selection.MergeCells := True;
          oGantt.Selection.WrapText:=True;
          oGantt.Selection.VerticalAlignment := xlCenter;
          oGantt.Selection.HorizontalAlignment := xlCenter;
          oGantt.Selection.Font.Bold := True;
          oGantt.Selection.Value :=TmpFechaT;

          if zrqPrograma.FieldByName('sTipoActividad').AsString = 'Paquete' then
          begin

            Connection.QryBusca.Active := False;
            Connection.QryBusca.Sql.Clear;
            Connection.QryBusca.Sql.Add('Select a.dPonderado, if((select ba.lCancelada from bitacoradeactividades ba where a.sContrato = ba.sContrato and a.sNumeroOrden = ba.sNumeroOrden and ba.sIdTipoMovimiento="ED" ' +
                                        'and a.swbs = ba.swbs and lCancelada = "Si" limit 1) ="Si",100,sum(b.dAvance)) as dAvance, if(sum(b.davance) > 100, a.dPonderado, if((select ba.lCancelada from bitacoradeactividades ba ' +
                                        'where a.sContrato = ba.sContrato and a.sNumeroOrden = ba.sNumeroOrden and ba.sIdTipoMovimiento="ED" and a.swbs = ba.swbs and lCancelada = "Si" and ba.didfecha <= :fecha limit 1) ="Si",a.dPonderado, ' +
                                        'sum(b.davance * (a.dPonderado))))as dAvancePonderado ' +
                                        'From actividadesxorden a inner join bitacoradeactividades b on (b.scontrato = a.scontrato and b.snumeroorden = a.snumeroorden and b.swbs = a.swbs and b.didfecha <= :fecha) ' +
                                        'Where a.sContrato = :contrato and a.sIdConvenio =:Convenio and a.sNumeroOrden = :orden and b.sIdTipoMovimiento="ED" And b.sWbs like concat(:wbs, ".%") group by a.swbs');

            Connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
            Connection.QryBusca.ParamByName('convenio').AsString := global_convenio;
            Connection.QryBusca.ParamByName('orden').AsString    := tsNumeroOrden.KeyValue;
            Connection.QryBusca.ParamByName('wbs').AsString      := zrqPrograma.FieldValues['swbs'];
            Connection.QryBusca.ParamByName('fecha').AsDate      := tdIdFecha.Date;
            Connection.QryBusca.Open;

            eAvance := 0;
            while Not Connection.QryBusca.Eof do
            begin
              eAvance := eAvance + Connection.QryBusca.FieldByName('dAvancePonderado').AsFloat;
              Connection.QryBusca.Next;
            end;
            Connection.QryBusca.Close;

            //Value := eAvance;




            SetSelect(oHoja,ColumnaNombre(ColEnd+2) + Trim(Inttostr(Ren)) + ':' +ColumnaNombre(ColEnd+2) + Trim(Inttostr(Ren+3)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Value := 0;

            SetSelect(oHoja,ColumnaNombre(ColEnd+3) + Trim(Inttostr(Ren)) + ':' +ColumnaNombre(ColEnd+3) + Trim(Inttostr(Ren+3)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
            oGantt.Selection.Value := (eAVance/100);
          end
          else
          Begin

            sTotalR:=sfnSumaHoras(sTotalR,sSumaTotalHrs);
            sTotalP:=sfnSumaHoras(sTotalP,sSumaTotalHrsP);

            SetSelect(oHoja,ColumnaNombre(ColEnd+1) + Trim(Inttostr(Ren)) + ':' +ColumnaNombre(ColEnd+1) + Trim(Inttostr(Ren+1)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
            oGantt.Selection.Value := zrqPrograma.FieldByName('dPonderado').AsFloat/100;


            SetSelect(oHoja,ColumnaNombre(ColEnd+2) + Trim(Inttostr(Ren)) + ':' +ColumnaNombre(ColEnd+2) + Trim(Inttostr(Ren+1)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
            oGantt.Selection.Value := 0;

             SetSelect(oHoja,ColumnaNombre(ColEnd+3) + Trim(Inttostr(Ren)) + ':' +ColumnaNombre(ColEnd+3) + Trim(Inttostr(Ren+1)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
             oGantt.Selection.Value :=((dAVance * zrqPrograma.FieldByName('dPonderado').AsFloat)/100);


            SetSelect(oHoja,ColumnaNombre(ColEnd+1) + Trim(Inttostr(Ren+2)) + ':' +ColumnaNombre(ColEnd+1) + Trim(Inttostr(Ren+3)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
            oGantt.Selection.Value :='PARCIAL';
            oGantt.Selection.Interior.Color:= 65535;//colorIndex := 6;
            oGantt.Selection.Interior.Pattern := xlSolid;

            SetSelect(oHoja,ColumnaNombre(ColEnd+2) + Trim(Inttostr(Ren+2)) + ':' +ColumnaNombre(ColEnd+2) + Trim(Inttostr(Ren+3)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Value := 0;
            oGantt.Selection.Interior.Color:= 65535;//colorIndex := 6;
            oGantt.Selection.Interior.Pattern := xlSolid;


            SetSelect(oHoja,ColumnaNombre(ColEnd+3) + Trim(Inttostr(Ren+2)) + ':' +ColumnaNombre(ColEnd+3) + Trim(Inttostr(Ren+3)));
            oGantt.Selection.MergeCells := True;
            oGantt.Selection.VerticalAlignment := xlCenter;
            oGantt.Selection.HorizontalAlignment := xlCenter;
            oGantt.Selection.Font.Name:='Arial';
            oGantt.Selection.Font.Size := 45;
            oGantt.Selection.Value := dAVance;
            oGantt.Selection.Interior.Color:= 65535;//colorIndex := 6;
            oGantt.Selection.Interior.Pattern := xlSolid;


          End;

          //* Terminar proceso de datos generales de partidas
          //*****************************************************************************************************************************

          inc(Ren, 4);
        end;



        Inc(Cuenta);
        zrqPrograma.Next;

        // Mostrar los datos del registro generado para que el usuario aprecie la generaci�n
        if CheckBox1.Checked and Graficar then
        Begin
          oGantt.ScreenUpdating := True;
          SetSelect(oHoja, 'A' + Trim(IntToStr(Ren)));
          oGantt.ScreenUpdating := False;
        End;
      End;

      if cbTipoGrafica.ItemIndex=2 then
        aCol:=SaltoCol;
      // Eliminar fechas de control final
      {oGantt.Range[ColumnaNombre(aCol + rAncho + 2) + '12:' + ColumnaNombre(aCol + rAncho + 4) + '15'].Select;
      oGantt.Selection.Clear;}

      // Generar las l�neas del �rea de gr�fica
      GeneraLineas(oGantt,'A10:' + Letra(aCol + rAncho + 3) + Trim(InttoStr(Ren - 1)));
      //FormatoDatos(oGantt,Ren - 1);

      SetSelect(oHoja, Letra(rAncho + 1) + '15:' + Letra(aCol + rAncho + 3) + Trim(IntToStr(Ren - 1)));
      if RadioButton1.Checked then
        oGantt.Selection.NumberFormat := '0.0000'
      else
        oGantt.Selection.NumberFormat := '0.0000%';

      oHoja.Rows[inttostr(ren)+':'+inttostr(ren)].RowHeight :=30;
      oHoja.Rows[inttostr(ren+1)+':'+inttostr(ren+3)].RowHeight :=100;
      oHoja.Rows[inttostr(ren+4)+':'+inttostr(ren+4)].RowHeight :=30;
      // Bajar dos renglones para posicionar correctamente el c�digo de colores
      Alto := Alto + RowHeight(oGantt,Ren);
      iRenFinaly:=Ren;
      Inc(Ren);
      Alto := Alto + RowHeight(oGantt,Ren);


      SetSelect(oHoja,'B' + Trim(IntToStr(Ren)) + ':C' + Trim(IntToStr(Ren)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlCenter;
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 44;
      oGantt.Selection.Font.Bold:=True;
      oGantt.Selection.Value := 'RESUMEN';

      inc(ren);
      SetSelect(oHoja,'B' + Trim(IntToStr(Ren)) + ':B' + Trim(IntToStr(Ren)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlRight;
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 44;
      oGantt.Selection.Font.Bold:=True;
      oGantt.Selection.Value := 'TOTAL  DE TIEMPO PROGRAMADO DE LAS ACTIVIDADES';

      SetSelect(oHoja,'C' + Trim(IntToStr(Ren)) + ':C' + Trim(IntToStr(Ren)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlCenter;
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 44;
      oGantt.Selection.Font.Bold:=True;
      oGantt.Selection.Value := sTotalP + ' Hrs.' ;

      inc(ren);

      SetSelect(oHoja,'B' + Trim(IntToStr(Ren)) + ':B' + Trim(IntToStr(Ren)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlRight;
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 44;
      oGantt.Selection.Font.Bold:=True;
      oGantt.Selection.Value := 'TOTAL DE TIEMPO REAL DE LAS ACTIVIDADES';

      SetSelect(oHoja,'C' + Trim(IntToStr(Ren)) + ':C' + Trim(IntToStr(Ren)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlCenter;
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 44;
      oGantt.Selection.Font.Bold:=True;
      oGantt.Selection.Value := sTotalR + ' Hrs.' ;

      GeneraLineas(oGantt,'B'+Trim(InttoStr(Ren - 2))+':C' + Trim(InttoStr(Ren)));


      SetSelect(oHoja,ColumnaNOmbre(ColEnd+1)+'10:'+ColumnaNOmbre(ColEnd+3)+'13');
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 48;
      oGantt.Selection.Value := 'AVANCES';
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlCenter;

      SetSelect(oHoja,ColumnaNOmbre(ColEnd+1)+'14:'+ColumnaNOmbre(ColEnd+1)+'14');
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 48;
      oGantt.Selection.Value := 'POND.';
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlCenter;

      SetSelect(oHoja,ColumnaNOmbre(ColEnd+2)+'14:'+ColumnaNOmbre(ColEnd+2)+'14');
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 48;
      oGantt.Selection.Value := 'PROG.';
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlCenter;

      SetSelect(oHoja,ColumnaNOmbre(ColEnd+3)+'14:'+ColumnaNOmbre(ColEnd+3)+'14');
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 48;
      oGantt.Selection.Value := 'REAL';
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlCenter;

      SetColumnWidth(oGantt,ColEnd+1,ColEnd+3,56.89);

      AnchoPage:=(SaltoCol * AnchoColumnasPixels);//+rAncho;
      AnchoPage:=AnchoPage + (oGantt.Columns[Letra(1)].Width+oGantt.Columns[Letra(2)].Width+oGantt.Columns[Letra(3)].Width+oGantt.Columns[Letra(4)].Width+
      oGantt.Columns[Letra(5)].Width+oGantt.Columns[Letra(5)].Width);
      AnchoPage:=AnchoPage +(oGantt.Columns[Letra(ColEnd+1)].Width * 3);
      BloquePage:=((AnchoPage-10)/6);


      //oGantt.Columns[Letra(rAncho + 1)]
      Inc(ren,2);
      //544.2
      oHoja.Shapes.AddShape(61, 0, 0, AnchoA - 50, (fAlto-(fAlto/4))).Select;
      sleep(500);
      oGantt.Selection.Cut;
      sleep(500);
      oHoja.Range['A' + IntToStr(Ren)].Select;
      sleep(500);
      oGantt.ActiveSheet.Paste;

      oGantt.Selection.ShapeRange.Fill.Solid;
      oGantt.Selection.ShapeRange.Fill.ForeColor.RGB:=RGB(7, 30, 169);//ForeColor.SchemeColor := 0;
      oGantt.Selection.ShapeRange.Line.Visible := msotrue;
      oGantt.Selection.ShapeRange.Line.Weight:=0.75;
      oGantt.Selection.ShapeRange.IncrementLeft(BloquePage/25);

      oHoja.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 700, (fAlto / 2)).Select;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Text:='PROGRAMADO';
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.FirstLineIndent := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.Alignment := msoAlignLeft;
      oGantt.Selection.ShapeRange.TextFrame2.VerticalAnchor := msoAnchorMiddle;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameComplexScript := 'Arial';
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameFarEast := '+mn-ea';
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Visible := msoTrue;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.ObjectThemeColor := msoThemeColorDark1;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.TintAndShade := 0;
      if (oGantt.Application.version >= 14) then
        oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.Brightness := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Transparency := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Solid;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Size := 40;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Name := 'Arial';
      oGantt.Selection.ShapeRange.Fill.Visible:= msoFalse;
      oGantt.Selection.ShapeRange.Line.Visible:= msoFalse;

      sleep(500);
      oGantt.Selection.Cut;
      sleep(500);
      oHoja.Range['A' + IntToStr(Ren)].Select;
      sleep(500);
      oGantt.ActiveSheet.Paste;
      oGantt.Selection.ShapeRange.IncrementLeft(BloquePage/4);
      oGantt.Selection.ShapeRange.incrementTop((fAlto / 4)-(fAlto / 10)  );

      PosCol:=1;

      QrMovimientos:=TzReadOnlyQuery.Create(nil);
      QrMovimientos.Connection:=Connection.zConnection;
      QrMovimientos.SQL.Text:='select * from tiposdemovimiento where sContrato=:Contrato and sClasificacion="Tiempo de Actividades" and lGrafica="Si"';
      QrMovimientos.ParamByName('Contrato').AsString:=Global_Contrato_Barco;
      QrMovimientos.Open;
      //PosCol:=2;
      while not QrMovimientos.Eof do
      begin
        if PosCol=6 then
        Begin
          PosCol:=0;
          Inc(ren,3);
          oHoja.Rows[inttostr(ren-1)+':'+inttostr(ren-1)].RowHeight :=20;
          oHoja.Shapes.AddShape(61, 0, 0, AnchoA - 50, (fAlto-(fAlto/4))).Select;
          sleep(500);
          oGantt.Selection.Cut;
          sleep(500);
          oHoja.Range['A' + IntToStr(Ren)].Select;
          sleep(500);
          oGantt.ActiveSheet.Paste;

          oGantt.Selection.ShapeRange.Fill.Solid;
          oGantt.Selection.ShapeRange.Fill.ForeColor.RGB:=ColorToRgb(tColor(QrMovimientos.FieldByName('iColor').AsInteger));
          oGantt.Selection.ShapeRange.Line.Visible := msotrue;
          oGantt.Selection.ShapeRange.Line.Weight:=0.75;
          oGantt.Selection.ShapeRange.IncrementLeft(BloquePage/25);

          oHoja.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 700, (fAlto / 2)).Select;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Text:=QrMovimientos.FieldByName('sDescripcion').AsString;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.FirstLineIndent := 0;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.Alignment := msoAlignLeft;
          oGantt.Selection.ShapeRange.TextFrame2.VerticalAnchor := msoAnchorMiddle;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameComplexScript := 'Arial';
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameFarEast := '+mn-ea';
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Visible := msoTrue;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.ObjectThemeColor := msoThemeColorDark1;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.TintAndShade := 0;
          if (oGantt.Application.version >= 14) then
            oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.Brightness := 0;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Transparency := 0;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Solid;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Size := 40;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Name := 'Arial';
          oGantt.Selection.ShapeRange.Fill.Visible:= msoFalse;
          oGantt.Selection.ShapeRange.Line.Visible:= msoFalse;

          sleep(500);
          oGantt.Selection.Cut;
          sleep(500);
          oHoja.Range['A' + IntToStr(Ren)].Select;
          sleep(500);
          oGantt.ActiveSheet.Paste;
          oGantt.Selection.ShapeRange.IncrementLeft(BloquePage/4);
          oGantt.Selection.ShapeRange.incrementTop((fAlto / 4)-(fAlto / 10));

        End
        else
        begin

          oHoja.Shapes.AddShape(61, 0, 0, AnchoA - 50, (fAlto-(fAlto/4))).Select;
          sleep(500);
          oGantt.Selection.Cut;
          sleep(500);
          oHoja.Range['A' + IntToStr(Ren)].Select;
          sleep(500);
          oGantt.ActiveSheet.Paste;

          oGantt.Selection.ShapeRange.Fill.Solid;
          oGantt.Selection.ShapeRange.Fill.ForeColor.RGB:=ColorToRgb(tColor(QrMovimientos.FieldByName('iColor').AsInteger));
          oGantt.Selection.ShapeRange.Line.Visible := msotrue;
          oGantt.Selection.ShapeRange.Line.Weight:=0.75;
          oGantt.Selection.ShapeRange.IncrementLeft(BloquePage * PosCol);//(BloquePage/25);

          oHoja.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 700, (fAlto / 2)).Select;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Text:=QrMovimientos.FieldByName('sDescripcion').AsString;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.FirstLineIndent := 0;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.Alignment := msoAlignLeft;
          oGantt.Selection.ShapeRange.TextFrame2.VerticalAnchor := msoAnchorMiddle;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameComplexScript := 'Arial';
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameFarEast := '+mn-ea';
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Visible := msoTrue;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.ObjectThemeColor := msoThemeColorDark1;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.TintAndShade := 0;
          if (oGantt.Application.version >= 14) then
            oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.Brightness := 0;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Transparency := 0;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Solid;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Size := 40;
          oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Name := 'Arial';
          oGantt.Selection.ShapeRange.Fill.Visible:= msoFalse;
          oGantt.Selection.ShapeRange.Line.Visible:= msoFalse;

          sleep(500);
          oGantt.Selection.Cut;
          sleep(500);
          oHoja.Range['A' + IntToStr(Ren)].Select;
          sleep(500);
          oGantt.ActiveSheet.Paste;
          oGantt.Selection.ShapeRange.IncrementLeft((BloquePage * PosCol) + (BloquePage/4));
          oGantt.Selection.ShapeRange.incrementTop((fAlto / 4)-(fAlto / 10));
        end;

        Inc(PosCol);
        QrMovimientos.Next;
      end;



      Inc(ren,3);
      QrMovimientos.Active:=false;
      QrMovimientos.SQL.Text:='select ba.* from bitacoradeactividades ba inner join tiposdemovimiento tm' + #10 +
                              'on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=ba.sIdTipoMovimiento)' + #10 + 
                              'where ba.sContrato=:Contrato and ba.dIdFecha between :Inicio and :Final' + #10 + 
                              'and ba.sNumeroOrden=:Orden and (tm.sClasificacion="Notas" or (ba.sIdTipoMovimiento="ED" and ba.sIdClasificacion<>"TE")) '+
                              'order by ba.dIdFecha,ba.sIdTipoMovimiento desc';
      QrMovimientos.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
      QrMovimientos.ParamByName('Contrato').AsString:=Global_Contrato;
      QrMovimientos.ParamByName('Orden').AsString:=tsNumeroOrden.KeyValue;
      QrMovimientos.ParamByName('Inicio').AsDate:=tdFechaInicio.Date;
      QrMovimientos.ParamByName('Final').AsDate:=tdIdFecha.Date;
      QrMovimientos.Open;

      //SetSelect(oHoja,ColumnaNOmbre(ColEnd+3)+'14:'+ColumnaNOmbre(ColEnd+3)+'14');
      SetSelect(oHoja,ColumnaNOmbre(1)+ Inttostr(ren)+':'+ColumnaNOmbre(3)+ Inttostr(ren));
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 45;
      oGantt.Selection.Font.bold:=true;
      oGantt.Selection.Value := 'NOTAS:';
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlLeft;
      SetRowHeight(oGantt,ren,88.5);

      Inc(ren);
      SetSelect(oHoja,ColumnaNOmbre(1)+ Inttostr(ren)+':'+ColumnaNOmbre(ColEnd+3)+ Inttostr(ren));
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 45;
      oGantt.Selection.Font.bold:=true;
      oGantt.Selection.Value := '*** ESTE PROGRAMA CONSIDERA JORNADAS LABORALES DE 12:00 HRS.';
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlLeft;
      SetRowHeight(oGantt,ren,57);

      Inc(ren);
      SetSelect(oHoja,ColumnaNOmbre(1)+ Inttostr(ren)+':'+ColumnaNOmbre(ColEnd+3)+ Inttostr(ren));
      oGantt.Selection.Font.Name:='Arial';
      oGantt.Selection.Font.Size := 45;
      oGantt.Selection.Font.bold:=true;
      oGantt.Selection.Value := '*** PARA ESTE PROGRAMA SE CONSIDERAN QUE TODOS LOS MATERIALES DE INSUMO Y HERRAMIENTA NECESARIA, SE ENCUENTRAN EN SITIO DE TRABAJO AL 100%.';
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.VerticalAlignment := xlCenter;
      oGantt.Selection.HorizontalAlignment := xlLeft;
      SetRowHeight(oGantt,ren,57);
      Inc(ren);
      while Not QrMovimientos.Eof do
      Begin
        Inc(ren);

        oHoja.Range[ColumnaNOmbre(1)+ Inttostr(ren)+':'+ColumnaNOmbre(ColEnd+3)+ Inttostr(ren)].Select;
        oGantt.Selection.MergeCells := true;

        oGantt.Selection.NumberFormat := '@' ;
        oGantt.Selection.Font.Size := 50;
        oGantt.Selection.Font.Name := 'Arial';
         oGantt.Selection.VerticalAlignment := xlCenter;
        oGantt.Selection.HorizontalAlignment := xlLeft;
        oGantt.Selection.wraptext:=true;
        oGantt.Selection.Value :='NOTA ' + Trim(QrMovimientos.FieldByName('mDescripcion').AsString);
        rangoE:=oGantt.Range[ColumnaNOmbre(1)+ Inttostr(ren)+':'+ColumnaNOmbre(ColEnd+3)+ Inttostr(ren)];
        AjustarTexto(rangoE,ColEnd+3);
        //oGantt.Rows[ren].AutoFit;
        QrMovimientos.Next;
      End;


      Freeandnil(QrMovimientos);

      Inc(Ren,6);
      rDiarioFirmas(global_contrato,tsNumeroOrden.KeyValue,global_turno,tdIdFecha.Date,self);

      oHoja.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 1400, (fAlto*2)).Select;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Text:='____________________________________________________'#13 + sSuperintendente + #13 + sPuestoSuperintendente   ;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.FirstLineIndent := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.Alignment := msoAlignCenter;
      //oGantt.Selection.ShapeRange.TextFrame2.VerticalAnchor := msoAnchorMiddle;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameComplexScript := 'Arial';
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameFarEast := '+mn-ea';
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Visible := msoTrue;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.ObjectThemeColor := msoThemeColorDark1;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.TintAndShade := 0;
      if (oGantt.Application.version >= 14) then
        oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.Brightness := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Transparency := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Solid;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Size := 45;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Name := 'Arial';
      oGantt.Selection.ShapeRange.Fill.Visible:= msoFalse;
      oGantt.Selection.ShapeRange.Line.Visible:= msoFalse;

      sleep(500);
      oGantt.Selection.Cut;
      sleep(500);
      oHoja.Range['A' + IntToStr(Ren)].Select;
      sleep(500);
      oGantt.ActiveSheet.Paste;
      oGantt.Selection.ShapeRange.IncrementLeft(BloquePage/4);
      oGantt.Selection.ShapeRange.incrementTop((fAlto));

          oHoja.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 1400, (fAlto*2)).Select;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Text:='____________________________________________________'#13 + sSupervisor + #13 + sPuestoSupervisor    ;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.FirstLineIndent := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.ParagraphFormat.Alignment := msoAlignCenter;
      //oGantt.Selection.ShapeRange.TextFrame2.VerticalAnchor := msoAnchorMiddle;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameComplexScript := 'Arial';
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.NameFarEast := '+mn-ea';
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Visible := msoTrue;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.ObjectThemeColor := msoThemeColorDark1;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.TintAndShade := 0;
      if (oGantt.Application.version >= 14) then
        oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.ForeColor.Brightness := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Transparency := 0;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Fill.Solid;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Size := 45;
      oGantt.Selection.ShapeRange.TextFrame2.TextRange.Characters.Font.Name := 'Arial';
      oGantt.Selection.ShapeRange.Fill.Visible:= msoFalse;
      oGantt.Selection.ShapeRange.Line.Visible:= msoFalse;

      sleep(500);
      oGantt.Selection.Cut;
      sleep(500);
      oHoja.Range['A' + IntToStr(Ren)].Select;
      sleep(500);
      oGantt.ActiveSheet.Paste;
      oGantt.Selection.ShapeRange.IncrementLeft(BloquePage * 4);
      oGantt.Selection.ShapeRange.incrementTop((fAlto));

      // Pintar las lineas separadoras de actividades
      Cuenta := 15;
      while Cuenta < iRenFinaly do//12 do
      begin
        oHoja.Range['A' + IntToStr(Cuenta) + ':' + Letra(Acol + rAncho +3) +IntToStr(Cuenta + 1)].Select;
        oGantt.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
        oGantt.Selection.Borders[xlEdgeTop].Weight := xlMedium;
        oGantt.Selection.Borders[xlEdgeTop].ColorIndex := 0;
        oGantt.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
        oGantt.Selection.Borders[xlEdgeBottom].Weight := xlMedium;
        oGantt.Selection.Borders[xlEdgeBottom].ColorIndex := 0;

        oHoja.Range['C' + IntToStr(Cuenta+2) + ':' + Letra(Acol + rAncho) +IntToStr(Cuenta + 3)].Select;
        oGantt.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
        oGantt.Selection.Borders[xlEdgeTop].Weight := xlMedium;
        oGantt.Selection.Borders[xlEdgeTop].ColorIndex := 0;
        oGantt.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
        oGantt.Selection.Borders[xlEdgeBottom].Weight := xlMedium;
        oGantt.Selection.Borders[xlEdgeBottom].ColorIndex := 0;

        Alto := Alto + 2;
        Inc(Cuenta,4);

        // Mostrar los datos del registro generado para que el usuario aprecie la generaci�n
        if CheckBox1.Checked and (Cuenta Mod 2 = 0)then
        Begin
          oGantt.ScreenUpdating := True;
          SetSelect(oHoja, 'A' + Trim(IntToStr(Cuenta)));
      //    oGantt.ScreenUpdating := False;
        End;
      end;

      // SetSelect(oHoja, 'A' );
      oHoja.Rows['10:11'].RowHeight :=50;
      oHoja.Rows['12:12'].RowHeight :=0;
      //     Rows("12:12").Select
      //oHoja.Range['F12'].Activate;
      //Range("F12").Activate
      //Selection.EntireRow.Hidden = True
      //oGantt.Selection.EntireRow.Hidden := True;


      ProgressBar1.Visible := False;
      if CheckBox1.Checked then
        oGantt.ScreenUpdating := True;
      oGantt.ActiveWindow.Zoom := 20;

     ConfigurarHoja(oGantt,oHoja);
     PFormatosExcel_EstablecerAreaDeImpresion(oHoja,1,1,ColEnd+3,Ren+5);
  
      // Grabar el archivo generado
      Reintentar := True;
      while Reintentar do
      Try
        oGantt.ActiveWorkbook.SaveAs(fName);
        Reintentar := False;
      Except
        Reintentar := MessageDlg('No fu� posible grabar el archivo.' + #10 + #10 + 'Es probable que otro archivo con el mismo nombre y ubicaci�n se encuentre abierto, cierre el archivo abierto y vuelva a intentar.' + #10 + #10 + 'Si el error persiste entonces es posible que la unidad destino del archivo se encuentre llena o pretegida contra escritura, verifique esto e intente nuevamente.', mtWarning, [mbOk, mbCancel], 0) = mrOk;
      End;

      if Not CheckBox1.Checked then
        if MessageDlg('�Desea ver ahora el archivo que fu� generado?', mtConfirmation, mbYesNo, 0) = mrYes then
        Begin
       
          oGantt.Visible := True;
          oGantt.ScreenUpdating := True;
          oGantt.DisplayAlerts := True;
        End
        Else
          oGantt.Quit;
      oGantt := Unassigned;

    except
      on e:exception do
      begin

        oGantt.ScreenUpdating := True;
        oGantt.Visible := True;
        if Mostrar then
        begin
          oGantt.Quit;
          messagedlg('Ha ocurrido un error al generar el archivo de excel.' + #13 + #13 + e.Message, mtWarning, [mbOk], 0);
        end
        else
          messagedlg('Ha ocurrido un error al generar el archivo de excel.', mtWarning, [mbOk], 0);
      end;
    end;
  finally
    PnlProcesando.Visible := False; 
  end;
End;


procedure TfrmGraficaGerencialDX.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  connection.zCommand.SQL.text := 'drop temporary table if exists tempgantt';
  connection.zcommand.ExecSQL;
  action := cafree ;
end;

procedure TfrmGraficaGerencialDX.tdIdFechaKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 then
        tsNumeroOrden.SetFocus 
end;

procedure TfrmGraficaGerencialDX.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmGraficaGerencialDX.tsNumeroOrdenExit(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_salida ;
end;

procedure TfrmGraficaGerencialDX.tdFechaInicioEnter(Sender: TObject);
begin
    tdFechaInicio.Color := global_color_entrada
end;

procedure TfrmGraficaGerencialDX.tdFechaInicioExit(Sender: TObject);
begin
    tdFechaInicio.Color := global_color_salida
end;

procedure TfrmGraficaGerencialDX.tdFechaInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tdIdFecha.SetFocus 
end;

procedure TfrmGraficaGerencialDX.tdIdFechaEnter(Sender: TObject);
begin
    tdIdFecha.Color := global_color_entrada
end;

procedure TfrmGraficaGerencialDX.tdIdFechaExit(Sender: TObject);
begin
    tdIdFecha.Color := global_color_salida
end;


procedure TfrmGraficaGerencialDX.Button1Click(Sender: TObject);
var
  reportediario : tzReadOnlyQuery;

begin
  reportediario := tzReadOnlyQuery.Create(frmGraficaGerencialDX);
  reportediario.Connection := connection.zconnection;
  reportediario.Active := False;
  reportediario.SQL.Clear;
  reportediario.SQL.Add('delete from ganttavances');
  reportediario.ExecSQL;
  reportediario.Active := False;
  reportediario.SQL.Clear;

  if Opciones.ItemIndex = 0 then
  begin
     ReporteDiario.Active := False;
     ReporteDiario.SQL.Clear;
//     ReporteDiario.SQL.Add('select * from reportediario where scontrato = :contrato and sidconvenio = :convenio  and didfecha <= :fecha group by didfecha');
     ReporteDiario.SQL.Add('select * from reportediario where scontrato = :contrato and didfecha <= :fecha group by didfecha');
     ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
     ReporteDiario.Params.ParamByName('contrato').Value := global_contrato;
//     ReporteDiario.Params.ParamByName('convenio').DataType := ftString;
//     ReporteDiario.Params.ParamByName('convenio').Value := global_convenio;
     ReporteDiario.Params.ParamByName('fecha').Value := self.tdIdFecha.datetime;

  end
  else
  begin
     ReporteDiario.Active := False;
     ReporteDiario.SQL.Clear;
//     ReporteDiario.SQL.Add('select * from reportediario where scontrato = :contrato and sidconvenio = :convenio and snumeroorden = :orden and didfecha <= :fecha ');
     ReporteDiario.SQL.Add('select * from reportediario where scontrato = :contrato and snumeroorden = :orden and didfecha <= :fecha ');
     ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
     ReporteDiario.Params.ParamByName('contrato').Value := global_contrato;
//     ReporteDiario.Params.ParamByName('convenio').DataType := ftString;
//     ReporteDiario.Params.ParamByName('convenio').Value := global_convenio;
     ReporteDiario.Params.ParamByName('orden').DataType := ftString;
     ReporteDiario.Params.ParamByName('orden').Value := ordenesdetrabajo.FieldValues['snumeroorden'];
     ReporteDiario.Params.ParamByName('fecha').Value := self.tdIdFecha.datetime;

  end;

  ReporteDiario.Open;
//  showmessage(datetostr(ReporteDiario.FieldValues['didfecha']));
  while Not reportediario.Eof do
  begin
//     cfnCalculaAvances1(ReporteDiario.FieldValues['sContrato'], ReporteDiario.FieldValues['sNumeroOrden'], ReporteDiario.FieldValues['sidconvenio'] , ReporteDiario.FieldValues['sIdTurno'] , True ,ReporteDiario.FieldValues['didfecha'] , 'Avanzada' , frmGraficaGerencialDX);
     cfnCalculaAvances2(ReporteDiario.FieldValues['sContrato'], ReporteDiario.FieldValues['sNumeroOrden'], ReporteDiario.FieldValues['sidconvenio'] , ReporteDiario.FieldValues['sIdTurno'] , True ,ReporteDiario.FieldValues['didfecha'] , 'Avanzada' , frmGraficaGerencialDX);

     reportediario.next;
  end;

end;





function TfrmGraficaGerencialDX.cfnCalculaAvances2(sParamContrato, sParamOrden, sParamConvenio, sParamTurno: string; lParamMultiple: Boolean; dParamFecha: tDate; sParamSeguridad: string; tOrigen: TComponent): Currency;
var
  QryPlataforma: tzReadOnlyQuery;
  QryContrato, QryAuxiliar, QryReporteDiario : tzReadOnlyQuery;
  dAvancePlataforma: Currency;
  dAvanceObra: Currency;
  dAvancePartida: Currency;
  lContinua: Boolean;
begin
//  showmessage('entro al programa');
  dAvancePlataforma := 0;
  dAvanceObra := 0;
  dAvancePartida := 0;

  QryPlataforma := tzReadOnlyQuery.Create(tOrigen);
  QryPlataforma.Connection := connection.zconnection;

  QryContrato := tzReadOnlyQuery.Create(tOrigen);
  QryContrato.Connection := connection.zconnection;

  QryAuxiliar := tzReadOnlyQuery.Create(tOrigen);
  QryAuxiliar.Connection := connection.zconnection;

  QryReporteDiario := tzReadOnlyQuery.Create(tOrigen);
  QryReporteDiario.Connection := connection.zconnection;

  if Opciones.ItemIndex <> 0 then
    lcontinua := true
  else
    lcontinua := false;


  if lContinua then
  begin
        // Partidas canceladas ....
        // Avance por Plataforma ...
    QryPlataforma.Active := False;
    QryPlataforma.SQL.Clear;
    QryPlataforma.SQL.Add('Select Sum(( b.dAvance * a.dPonderado) / 100 ) as dAvance From actividadescanceladas b ' +
      'INNER JOIN actividadesxorden a ON (b.sContrato = a.sContrato And b.sNumeroOrden = a.sNumeroOrden And ' +
      'a.sIdConvenio = :Convenio And b.sWbs = a.sWbs And b.sNumeroActividad = a.sNumeroActividad) ' +
      'Where b.sContrato = :contrato and b.sNumeroOrden = :Orden and b.dIdFecha = :Fecha group by b.sNumeroOrden');
    QryPlataforma.Params.ParamByName('contrato').DataType := ftString;
    QryPlataforma.Params.ParamByName('contrato').Value := sParamContrato;
    QryPlataforma.Params.ParamByName('orden').DataType := ftString;
    QryPlataforma.Params.ParamByName('orden').Value := sParamOrden;
    QryPlataforma.Params.ParamByName('convenio').DataType := ftString;
    QryPlataforma.Params.ParamByName('convenio').Value := sParamConvenio;
    QryPlataforma.Params.ParamByName('Fecha').DataType := ftDate;
    QryPlataforma.Params.ParamByName('Fecha').Value := dParamFecha;
    QryPlataforma.Open;
    if QryPlataforma.RecordCount > 0 then
    begin
       showmessage('existen partidas con cancelaciones');
       dAvancePlataforma := dAvancePlataforma + QryPlataforma.FieldValues['dAvance'];

    end;





        // Avance de Partidas Reportadas ....
            // por Plataforma
    QryPlataforma.Active := False;
    QryPlataforma.SQL.Clear;
    QryPlataforma.SQL.Add('Select (( b.dAvance * a.dPonderado) / 100 )  as dAvance, b.snumeroactividad, b.swbs, a.swbscontrato From bitacoradeactividades b ' +
      'INNER JOIN actividadesxorden a ON (b.sContrato = a.sContrato And b.sNumeroOrden = a.sNumeroOrden And a.sIdConvenio = :Convenio And ' +
      'b.sWbs = a.sWbs) ' +
      'Where b.sContrato = :contrato and b.dIdFecha = :Fecha And b.sNumeroOrden = :Orden ');
    QryPlataforma.Params.ParamByName('contrato').DataType := ftString;
    QryPlataforma.Params.ParamByName('contrato').Value := sParamContrato;
    QryPlataforma.Params.ParamByName('convenio').DataType := ftString;
    QryPlataforma.Params.ParamByName('convenio').Value := sParamConvenio;
    QryPlataforma.Params.ParamByName('Fecha').DataType := ftDate;
    QryPlataforma.Params.ParamByName('Fecha').Value := dParamFecha;
    QryPlataforma.Params.ParamByName('Orden').DataType := ftString;
    QryPlataforma.Params.ParamByName('Orden').Value := sParamOrden;
    QryPlataforma.Open;
//    showmessage(QryPlataforma.SQL.text);
    if QryPlataforma.RecordCount > 0 then
    begin

    //cliclo para acumular los avances por orden
    while not QryPlataforma.eof do
    begin
         davanceplataforma := QryPlataforma.FieldValues['davance'];

         QryAuxiliar.Active := False;
         QryAuxiliar.SQL.Clear;
         QryAuxiliar.SQL.Add('Select sum(dAvance) as davance from ganttavances ' +
           'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
           'dIdFecha < :fecha and sNumeroOrden = :Orden and swbs = :wbs');
         QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
         QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
         QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
         QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
         QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
         QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
         QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
         QryAuxiliar.Params.ParamByName('Orden').Value := sParamOrden;
         QryAuxiliar.Params.ParamByName('wbs').DataType := ftString;
         QryAuxiliar.Params.ParamByName('wbs').Value := QryPlataforma.FieldValues['swbscontrato'];

         QryAuxiliar.Open;
         if QryAuxiliar.RecordCount > 0 then
           if (QryAuxiliar.FieldValues['dAvance'] + dAvancePlataforma) > 100 then
             dAvancePlataforma := 100 - QryAuxiliar.FieldValues['dAvance'];

         QryAuxiliar.Active := False;
         QryAuxiliar.SQL.Clear;
         QryAuxiliar.SQL.Add('Select dIdFecha, davance from ganttavances ' +
           'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
           'dIdFecha = :fecha and sNumeroOrden = :Orden and swbs = :wbs ');
         QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
         QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
         QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
         QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
         QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
         QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
         QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
         QryAuxiliar.Params.ParamByName('Orden').Value := sParamOrden;
         {QryAuxiliar.Params.ParamByName('Turno').DataType := ftString;
         QryAuxiliar.Params.ParamByName('Turno').Value := 'A';}
         QryAuxiliar.Params.ParamByName('wbs').DataType := ftString;
         QryAuxiliar.Params.ParamByName('wbs').Value := QryPlataforma.FieldValues['swbscontrato'];

         QryAuxiliar.Open;
         if QryAuxiliar.RecordCount > 0 then
         begin
           davanceplataforma := davanceplataforma + QryAuxiliar.FieldValues['davance'];

           connection.zCommand.Active := False;
           connection.zCommand.SQL.Clear;
           connection.zCommand.SQL.Add('Update ganttavances SET dAvance = :Avance ' +
             'Where sContrato = :Contrato And dIdFecha = :Fecha And ' +
             'sNumeroOrden = :Orden And sIdConvenio = :Convenio And sIdTurno = :Turno and snumeroactividad = :actividad');
           connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
           connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
           connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
           connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
           connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
           connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
           connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
           connection.zCommand.Params.ParamByName('Orden').Value := sParamOrden;
           connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
           connection.zCommand.Params.ParamByName('Turno').Value := 'A';
           connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
           connection.zCommand.Params.ParamByName('Avance').Value := dAvancePlataforma;
           connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
           connection.zCommand.Params.ParamByName('Actividad').Value := QryPlataforma.FieldValues['snumeroactividad'];
           connection.zCommand.ExecSQL
         end
         else
         begin
           connection.zCommand.Active := False;
           connection.zCommand.SQL.Clear;
           connection.zCommand.SQL.Add('INSERT INTO ganttavances (sContrato, sNumeroOrden, dIdFecha, sIdConvenio, sIdTurno, dAvance, snumeroactividad, swbs) ' +
             'Values (:Contrato, :Orden, :Fecha, :Convenio, :Turno, :Avance, :actividad, :swbs)');
           connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
           connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
           connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
           connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
           connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
           connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
           connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
           connection.zCommand.Params.ParamByName('Orden').Value := sParamOrden;
           connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
           connection.zCommand.Params.ParamByName('Turno').Value := 'A';
           connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
           connection.zCommand.Params.ParamByName('Avance').Value := dAvancePlataforma;
           connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
           connection.zCommand.Params.ParamByName('Actividad').Value := QryPlataforma.FieldValues['snumeroactividad'];
           connection.zCommand.Params.ParamByName('swbs').DataType := ftString;
           connection.zCommand.Params.ParamByName('swbs').Value := QryPlataforma.FieldValues['swbscontrato'];
           connection.zCommand.ExecSQL;
         end;

         QryPlataforma.Next;
      end;


    end;
  end
  else
  begin



        //Avance por Contrato ...
    if sParamSeguridad = 'Avanzada' then
    begin
      QryContrato.Active := False;
      QryContrato.SQL.Clear;
      QryContrato.SQL.Add('Select a.sNumeroActividad, a.dPonderado, a.dCantidadAnexo, Sum(b.dCantidad) as dCantidad From actividadescanceladas b ' +
        'INNER JOIN actividadesfcanceladas af ON (af.sContrato = b.sContrato And af.dIdFecha = b.dIdFecha And af.lAfectaContrato = "Si") ' +
        'INNER JOIN actividadesxanexo a ON (b.sContrato = a.sContrato And a.sIdConvenio = :Convenio And b.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where b.sContrato = :contrato and b.dIdFecha = :Fecha group by b.sNumeroActividad');
      QryContrato.Params.ParamByName('contrato').DataType := ftString;
      QryContrato.Params.ParamByName('contrato').Value := sParamContrato;
      QryContrato.Params.ParamByName('convenio').DataType := ftString;
      QryContrato.Params.ParamByName('convenio').Value := sParamConvenio;
      QryContrato.Params.ParamByName('Fecha').DataType := ftDate;
      QryContrato.Params.ParamByName('Fecha').Value := dParamFecha;
      QryContrato.Open;
      while not QryContrato.Eof do
      begin
        QryAuxiliar.Active := False;
        QryAuxiliar.SQL.Clear;
        QryAuxiliar.SQL.Add('Select sum(dCantidad) as dAcumulado From bitacoradeactividades ' +
          'Where sContrato = :contrato And sNumeroActividad = :Actividad And dIdFecha <= :Fecha Group By sNumeroActividad');
        QryAuxiliar.Params.ParamByName('Contrato').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Contrato').Value := sParamContrato;
        QryAuxiliar.Params.ParamByName('Actividad').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Actividad').Value := QryContrato.FieldValues['sNumeroActividad'];
        QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
        QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
        QryAuxiliar.Open;
        dAvancePartida := 0;
        if ((QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad']) >= QryContrato.FieldValues['dCantidadAnexo']) then
          dAvancePartida := 0
        else
          if (QryAuxiliar.FieldValues['dAcumulado'] > QryContrato.FieldValues['dCantidadAnexo']) then
            dAvancePartida := ((QryContrato.FieldValues['dCantidadAnexo'] - (QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad'])) * 100) / QryContrato.FieldValues['dCantidadAnexo']
          else
            dAvancePartida := (QryContrato.FieldValues['dCantidad'] * 100) / QryContrato.FieldValues['dCantidadAnexo'];
        dAvanceObra := dAvanceObra + (dAvancePartida * QryContrato.FieldValues['dPonderado']);
        QryContrato.Next;
      end
    end;
    if davanceobra > 0 then
       showmessage('tiene valor');

        // Por Contrato de Partidas reportadas en bitacora
        // Por Contrato de Partidas reportadas por Alcances
    if sParamSeguridad = 'Avanzada' then
    begin
      QryContrato.Active := False;
      QryContrato.SQL.Clear;
      QryContrato.SQL.Add('Select a.sNumeroActividad, a.dCantidadAnexo, a.dPonderado, Sum(b.dCantidad) as dCantidad, a.swbs  From bitacoradeactividades b ' +
        'INNER JOIN actividadesxanexo a ON (b.sContrato = a.sContrato And a.sIdConvenio = :Convenio And b.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where b.sContrato = :contrato and b.dIdFecha = :Fecha And b.lAlcance = "No" group by b.sNumeroActividad order by a.iItemOrden');
      QryContrato.Params.ParamByName('contrato').DataType := ftString;
      QryContrato.Params.ParamByName('contrato').Value := sParamContrato;
      QryContrato.Params.ParamByName('convenio').DataType := ftString;
      QryContrato.Params.ParamByName('convenio').Value := sParamConvenio;
      QryContrato.Params.ParamByName('Fecha').DataType := ftDate;
      QryContrato.Params.ParamByName('Fecha').Value := dParamFecha;
      QryContrato.Open;

      while not QryContrato.Eof do
      begin
        davanceobra := 0;
        QryAuxiliar.Active := False;
        QryAuxiliar.SQL.Clear;
        QryAuxiliar.SQL.Add('Select sum(dCantidad) as dAcumulado From bitacoradeactividades ' +
          'Where sContrato = :contrato And sNumeroActividad = :Actividad And dIdFecha <= :Fecha Group By sNumeroActividad');
        QryAuxiliar.Params.ParamByName('Contrato').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Contrato').Value := sParamContrato;
        QryAuxiliar.Params.ParamByName('Actividad').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Actividad').Value := QryContrato.FieldValues['sNumeroActividad'];
        QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
        QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
        QryAuxiliar.Open;
        dAvancePartida := 0;
        if ((QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad']) >= QryContrato.FieldValues['dCantidadAnexo']) then
          dAvancePartida := 0
        else
          if (QryAuxiliar.FieldValues['dAcumulado'] > QryContrato.FieldValues['dCantidadAnexo']) then
          begin
            try
              if QryContrato.FieldValues['dCantidadAnexo'] <> 0 then
                dAvancePartida := ((QryContrato.FieldValues['dCantidadAnexo'] - (QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad'])) * 100) / QryContrato.FieldValues['dCantidadAnexo']
              else
                dAvancePartida := 0;
            except
              dAvancePartida := 0;
            end
          end
          else
            dAvancePartida := (QryContrato.FieldValues['dCantidad'] * 100) / QryContrato.FieldValues['dCantidadAnexo'];
        dAvanceObra := dAvanceObra + (dAvancePartida * QryContrato.FieldValues['dPonderado']);
        davanceobra := davanceobra /100;


      // acumulaavances
      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select sum(dAvance) as dAvance from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha < :fecha and sNumeroOrden = :Orden and snumeroactividad = :actividad');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := '';
      QryAuxiliar.Params.ParamByName('actividad').DataType := ftString;
      QryAuxiliar.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];

      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
        if (QryAuxiliar.FieldValues['dAvance'] + dAvanceObra) > 100 then
          dAvanceObra := 100 - QryAuxiliar.FieldValues['dAvance'];

      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select dIdFecha from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha = :fecha and sNumeroOrden = :Orden And sIdTurno = :Turno and snumeroactividad = :actividad');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := '';
      QryAuxiliar.Params.ParamByName('Turno').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Turno').Value := 'A';
      QryAuxiliar.Params.ParamByName('actividad').DataType := ftString;
      QryAuxiliar.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];

      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update ganttavances SET dAvance = :Avance Where sContrato = :Contrato And dIdFecha = :Fecha And sNumeroOrden = :Orden And sIdConvenio = :Convenio And sIdTurno = :Turno and snumeroactividad = :actividad');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := '';
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvanceObra;
        connection.zCommand.Params.ParamByName('actividad').DataType := ftstring;
        connection.zCommand.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];

        connection.zCommand.ExecSQL
      end
      else
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('INSERT INTO ganttavances (sContrato, sNumeroOrden, dIdFecha, sIdConvenio, sIdTurno, dAvance, snumeroactividad, swbs) ' +
          'Values (:Contrato, :Orden, :Fecha, :Convenio, :Turno, :Avance, :actividad, :swbs )');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := '';
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvanceObra;
        connection.zCommand.Params.ParamByName('actividad').DataType := ftstring;
        connection.zCommand.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];
        connection.zCommand.Params.ParamByName('swbs').DataType := ftstring;
        connection.zCommand.Params.ParamByName('swbs').Value := QryContrato.FieldValues['swbs'];

        connection.zCommand.ExecSQL;
      end;


        QryContrato.Next;
      end;

      QryContrato.Active := False;
      QryContrato.SQL.Clear;
      QryContrato.SQL.Add('Select a.sNumeroActividad, a.dCantidadAnexo, a.dPonderado, b.iFase, Sum(b.dCantidad) as dCantidad, a.swbs From bitacoradealcances b ' +
        'INNER JOIN actividadesxanexo a ON (b.sContrato = a.sContrato And a.sIdConvenio = :Convenio And b.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where b.sContrato = :contrato and b.dIdFecha = :Fecha group by b.sNumeroActividad, b.iFase order by a.iItemOrden, b.iFase');
      QryContrato.Params.ParamByName('contrato').DataType := ftString;
      QryContrato.Params.ParamByName('contrato').Value := sParamContrato;
      QryContrato.Params.ParamByName('convenio').DataType := ftString;
      QryContrato.Params.ParamByName('convenio').Value := sParamConvenio;
      QryContrato.Params.ParamByName('Fecha').DataType := ftDate;
      QryContrato.Params.ParamByName('Fecha').Value := dParamFecha;
      QryContrato.Open;
      while not QryContrato.Eof do
      begin
        davanceobra := 0;
        davancepartida := 0;
        QryAuxiliar.Active := False;
        QryAuxiliar.SQL.Clear;
        QryAuxiliar.SQL.Add('Select Sum(b.dCantidad) as dCantidad, a.dAvance From bitacoradealcances b ' +
          'INNER JOIN alcancesxactividad a ON (b.sContrato = a.sContrato And b.sNumeroActividad = a.sNumeroActividad And b.iFase = a.iFase) ' +
          'Where b.sContrato = :Contrato And b.sNumeroActividad = :Actividad And b.iFase = :Fase And b.dIdFecha <= :Fecha  Group By b.sNumeroActividad');
        QryAuxiliar.Params.ParamByName('Contrato').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Contrato').Value := sParamContrato;
        QryAuxiliar.Params.ParamByName('Actividad').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Actividad').Value := QryContrato.FieldValues['sNumeroActividad'];
        QryAuxiliar.Params.ParamByName('Fase').DataType := ftInteger;
        QryAuxiliar.Params.ParamByName('Fase').Value := QryContrato.FieldValues['iFase'];
        QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
        QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
        QryAuxiliar.Open;
        dAvancePartida := 0;
        if QryAuxiliar.RecordCount > 0 then
        begin
          dAvancePartida := QryAuxiliar.FieldValues['dCantidad'];
          if (dAvancePartida - QryContrato.FieldValues['dCantidad']) >= QryContrato.FieldValues['dCantidadAnexo'] then
            dAvancePartida := 0
          else
            if dAvancePartida > QryContrato.FieldValues['dCantidadAnexo'] then
              dAvancePartida := ((QryContrato.FieldValues['dCantidadAnexo'] - (dAvancePartida - QryContrato.FieldValues['dCantidad'])) * QryAuxiliar.FieldValues['dAvance']) / QryContrato.FieldValues['dCantidadAnexo']
            else
              dAvancePartida := (QryContrato.FieldValues['dCantidad'] * QryAuxiliar.FieldValues['dAvance']) / QryContrato.FieldValues['dCantidadAnexo'];
        end;
        dAvanceObra := dAvanceObra + (dAvancePartida * QryContrato.FieldValues['dPonderado']);
        dAvanceObra := dAvanceObra / 100;


      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select sum(dAvance) as dAvance from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha < :fecha and sNumeroOrden = :Orden and snumeroactividad = :actividad');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := '';
      QryAuxiliar.Params.ParamByName('actividad').DataType := ftString;
      QryAuxiliar.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];

      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
        if (QryAuxiliar.FieldValues['dAvance'] + dAvanceObra) > 100 then
          dAvanceObra := 100 - QryAuxiliar.FieldValues['dAvance'];

      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select dIdFecha, davance from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha = :fecha and sNumeroOrden = :Orden And sIdTurno = :Turno and snumeroactividad = :actividad');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := '';
      QryAuxiliar.Params.ParamByName('Turno').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Turno').Value := 'A';
      QryAuxiliar.Params.ParamByName('actividad').DataType := ftString;
      QryAuxiliar.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];

      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
      begin
        davanceobra := davanceobra + QryAuxiliar.FieldValues['dAvance'];
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update ganttavances SET dAvance = :Avance Where sContrato = :Contrato And dIdFecha = :Fecha And sNumeroOrden = :Orden And sIdConvenio = :Convenio And sIdTurno = :Turno and snumeroactividad = :actividad');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := '';
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvanceObra;
        connection.zCommand.Params.ParamByName('actividad').DataType := ftstring;
        connection.zCommand.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];

        connection.zCommand.ExecSQL
      end
      else
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('INSERT INTO ganttavances (sContrato, sNumeroOrden, dIdFecha, sIdConvenio, sIdTurno, dAvance, snumeroactividad, swbs) ' +
          'Values (:Contrato, :Orden, :Fecha, :Convenio, :Turno, :Avance, :actividad, :swbs )');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := '';
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvanceObra;
        connection.zCommand.Params.ParamByName('actividad').DataType := ftstring;
        connection.zCommand.Params.ParamByName('actividad').Value := QryContrato.FieldValues['snumeroactividad'];
        connection.zCommand.Params.ParamByName('swbs').DataType := ftstring;
        connection.zCommand.Params.ParamByName('swbs').Value := QryContrato.FieldValues['swbs'];

        connection.zCommand.ExecSQL;
      end;

        QryContrato.Next
      end;
    end;

        // Almacenamiento de Avances ...
        // Primero se ajusta tanto el avance x programa de trabajo ...

        // Avances del Contrato ...
    if sParamSeguridad = 'Avanzada' then
    begin
    end;
  end;

  Result := dAvanceObra;
  QryAuxiliar.Destroy;
  QryContrato.Destroy;
  QryPlataforma.Destroy;
end;

procedure TfrmGraficaGerencialDX.cbPeriodoOrdenChange(Sender: TObject);
var
  vPeriodo: TPeriodo;
begin
  if tsNumeroOrden.KeyValue <> Null then
  begin
    vPeriodo := TPeriodo.Create;
    case cbPeriodoOrden.ItemIndex of
      0: begin
        // Completo
        vPeriodo := TPeriodo(TotalPeriodos.Objects[TotalPeriodos.IndexOf(OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString)]);
        tdFechaInicio.DateTime := vPeriodo.Inicio;
        tdIdFecha.DateTime := vPeriodo.Termino;
      end;
      1: begin
        // Programado
        vPeriodo := TPeriodo(ListaPeriodos.Objects[ListaPeriodos.IndexOf(OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString)]);
        tdfechainicio.DateTime := vPeriodo.Inicio;
        tdIdfecha.DateTime := vPeriodo.Termino;
      end;
      2: begin
        // Real
        vPeriodo := TPeriodo(RealPeriodos.Objects[RealPeriodos.IndexOf(OrdenesdeTrabajo.FieldByName('sNumeroOrden').AsString)]);
        tdfechainicio.DateTime := vPeriodo.Inicio;
        tdIdfecha.DateTime := vPeriodo.Termino;
      end;
    end;
  end;
end;

procedure TfrmGraficaGerencialDX.cbPeriodoChange(Sender: TObject);
begin
  mdPeriodo.RecNo := cbPeriodo.ItemIndex + 1;

  dInicio.Date := mdPeriodo.FieldByName('dInicio').AsDateTime;
  dTermino.Date := mdPeriodo.FieldByName('dTermino').AsDateTime;
end;

procedure TfrmGraficaGerencialDX.cbTipoGraficaChange(Sender: TObject);
begin
  if cbTipoGrafica.ItemIndex = 0 then
  begin
    // Respaldar el estado original del checkbox de partidas a graficar
    OldCheck := cbPartidas.Checked;
    cbPartidas.Checked := True;
    cbPartidas.Enabled := False;
  end
  else
  begin
    cbPartidas.Checked := OldCheck;
    cbPartidas.Enabled := True;
  end;
end;

function TfrmGraficaGerencialDX.cfnCalculaAvances1(sParamContrato, sParamOrden, sParamConvenio, sParamTurno: string; lParamMultiple: Boolean; dParamFecha: tDate; sParamSeguridad: string; tOrigen: TComponent): Currency;
var
  QryPlataforma: tzReadOnlyQuery;
  QryContrato, QryAuxiliar, QryReporteDiario : tzReadOnlyQuery;
  dAvancePlataforma: Currency;
  dAvanceObra: Currency;
  dAvancePartida: Currency;
  lContinua: Boolean;
begin
//  showmessage('entro al programa');
  dAvancePlataforma := 0;
  dAvanceObra := 0;
  dAvancePartida := 0;

  QryPlataforma := tzReadOnlyQuery.Create(tOrigen);
  QryPlataforma.Connection := connection.zconnection;

  QryContrato := tzReadOnlyQuery.Create(tOrigen);
  QryContrato.Connection := connection.zconnection;

  QryAuxiliar := tzReadOnlyQuery.Create(tOrigen);
  QryAuxiliar.Connection := connection.zconnection;

  QryReporteDiario := tzReadOnlyQuery.Create(tOrigen);
  QryReporteDiario.Connection := connection.zconnection;

  lcontinua := true;

  if lContinua then
  begin
        // Partidas canceladas ....
        // Avance por Plataforma ...
    QryPlataforma.Active := False;
    QryPlataforma.SQL.Clear;
    QryPlataforma.SQL.Add('Select Sum(( b.dAvance * a.dPonderado) / 100 ) as dAvance From actividadescanceladas b ' +
      'INNER JOIN actividadesxorden a ON (b.sContrato = a.sContrato And b.sNumeroOrden = a.sNumeroOrden And ' +
      'a.sIdConvenio = :Convenio And b.sWbs = a.sWbs And b.sNumeroActividad = a.sNumeroActividad) ' +
      'Where b.sContrato = :contrato and b.sNumeroOrden = :Orden and b.dIdFecha = :Fecha group by b.sNumeroOrden');
    QryPlataforma.Params.ParamByName('contrato').DataType := ftString;
    QryPlataforma.Params.ParamByName('contrato').Value := sParamContrato;
    QryPlataforma.Params.ParamByName('orden').DataType := ftString;
    QryPlataforma.Params.ParamByName('orden').Value := sParamOrden;
    QryPlataforma.Params.ParamByName('convenio').DataType := ftString;
    QryPlataforma.Params.ParamByName('convenio').Value := sParamConvenio;
    QryPlataforma.Params.ParamByName('Fecha').DataType := ftDate;
    QryPlataforma.Params.ParamByName('Fecha').Value := dParamFecha;
    QryPlataforma.Open;
    if QryPlataforma.RecordCount > 0 then
      dAvancePlataforma := dAvancePlataforma + QryPlataforma.FieldValues['dAvance'];

        //Avance por Contrato ...
    if sParamSeguridad = 'Avanzada' then
    begin
      QryContrato.Active := False;
      QryContrato.SQL.Clear;
      QryContrato.SQL.Add('Select a.sNumeroActividad, a.dPonderado, a.dCantidadAnexo, Sum(b.dCantidad) as dCantidad From actividadescanceladas b ' +
        'INNER JOIN actividadesfcanceladas af ON (af.sContrato = b.sContrato And af.dIdFecha = b.dIdFecha And af.lAfectaContrato = "Si") ' +
        'INNER JOIN actividadesxanexo a ON (b.sContrato = a.sContrato And a.sIdConvenio = :Convenio And b.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where b.sContrato = :contrato and b.dIdFecha = :Fecha group by b.sNumeroActividad');
      QryContrato.Params.ParamByName('contrato').DataType := ftString;
      QryContrato.Params.ParamByName('contrato').Value := sParamContrato;
      QryContrato.Params.ParamByName('convenio').DataType := ftString;
      QryContrato.Params.ParamByName('convenio').Value := sParamConvenio;
      QryContrato.Params.ParamByName('Fecha').DataType := ftDate;
      QryContrato.Params.ParamByName('Fecha').Value := dParamFecha;
      QryContrato.Open;
      while not QryContrato.Eof do
      begin
        QryAuxiliar.Active := False;
        QryAuxiliar.SQL.Clear;
        QryAuxiliar.SQL.Add('Select sum(dCantidad) as dAcumulado From bitacoradeactividades ' +
          'Where sContrato = :contrato And sNumeroActividad = :Actividad And dIdFecha <= :Fecha Group By sNumeroActividad');
        QryAuxiliar.Params.ParamByName('Contrato').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Contrato').Value := sParamContrato;
        QryAuxiliar.Params.ParamByName('Actividad').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Actividad').Value := QryContrato.FieldValues['sNumeroActividad'];
        QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
        QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
        QryAuxiliar.Open;
        dAvancePartida := 0;
        if ((QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad']) >= QryContrato.FieldValues['dCantidadAnexo']) then
          dAvancePartida := 0
        else
          if (QryAuxiliar.FieldValues['dAcumulado'] > QryContrato.FieldValues['dCantidadAnexo']) then
            dAvancePartida := ((QryContrato.FieldValues['dCantidadAnexo'] - (QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad'])) * 100) / QryContrato.FieldValues['dCantidadAnexo']
          else
            dAvancePartida := (QryContrato.FieldValues['dCantidad'] * 100) / QryContrato.FieldValues['dCantidadAnexo'];
        dAvanceObra := dAvanceObra + (dAvancePartida * QryContrato.FieldValues['dPonderado']);
        QryContrato.Next;
      end
    end;

        // Avance de Partidas Reportadas ....
            // por Plataforma
    QryPlataforma.Active := False;
    QryPlataforma.SQL.Clear;
    QryPlataforma.SQL.Add('Select Sum(( b.dAvance * a.dPonderado) / 100 ) as dAvance From bitacoradeactividades b ' +
      'INNER JOIN actividadesxorden a ON (b.sContrato = a.sContrato And b.sNumeroOrden = a.sNumeroOrden And a.sIdConvenio = :Convenio And ' +
      'b.sWbs = a.sWbs And b.sNumeroActividad = a.sNumeroActividad) ' +
      'Where b.sContrato = :contrato and b.dIdFecha = :Fecha And b.sNumeroOrden = :Orden group by b.sNumeroOrden');
    QryPlataforma.Params.ParamByName('contrato').DataType := ftString;
    QryPlataforma.Params.ParamByName('contrato').Value := sParamContrato;
    QryPlataforma.Params.ParamByName('convenio').DataType := ftString;
    QryPlataforma.Params.ParamByName('convenio').Value := sParamConvenio;
    QryPlataforma.Params.ParamByName('Fecha').DataType := ftDate;
    QryPlataforma.Params.ParamByName('Fecha').Value := dParamFecha;
    QryPlataforma.Params.ParamByName('Orden').DataType := ftString;
    QryPlataforma.Params.ParamByName('Orden').Value := sParamOrden;
    QryPlataforma.Open;
//    showmessage(QryPlataforma.SQL.text);
    if QryPlataforma.RecordCount > 0 then
      dAvancePlataforma := dAvancePlataforma + QryPlataforma.FieldValues['dAvance'];

        // Por Contrato de Partidas reportadas en bitacora
        // Por Contrato de Partidas reportadas por Alcances
    if sParamSeguridad = 'Avanzada' then
    begin
      QryContrato.Active := False;
      QryContrato.SQL.Clear;
      QryContrato.SQL.Add('Select a.sNumeroActividad, a.dCantidadAnexo, a.dPonderado, Sum(b.dCantidad) as dCantidad From bitacoradeactividades b ' +
        'INNER JOIN actividadesxanexo a ON (b.sContrato = a.sContrato And a.sIdConvenio = :Convenio And b.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where b.sContrato = :contrato and b.dIdFecha = :Fecha And b.lAlcance = "No" group by b.sNumeroActividad order by a.iItemOrden');
      QryContrato.Params.ParamByName('contrato').DataType := ftString;
      QryContrato.Params.ParamByName('contrato').Value := sParamContrato;
      QryContrato.Params.ParamByName('convenio').DataType := ftString;
      QryContrato.Params.ParamByName('convenio').Value := sParamConvenio;
      QryContrato.Params.ParamByName('Fecha').DataType := ftDate;
      QryContrato.Params.ParamByName('Fecha').Value := dParamFecha;
      QryContrato.Open;
      while not QryContrato.Eof do
      begin
        QryAuxiliar.Active := False;
        QryAuxiliar.SQL.Clear;
        QryAuxiliar.SQL.Add('Select sum(dCantidad) as dAcumulado From bitacoradeactividades ' +
          'Where sContrato = :contrato And sNumeroActividad = :Actividad And dIdFecha <= :Fecha Group By sNumeroActividad');
        QryAuxiliar.Params.ParamByName('Contrato').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Contrato').Value := sParamContrato;
        QryAuxiliar.Params.ParamByName('Actividad').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Actividad').Value := QryContrato.FieldValues['sNumeroActividad'];
        QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
        QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
        QryAuxiliar.Open;
        dAvancePartida := 0;
        if ((QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad']) >= QryContrato.FieldValues['dCantidadAnexo']) then
          dAvancePartida := 0
        else
          if (QryAuxiliar.FieldValues['dAcumulado'] > QryContrato.FieldValues['dCantidadAnexo']) then
          begin
            try
              if QryContrato.FieldValues['dCantidadAnexo'] <> 0 then
                dAvancePartida := ((QryContrato.FieldValues['dCantidadAnexo'] - (QryAuxiliar.FieldValues['dAcumulado'] - QryContrato.FieldValues['dCantidad'])) * 100) / QryContrato.FieldValues['dCantidadAnexo']
              else
                dAvancePartida := 0;
            except
              dAvancePartida := 0;
            end
          end
          else
            dAvancePartida := (QryContrato.FieldValues['dCantidad'] * 100) / QryContrato.FieldValues['dCantidadAnexo'];
        dAvanceObra := dAvanceObra + (dAvancePartida * QryContrato.FieldValues['dPonderado']);
        QryContrato.Next;
      end;

      QryContrato.Active := False;
      QryContrato.SQL.Clear;
      QryContrato.SQL.Add('Select a.sNumeroActividad, a.dCantidadAnexo, a.dPonderado, b.iFase, Sum(b.dCantidad) as dCantidad From bitacoradealcances b ' +
        'INNER JOIN actividadesxanexo a ON (b.sContrato = a.sContrato And a.sIdConvenio = :Convenio And b.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where b.sContrato = :contrato and b.dIdFecha = :Fecha group by b.sNumeroActividad, b.iFase order by a.iItemOrden, b.iFase');
      QryContrato.Params.ParamByName('contrato').DataType := ftString;
      QryContrato.Params.ParamByName('contrato').Value := sParamContrato;
      QryContrato.Params.ParamByName('convenio').DataType := ftString;
      QryContrato.Params.ParamByName('convenio').Value := sParamConvenio;
      QryContrato.Params.ParamByName('Fecha').DataType := ftDate;
      QryContrato.Params.ParamByName('Fecha').Value := dParamFecha;
      QryContrato.Open;
      while not QryContrato.Eof do
      begin
        QryAuxiliar.Active := False;
        QryAuxiliar.SQL.Clear;
        QryAuxiliar.SQL.Add('Select Sum(b.dCantidad) as dCantidad, a.dAvance From bitacoradealcances b ' +
          'INNER JOIN alcancesxactividad a ON (b.sContrato = a.sContrato And b.sNumeroActividad = a.sNumeroActividad And b.iFase = a.iFase) ' +
          'Where b.sContrato = :Contrato And b.sNumeroActividad = :Actividad And b.iFase = :Fase And b.dIdFecha <= :Fecha  Group By b.sNumeroActividad');
        QryAuxiliar.Params.ParamByName('Contrato').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Contrato').Value := sParamContrato;
        QryAuxiliar.Params.ParamByName('Actividad').DataType := ftString;
        QryAuxiliar.Params.ParamByName('Actividad').Value := QryContrato.FieldValues['sNumeroActividad'];
        QryAuxiliar.Params.ParamByName('Fase').DataType := ftInteger;
        QryAuxiliar.Params.ParamByName('Fase').Value := QryContrato.FieldValues['iFase'];
        QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
        QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
        QryAuxiliar.Open;
        dAvancePartida := 0;
        if QryAuxiliar.RecordCount > 0 then
        begin
          dAvancePartida := QryAuxiliar.FieldValues['dCantidad'];
          if (dAvancePartida - QryContrato.FieldValues['dCantidad']) >= QryContrato.FieldValues['dCantidadAnexo'] then
            dAvancePartida := 0
          else
            if dAvancePartida > QryContrato.FieldValues['dCantidadAnexo'] then
              dAvancePartida := ((QryContrato.FieldValues['dCantidadAnexo'] - (dAvancePartida - QryContrato.FieldValues['dCantidad'])) * QryAuxiliar.FieldValues['dAvance']) / QryContrato.FieldValues['dCantidadAnexo']
            else
              dAvancePartida := (QryContrato.FieldValues['dCantidad'] * QryAuxiliar.FieldValues['dAvance']) / QryContrato.FieldValues['dCantidadAnexo'];
        end;
        dAvanceObra := dAvanceObra + (dAvancePartida * QryContrato.FieldValues['dPonderado']);
        QryContrato.Next
      end;
      dAvanceObra := dAvanceObra / 100;
    end;

        // Almacenamiento de Avances ...
        // Primero se ajusta tanto el avance x programa de trabajo ...
    if lParamMultiple then
    begin
      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select sum(dAvance) as dAvance from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha < :fecha and sNumeroOrden = :Orden Group By sNumeroOrden');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := sParamOrden;
      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
        if (QryAuxiliar.FieldValues['dAvance'] + dAvancePlataforma) > 100 then
          dAvancePlataforma := 100 - QryAuxiliar.FieldValues['dAvance'];

      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select dIdFecha from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha = :fecha and sNumeroOrden = :Orden And sIdTurno = :Turno Group By sNumeroOrden');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := sParamOrden;
      QryAuxiliar.Params.ParamByName('Turno').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Turno').Value := 'A';
      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update ganttavances SET dAvance = :Avance ' +
          'Where sContrato = :Contrato And dIdFecha = :Fecha And ' +
          'sNumeroOrden = :Orden And sIdConvenio = :Convenio And sIdTurno = :Turno');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := sParamOrden;
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvancePlataforma;
        connection.zCommand.ExecSQL
      end
      else
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('INSERT INTO ganttavances (sContrato, sNumeroOrden, dIdFecha, sIdConvenio, sIdTurno, dAvance) ' +
          'Values (:Contrato, :Orden, :Fecha, :Convenio, :Turno, :Avance)');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := sParamOrden;
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvancePlataforma;
        connection.zCommand.ExecSQL
      end
    end;

        // Avances del Contrato ...
    if sParamSeguridad = 'Avanzada' then
    begin
      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select sum(dAvance) as dAvance from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha < :fecha and sNumeroOrden = :Orden Group By sContrato');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := '';
      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
        if (QryAuxiliar.FieldValues['dAvance'] + dAvanceObra) > 100 then
          dAvanceObra := 100 - QryAuxiliar.FieldValues['dAvance'];

      QryAuxiliar.Active := False;
      QryAuxiliar.SQL.Clear;
      QryAuxiliar.SQL.Add('Select dIdFecha from ganttavances ' +
        'where sContrato = :Contrato and sIdConvenio = :Convenio And ' +
        'dIdFecha = :fecha and sNumeroOrden = :Orden And sIdTurno = :Turno Group By sContrato');
      QryAuxiliar.Params.ParamByName('contrato').DataType := ftString;
      QryAuxiliar.Params.ParamByName('contrato').Value := sParamContrato;
      QryAuxiliar.Params.ParamByName('convenio').DataType := ftString;
      QryAuxiliar.Params.ParamByName('convenio').Value := sParamConvenio;
      QryAuxiliar.Params.ParamByName('Fecha').DataType := ftDate;
      QryAuxiliar.Params.ParamByName('Fecha').Value := dParamFecha;
      QryAuxiliar.Params.ParamByName('Orden').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Orden').Value := '';
      QryAuxiliar.Params.ParamByName('Turno').DataType := ftString;
      QryAuxiliar.Params.ParamByName('Turno').Value := 'A';
      QryAuxiliar.Open;
      if QryAuxiliar.RecordCount > 0 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update ganttavances SET dAvance = :Avance Where sContrato = :Contrato And dIdFecha = :Fecha And sNumeroOrden = :Orden And sIdConvenio = :Convenio And sIdTurno = :Turno');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := '';
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvanceObra;
        connection.zCommand.ExecSQL
      end
      else
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('INSERT INTO ganttavances (sContrato, sNumeroOrden, dIdFecha, sIdConvenio, sIdTurno, dAvance) ' +
          'Values (:Contrato, :Orden, :Fecha, :Convenio, :Turno, :Avance)');
        connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('contrato').Value := sParamContrato;
        connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
        connection.zCommand.Params.ParamByName('convenio').Value := sParamConvenio;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha;
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').Value := '';
        connection.zCommand.Params.ParamByName('Turno').DataType := ftString;
        connection.zCommand.Params.ParamByName('Turno').Value := 'A';
        connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Avance').Value := dAvanceObra;
        connection.zCommand.ExecSQL
      end
    end
  end;

  Result := dAvanceObra;
  QryAuxiliar.Destroy;
  QryContrato.Destroy;
  QryPlataforma.Destroy;
end;

procedure TfrmGraficaGerencialDX.GenerarNuevaGrafica;
type
  wbsAnterior = record
    wbs : string;
    valor : extended;
  end;

Const
  rAlto = 11;       // N�mero de renglones correspondientes al encabezado
  rAncho = 16;      // N�mero de columnas correspondientes a datos generales de partidas
  OrAnchoColumnasGrafica = 6;
  AnchoColumnasGrafica = 6;
  fAlto = 10;
  nMes: Array[1..12] of String = ('ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
                                  'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE');

var
  pInicioP, pTerminoP: Extended;   // Inicio y final del proyecto
  InicioG, TerminoG: Extended;     // Inicio y final del reporte
  aCol, nCol: Integer;
  RangoInicio,
  BarraRango,
  wCol: Extended;
  version,MaxCol,MaxRow,CeldasAUsar: dword;
  TipoReporte: Integer;
  t1,t2,t3,t4: Extended;
  rt1,rt2: Extended;
  desplaze: Extended;
  myWidth: Extended;
  InicioGrafica,InicioProyecto: TDateTime;
  Primero: Boolean;
  tColor: Integer;

  OpenDialog: TOpenDialog;
  Inicio_Semana: Integer;
  Tipo_Programa: Integer;
  fName: TFileName;
  Ext: String;
  OGantt, oHoja,
  oDetalle,
  oLibro: Variant;
  Busca: TZReadOnlyQuery;
  Ren, Ren1, RenAnt: Integer;
  Cuenta, cNivel: Integer;
  myInicio, myTermino: TDate;
  ColBarra: Integer;
  AcumAvance, DurBarra,
  AcumAnterior: Extended;
  Procede: Boolean;
  Mensaje: String;
  DiaDeLaSemana,DiasEnElMes,DiaActual,MesActual,AnioActual:Byte;
  wbsxorden: array of wbsAnterior;
  Indice,Cont: Integer;

  roqActividades: TZReadOnlyQuery;

// Obtiene el indice de periodo correspondiente a la fecha especificada
Function CalculaPeriodo(dtInicio: TDateTime;dtValor: TDateTime; Periodo: Integer): integer;
Var
  iTemp: integer;
Begin
  case Periodo of
    0:  Begin  // Diario
          CalculaPeriodo := DaysBetween(dtValor,dtInicio) + 1;
        End;
    1:  Begin    // Semanal
          iTemp := DaysBetween(dtValor,dtInicio);
          CalculaPeriodo := TRUNC(iTemp/7) + 1;
        End;
    2:  Begin     // Mensual
          if YearOf(dtInicio)=YearOf(dtValor) then
            CalculaPeriodo := MonthOf(dtValor) - MonthOf(dtInicio) + 1
          else
            begin
              iTemp := YearOf(dtValor) - YearOf(dtInicio) - 1;
              CalculaPeriodo := 12 * iTemp + (13 - MonthOf(dtInicio)) + MonthOf(dtValor);
            end;
        end;
    3:  Begin   // Anual
          CalculaPeriodo := YearOf(dtValor) - YearOf(dtInicio) + 1;
        End;
  end;
End;

// Calcula la diferencia de meses entre dos fechas (no importan los dias)
// incluyendo el actual
Function MesesEntre(dtInicio: TDateTime; dtFinal: TDateTime): integer;
Var
  iTemp: integer;
Begin
  if YearOf(dtInicio)=YearOf(dtFinal) then
    MesesEntre := MonthOf(dtFinal) - MonthOf(dtInicio)
  else
  begin
    iTemp := YearOf(dtFinal) - YearOf(dtInicio) - 1;
    MesesEntre := (12 * iTemp) + (12 - MonthOf(dtInicio)) + MonthOf(dtFinal);
  end;

End;

// Calcula la fecha de inicio del periodo seg�n la fecha especificada
Function CalcularInicioPeriodo(dtFecha: TDateTime; Tipo: Integer; Inicio_Semana: Integer): TDateTime;
begin
    if (Tipo=1) And (Inicio_Semana <> dayofweek(dtFecha)) then
  begin
    if Inicio_Semana > dayofweek(dtFecha) then
      CalcularInicioPeriodo := IncDay(dtFecha,(Inicio_Semana - dayofweek(dtFecha))-7)
    else
      CalcularInicioPeriodo := IncDay(dtFecha,Inicio_Semana - dayofweek(dtFecha));
  end
  else
    if Tipo=2 then  // Mensual
      CalcularInicioPeriodo := IncDay(dtFecha,1-DayOfTheMonth(dtFecha))
    else
      if Tipo=3 then  // Anual
        CalcularInicioPeriodo := IncDay(dtFecha,1-DayOfTheYear(dtFecha))
      else
        CalcularInicioPeriodo := dtFecha;
end;

// Ajustar el desfase entre fechas de acuerdo al periodo
Function CalcularDesfase(dtFecha: TDateTime; Tipo: Integer; Inicio_Semana: Integer): Extended;
Var
  tFecha: TDateTime;
  tDias: Integer;
begin
  // Obtener primero la fecha de inicio del periodo para la fecha
  tFecha := CalcularInicioPeriodo(dtFecha,Tipo,Inicio_Semana);
  tDias := DaysBetween(tFecha,dtFecha);
  case Tipo of
    0:  CalcularDesfase := tDias;
    1:  CalcularDesfase := tDias/7;
    2:  CalcularDesfase := tDias/DaysInAMonth(YearOf(dtFecha),MonthOf(dtFecha));
    3:  CalcularDesfase := tDias/DaysInAYear(YearOf(dtFecha));
  end;
end;

Procedure FormatoEncabezado1(Objeto: Variant; Rango: String; Valor: String;
            VAlign: integer=xlCenter;
            HAlign: integer=xlCenter);
begin
    if oGantt.Visible then
      Objeto.Select;
    SetSelect(Objeto,Rango);
    oGantt.Selection.Value := Valor;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.VerticalAlignment := VAlign;
    oGantt.Selection.HorizontalAlignment := HAlign;
    oGantt.Selection.WrapText := True;
end;

procedure GenerarBarraReal(oSheet:variant; iRow: Integer; Tipo: Integer; SchemeColor: Integer);
Var
  FechaIni,
  FechaFin,
  FechaSig    : TDateTime;
  Proporcion,
  desplaze    : Extended;
  tColBarra   : integer;
begin
  myWidth := oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Width;
  t3 := (oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Height - (fAlto/2))/2;

  roqActividades.First;
  while not roqActividades.eof do
  begin
    FechaIni := roqActividades.FieldByName('dIdFecha').AsDateTime;
    FechaSig := IncDay(roqActividades.FieldByName('dIdFecha').AsDateTime);
    roqActividades.Next;
    while (not roqActividades.eof) and (FechaSig = FechaIni) do
      roqActividades.Next;
    while (not roqActividades.eof) and (FechaSig = roqActividades.FieldByName('dIdFecha').AsDateTime) do
    begin
      FechaSig := IncDay(roqActividades.FieldByName('dIdFecha').AsDateTime);
      roqActividades.Next;
    end;
    // Regresar al registro anterior para tener el rango completo
    if roqActividades.eof then
      roqActividades.Last
    else
      roqActividades.Prior;
    FechaFin := roqActividades.FieldByName('dIdFecha').AsDateTime;

      case Tipo of
        0:  Proporcion := (FechaFin - FechaIni + 1);
        1:  Proporcion := (FechaFin - FechaIni + 1)/7;
        2:  Begin
              if monthof(FechaIni)=monthof(FechaFin) then
                Proporcion := (FechaFin - FechaIni + 1)/DaysInAMonth(YearOf(FechaIni),MonthOf(FechaIni))
              else
                Proporcion := (MesesEntre(FechaIni,FechaFin) - 1) +
                       (1 + DaysInAMonth(YearOf(FechaIni),MonthOf(FechaIni)) - DayOfTheMonth(FechaIni))/DaysInAMonth(YearOf(FechaIni),MonthOf(FechaIni)) +
                       DayOfTheMonth(FechaFin)/DaysInAMonth(YearOf(FechaFin),MonthOf(FechaFin));
            End;
        3:  Begin
              if YearOf(FechaIni)=YearOf(FechaFin) then
                Proporcion := (DayOfTheYear(FechaFin) - DayOfTheYear(FechaIni) + 1)/DaysInAYear(YearOf(FechaIni))
              else
                Proporcion := (YearOf(FechaFin) - YearOf(FechaIni) - 1) +
                       (1 + DaysInAYear(YearOf(FechaIni)) - DayOfTheYear(FechaIni))/DaysInAYear(YearOf(FechaIni))
                        + DayOfTheYear(FechaFin)/DaysInAYear(YearOf(FechaFin));
            End;
      end;

    tColBarra := CalculaPeriodo(InicioGrafica,FechaIni,TipoReporte) - 1;

    desplaze := CalcularDesfase(FechaIni,Tipo,Inicio_Semana) * myWidth;



  	t1:=oSheet.Cells[iRow - 1, rAncho + tColBarra + 1].Left + desplaze;
  	t2:=oSheet.Cells[iRow - 1, rAncho + tColBarra + 1].Top;
  	DurBarra := Proporcion * myWidth;

    oSheet.Select;
    oSheet.Cells[iRow - 1, rAncho + tColBarra + 1].Select;
    //oSheet.Cells[iRow - 1, rAncho + tColBarra + 1].Value := tColor;
  	oSheet.Shapes.AddShape(61,t1, t2+t3, DurBarra, (fAlto / 2)).Select;
    oGantt.Selection.ShapeRange.Left := oGantt.Selection.ShapeRange.Left + 2;
    oGantt.Selection.ShapeRange.Left := oGantt.Selection.ShapeRange.Left - 2;
  	oGantt.Selection.ShapeRange.Fill.Solid;
  	oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := SchemeColor;
  	oGantt.Selection.ShapeRange.Fill.OneColorGradient(msoGradientHorizontal, 4, 0.23);
  	oGantt.Selection.ShapeRange.Line.Visible := msoFalse;
    inc(tColor);

    roqActividades.Next;
  end;
end;


Procedure GenerarBarra(oSheet: Variant; iRow: Integer);
var
shape1,shape2,shape3: Variant;
hay1,hay2: Boolean;
begin
      hay1 := false;
      hay2 := false;
      oSheet.Select;

      myWidth := oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Width;
      t3 := (oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Height - (fAlto/2))/2;

      // Verificar si es una partida concentradora
      if roqAnexo.FieldByName('sTipoActividad').AsString = 'Paquete' then
      Begin
        // Color amarillo para las concentradoras
        SetSelect(oSheet,'A' + Trim(inttostr(iRow - 1)) + ':' + Letra(rAncho - 1) + Trim(Inttostr(iRow)));
        oGantt.Selection.Interior.colorIndex := 6;
        oGantt.Selection.Interior.Pattern := xlSolid;

        if ColBarra >= 0 then
        begin
          // Colocar barra concentradora en el rengl�n adecuado
          oSheet.Select;
            t1 := oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Left + desplaze;
            t2 := oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Top;
            DurBarra := rt1 * myWidth;
            shape1 := oSheet.Shapes.AddShape(61,t1, t2+t3, DurBarra, (fAlto / 2));
            shape1.Select;
            {oGantt.Selection.Cut;
            oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Select;
            oSheet.Paste;
            oGantt.Selection.ShapeRange.Left := oGantt.Selection.ShapeRange.Left + desplaze;
            }
            oGantt.Selection.ShapeRange.Fill.Solid;
            oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 0;
            oGantt.Selection.ShapeRange.Line.Visible := msoFalse;

          if (myInicio = roqAnexo.FieldByName('dFechaInicio').AsFloat) and (DurBarra >= fAlto) then
          Begin
            // Colocar indicador de inicio de concentradora en el rengl�n adecuado
            //oSheet.Select;
              //oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Select;
              t1:=oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Left + desplaze + (fAlto / 4);
              t2:=oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Top;

              shape2 := oSheet.Shapes.AddShape(msoShapeRightTriangle, t1, t2+t3, (fAlto / 2), fAlto);
              shape2.Select;
              oGantt.Selection.ShapeRange.IncrementRotation(90);
              oGantt.Selection.ShapeRange.Fill.Solid;
              oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 0;
              oGantt.Selection.ShapeRange.Line.Visible := msoFalse;
              hay1 := true;
          End;

          if (myTermino =roqAnexo.FieldByName('dFechaFinal').AsFloat) and (DurBarra >= fAlto) then
          Begin
            // Colocar indicador de termino de concentradora
            //oSheet.Select;
              //oSheet.Cells[iRow - 1, rAncho + ColBarra + 2].Select;
              t1:=DurBarra + oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Left + desplaze + (fAlto/4);
              t2:=oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Top;

              shape3 := oSheet.Shapes.AddShape(msoShapeRightTriangle, t1, t2+t3, fAlto/2, fAlto);
              shape3.Select;
              oGantt.Selection.ShapeRange.IncrementRotation(90);
              oGantt.Selection.ShapeRange.Flip(0);
              oGantt.Selection.ShapeRange.IncrementLeft(fAlto * -1);
              oGantt.Selection.ShapeRange.Fill.Solid;
              oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 0;
              oGantt.Selection.ShapeRange.Line.Visible := msoFalse;
              hay2 := true;
          End;

          shape1.select(true);
          if hay1 then shape2.select(false);
          if hay2 then shape3.select(false);
          if hay1 OR hay2 then oGantt.Selection.shaperange.Group;
        end;
      end
      Else
      Begin
        // Colocar barra de actividad programada
        if ColBarra >= 0 then
        begin
          //oSheet.Select;
            //oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Select;
            t1:=oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Left + desplaze;
            t2:=oSheet.Cells[iRow - 1, rAncho + ColBarra + 1].Top;
            DurBarra := rt1 * myWidth;

            oSheet.Shapes.AddShape(61,t1, t2+t3, DurBarra, (fAlto / 2)).Select(true);
            oGantt.Selection.ShapeRange.Fill.Solid;
            oGantt.Selection.ShapeRange.Fill.ForeColor.SchemeColor := 48;
            oGantt.Selection.ShapeRange.Fill.OneColorGradient(msoGradientHorizontal, 4, 0.23);
            oGantt.Selection.ShapeRange.Line.Visible := msoFalse;
        end;
      End;

end;

// Enmarcar el rango con lineas gruesas externas y lineas delgadas internas
Procedure EnmarcarArea(oSheet: Variant; sRango: String; Externo: Integer; Interno: Integer; Dentro:Boolean);
Var
  oRango: Variant;
begin
    oSheet.Select;
    oRango := oSheet.Range[sRango];
    oRango.Select;
    oRango.Borders[xlDiagonalDown].LineStyle := xlNone;
    oRango.Borders[xlDiagonalUp].LineStyle := xlNone;
    oRango.Borders[xlEdgeLeft].LineStyle := xlContinuous;
    oRango.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
    oRango.Borders[xlEdgeLeft].Weight := Externo;
    oRango.Borders[xlEdgeTop].LineStyle := xlContinuous;
    oRango.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
    oRango.Borders[xlEdgeTop].Weight := Externo;
    oRango.Borders[xlEdgeBottom].LineStyle := xlContinuous;
    oRango.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
    oRango.Borders[xlEdgeBottom].Weight := Externo;
    oRango.Borders[xlEdgeRight].LineStyle := xlContinuous;
    oRango.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
    oRango.Borders[xlEdgeRight].Weight := Externo;
    oRango.Borders[xlInsideVertical].LineStyle := xlContinuous;
    oRango.Borders[xlInsideVertical].ColorIndex := xlAutomatic;
    oRango.Borders[xlInsideVertical].Weight := Interno;
    if Dentro then
    begin
      oRango.Borders[xlInsideHorizontal].LineStyle := xlContinuous;
      oRango.Borders[xlInsideHorizontal].ColorIndex := xlAutomatic;
      oRango.Borders[xlInsideHorizontal].Weight := Interno;
    end;
end;

Function xEncabezado: Boolean;
Var
  myCol: Variant;
  myRow: Variant;
  myCell: Variant;
  myTiempo: Real;
  Bloque: Extended;
  mFecha: TDateTime;
  OldFecha: Integer;
  OldCol, xaCol: Integer;
  sValor: String;
  Cadena, tmpNombre: String;
  fs: TStream;

function GetTempFile(const Extension: string): string;
var
  Buffer: array[0..MAX_PATH] of Char;
begin
  Repeat
    GetTempPath(SizeOf(Buffer) - 1, Buffer);
    GetTempFileName(Buffer, '~', 0, Buffer);
    Result := ChangeFileExt(Buffer, Extension);
  until not FileExists(Result);
end;

Begin
  Result := False;

  oHoja.Select;
  oGantt.ActiveWindow.Zoom := 75;
  SetColumnWidth(oGantt,3,3,50);  //Columns['B:B'].ColumnWidth := 50;
  SetColumnWidth(oGantt,rAncho,rAncho,0.5);

  oDetalle.Select;
  oGantt.ActiveWindow.Zoom := 75;
  SetColumnWidth(oGantt,3,3,50);  //Columns['B:B'].ColumnWidth := 50;
  SetColumnWidth(oGantt,rAncho,rAncho,0.5);

  // Dar formato general a las columnas
    oHoja.Select;
    myCol := oHoja.Columns['A:A'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '@';
    oGantt.Selection.Font.Size := 8;
    oGantt.Selection.ColumnWidth := 10.71;

    oDetalle.Select;
    myCol := oDetalle.Columns['A:A'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '@';
    oGantt.Selection.Font.Size := 8;
    oGantt.Selection.ColumnWidth := 10.71;

    oHoja.Select;
    myCol := oHoja.Columns['B:B'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlGeneral;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.Font.Size := 8;

    oDetalle.Select;
    myCol := oDetalle.Columns['B:B'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlGeneral;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.Font.Size := 8;

    oHoja.Select;
    myCol := oHoja.Columns['C:C'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlGeneral;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.Font.Size := 8;

    oDetalle.Select;
    myCol := oDetalle.Columns['C:C'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlGeneral;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.Font.Size := 8;

    oHoja.Select;
    myCol := oHoja.Columns['D:J'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.VerticalAlignment := xlTop;

    oDetalle.Select;
    myCol := oDetalle.Columns['D:J'];
    myCol.Select;
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.VerticalAlignment := xlTop;

    oHoja.Select;
    myCol := oHoja.Columns['F:F'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oDetalle.Select;
    myCol := oDetalle.Columns['F:F'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oHoja.Select;
    myCol := oHoja.Columns['G:G'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000';
    oGantt.Selection.ColumnWidth := 0;

    oDetalle.Select;
    myCol := oDetalle.Columns['G:G'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000';
    oGantt.Selection.ColumnWidth := 0;

    oHoja.Select;
    myCol := oHoja.Columns['H:H'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000';
    oGantt.Selection.ColumnWidth := 0;

    oDetalle.Select;
    myCol := oDetalle.Columns['H:H'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000';
    oGantt.Selection.ColumnWidth := 0;

    oHoja.Select;
    myCol := oHoja.Columns['I:I'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oDetalle.Select;
    myCol := oDetalle.Columns['I:I'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oHoja.Select;
    myCol := oHoja.Columns['J:J'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.00';
    oGantt.Selection.ColumnWidth := 0;

    oDetalle.Select;
    myCol := oDetalle.Columns['J:J'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.00';
    oGantt.Selection.ColumnWidth := 0;

    oHoja.Select;
    myCol := oHoja.Columns['K:K'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.00';
    oGantt.Selection.ColumnWidth := 0;

    oDetalle.Select;
    myCol := oDetalle.Columns['K:K'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.00';
    oGantt.Selection.ColumnWidth := 0;

    oHoja.Select;
    myCol := oHoja.Columns['L:L'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oDetalle.Select;
    myCol := oDetalle.Columns['L:L'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oHoja.Select;
    myCol := oHoja.Columns['M:M'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oDetalle.Select;
    myCol := oDetalle.Columns['M:M'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oHoja.Select;
    myCol := oHoja.Columns['N:N'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oDetalle.Select;
    myCol := oDetalle.Columns['N:N'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := '0.0000%';

    oHoja.Select;
    myCol := oHoja.Columns['O:O'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.WrapText := True;
    oGantt.Selection.ColumnWidth := 12;

    oDetalle.Select;
    myCol := oDetalle.Columns['O:O'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.WrapText := True;
    oGantt.Selection.ColumnWidth := 12;

    oHoja.Select;
    myCol := oHoja.Columns['P:P'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.WrapText := True;
    oGantt.Selection.ColumnWidth := 12;

    oDetalle.Select;
    myCol := oDetalle.Columns['P:P'];
    myCol.Select;
    oGantt.Selection.VerticalAlignment := xlTop;
    oGantt.Selection.NumberFormat := 'dd/mm/aaaa h:mm';
    oGantt.Selection.HorizontalAlignment := xlCenter;
    oGantt.Selection.WrapText := True;
    oGantt.Selection.ColumnWidth := 12;

    oHoja.Select;
    SetRowHeight(oGantt,10,13.8);
    myRow := oGantt.Rows['10:10'];
    myRow.WrapText := True;

    oDetalle.Select;
    SetRowHeight(oGantt,10,13.8);
    myRow := oGantt.Rows['10:10'];
    myRow.WrapText := True;

    oHoja.Select;
    myRow := oGantt.Rows['10:10'];
    myRow.VerticalAlignment := xlTop;

    oDetalle.Select;
    myRow := oGantt.Rows['10:10'];
    myRow.VerticalAlignment := xlTop;

    FormatoEncabezado1(oHoja,'A09:A11','PARTIDA');
    FormatoEncabezado1(oDetalle,'A09:A11','PARTIDA');

    FormatoEncabezado1(oHoja,'B09:B11','PARTIDA');
    FormatoEncabezado1(oDetalle,'B09:B11','PARTIDA');

    FormatoEncabezado1(oHoja,'C09:C11','DESCRIPCI�N');
    FormatoEncabezado1(oDetalle,'C09:C11','DESCRIPCI�N');

    FormatoEncabezado1(oHoja,'D09:D11','TIEMPO PROGRAMADO');
    FormatoEncabezado1(oDetalle,'D09:D11','TIEMPO PROGRAMADO');

    FormatoEncabezado1(oHoja,'E09:E11','DIAS');
    FormatoEncabezado1(oDetalle,'E09:E11','DIAS');

    FormatoEncabezado1(oHoja,'F09:F11','POND.');
    FormatoEncabezado1(oDetalle,'F09:F11','POND.');

    FormatoEncabezado1(oHoja,'G09:G11','COSTO TOTAL MN');
    FormatoEncabezado1(oDetalle,'G09:G11','COSTO TOTAL MN');

    FormatoEncabezado1(oHoja,'H09:H11','COSTO TOTAL USD');
    FormatoEncabezado1(oDetalle,'H09:H11','COSTO TOTAL USD');

    FormatoEncabezado1(oHoja,'I09:I11','AVANCE PONDERADO');
    FormatoEncabezado1(oDetalle,'I09:I11','AVANCE PONDERADO');

    FormatoEncabezado1(oHoja,'J09:J11','COSTO EJECUTADO MN');
    FormatoEncabezado1(oDetalle,'J09:J11','COSTO EJECUTADO MN');

    FormatoEncabezado1(oHoja,'K09:K11','COSTO EJECUTADO USD');
    FormatoEncabezado1(oDetalle,'K09:K11','COSTO EJECUTADO USD');

    FormatoEncabezado1(oHoja,'L09:L11','AVANCE ANTERIOR AL PERIODO');
    FormatoEncabezado1(oDetalle,'L09:L11','AVANCE ANTERIOR AL PERIODO');

    FormatoEncabezado1(oHoja,'M09:M11','AVANCE PERIODO');
    FormatoEncabezado1(oDetalle,'M09:M11','AVANCE PERIODO');

    FormatoEncabezado1(oHoja,'N09:N11','AVANCE TOTAL AL PERIODO');
    FormatoEncabezado1(oDetalle,'N09:N11','AVANCE TOTAL AL PERIODO');

    FormatoEncabezado1(oHoja,'O09:O11','FECHA PROG. DE INICIO');
    FormatoEncabezado1(oDetalle,'O09:O11','FECHA PROG. DE INICIO');

    FormatoEncabezado1(oHoja,'P09:P11','FECHA PROG. DE TERMINO');
    FormatoEncabezado1(oDetalle,'P09:P11','FECHA PROG. DE TERMINO');

    oHoja.Select;

  // Determinar la altura total de los titulos
    myRow := oHoja.Rows['1:6'];
    myRow.RowHeight := 30;
    myRow := oHoja.Rows['7:7'];
    myRow.RowHeight := 40;
    myRow := oHoja.Rows['8:8'];
    myRow.RowHeight := 32;

    myRow := oDetalle.Rows['1:6'];
    myRow.RowHeight := 30;
    myRow := oDetalle.Rows['7:7'];
    myRow.RowHeight := 40;
    myRow := oDetalle.Rows['8:8'];
    myRow.RowHeight := 32;

  pInicioP := mdPeriodo.FieldByName('dInicio').AsDateTime;
  pTerminoP := mdPeriodo.FieldByName('dTermino').AsDateTime;

  {
  // Ajustar periodo de inicio al bloque completo
  myTiempo := 1440 / 24 / 60;     // Convertir D�as a formato de tiempo de excel
  Bloque := Trunc(pInicioP / myTiempo);     // Obtener bloque de tiempo de acuerdo a periodo seleccionado
  pInicioP := Bloque * myTiempo;            // Cerrar periodo de inicio de acuerdo a bloque de tiempo seleccionado

  nCol := Trunc(pTerminoP - pInicioP);
  }
  // Establecer las fechas en titulos de barras

  // Checar si el reporte es semanal, para cambiar la fecha de inicio de la gr�fica
  // al primer d�a habil parametrizado
  mFecha := CalcularInicioPeriodo(InicioG,TipoReporte,Inicio_Semana);

  OldFecha := 0;
  OldCol := 0;

  InicioGrafica := mFecha;
  InicioProyecto := CalcularInicioPeriodo(pInicioP,TipoReporte,Inicio_Semana);

  for xaCol := 0 to CeldasAUsar - 1 do
  begin


    // Con este llamado puedo saber cual es el n�mero del periodo de acuerdo a como
    // esta la gr�fica
    //myCell.Value := CalculaPeriodo(InicioGrafica,mFecha,TipoReporte);

    // Con este llamado puedo saber cual es el n�mero de acuerdo a como
    // esta en el proyecto (No necesariamente igual al de arriba)
    //myCell.Value := CalculaPeriodo(InicioProyecto,mFecha,TipoReporte);

    OldCol := CalculaPeriodo(InicioProyecto,mFecha,TipoReporte);

    sValor := Letra(rAncho + xaCol + 1) + '11';
    oHoja.Select;
    myCell := oHoja.Range[sValor];
    myCell.HorizontalAlignment := xlCenter;
    myCell.Value := OldCol;

    oDetalle.Select;
    myCell := oDetalle.Range[sValor];
    myCell.HorizontalAlignment := xlCenter;
    myCell.Value := OldCol;




    sValor := Letra(rAncho + xaCol + 1) + '10';

    case TipoReporte of
      0,1:Begin  // diario y semanal
            oHoja.Select;
            myCell := oHoja.Range[sValor];
            myCell.NumberFormat := 'dd/mm/aaaa';
            myCell.Font.Size := 10;
            myCell.HorizontalAlignment := xlCenter;
            myCell.Value := mFecha;
            if TipoReporte=0 then
            begin
              SetSelect(oHoja,Letra(rAncho + xaCol + 1) + '8:' + Letra(rAncho + xaCol + 1) + '10');
              oGantt.Selection.MergeCells := True;
              oGantt.Selection.HorizontalAlignment := xlCenter;
              oGantt.Selection.VerticalAlignment := -4107;
              oGantt.Selection.Orientation := -90;
              oGantt.Selection.ReadingOrder := -5002;
            end;


            oDetalle.Select;
            myCell := oDetalle.Range[sValor];
            myCell.Font.Size := 10;
            myCell.NumberFormat := 'dd/mm/aaaa';
            myCell.HorizontalAlignment := xlCenter;
            myCell.Value := mFecha;
            if TipoReporte=0 then
            begin
              SetSelect(oDetalle,Letra(rAncho + xaCol + 1) + '8:' + Letra(rAncho + xaCol + 1) + '10');
              oGantt.Selection.MergeCells := True;
              oGantt.Selection.HorizontalAlignment := xlCenter;
              oGantt.Selection.VerticalAlignment := -4107;
              oGantt.Selection.Orientation := -90;
              oGantt.Selection.ReadingOrder := -5002;
            end;

            if TipoReporte=0 then
              mFecha := IncDay(mFecha,1)
            else
              mFecha := IncDay(mFecha,7)
          End;
      2:  Begin  // mensual
            oHoja.Select;
            myCell := oHoja.Range[sValor];
            myCell.NumberFormat := 'mm/aaaa';
            myCell.HorizontalAlignment := xlCenter;
            myCell.Value := mFecha;

            oDetalle.Select;
            myCell := oDetalle.Range[sValor];
            myCell.NumberFormat := 'mmm/aaaa';
            myCell.HorizontalAlignment := xlCenter;
            myCell.Value := mFecha;

            mFecha := IncMonth(mFecha,1)
          End;
      3:  Begin  // anual
            oHoja.Select;
            myCell := oHoja.Range[sValor];
            myCell.NumberFormat := 'aaaa';
            myCell.HorizontalAlignment := xlCenter;
            myCell.Value := mFecha;

            oDetalle.Select;
            myCell := oDetalle.Range[sValor];
            myCell.NumberFormat := 'aaaa';
            myCell.HorizontalAlignment := xlCenter;
            myCell.Value := mFecha;

            mFecha := IncYear(mFecha,1)
          End;
    end;
  end;

  oHoja.Select;
    // Establecer el ancho de columnas  (Multiplicar por un factor exponencial seg�nse incremente el tipo de reporte)
    SetColumnWidth(oGantt,rAncho + 1,rAncho + xaCol, AnchoColumnasGrafica * (1 + TipoReporte*TipoReporte*2));
    myCol := oGantt.Columns[Letra(rAncho + 1)]; // + Letra(xaCol + rAncho)];
    myCol.Select;

  oDetalle.Select;
    // Establecer el ancho de columnas
    SetColumnWidth(oGantt,rAncho + 1,rAncho + xaCol, AnchoColumnasGrafica * (1 + TipoReporte*TipoReporte*2));
    myCol := oGantt.Columns[Letra(rAncho + 1)]; // + Letra(xaCol + rAncho)];
    myCol.Select;

    wCol := myCol.Width * xaCol;
    Dec(xaCol);
    RangoInicio := pInicioP;
    BarraRango := myTiempo;

    oHoja.Select;
    if xaCol > 8 then
      SetSelect(oHoja,'C1:' + Letra(rAncho + xaCol - 9) + '1')
    else
      SetSelect(oHoja,'C1:I1');


  oDetalle.Select;
    if xaCol > 8 then
      SetSelect(oDetalle,'C1:' + Letra(rAncho + xaCol - 9) + '1')
    else
      SetSelect(oDetalle,'C1:I1');


  oHoja.Select;
    oGantt.Selection.HorizontalAlignment := xlLeft;
    oGantt.Selection.Font.Size := 24;
    oGantt.Selection.Font.Bold := True;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.WrapText := False;
    oGantt.Selection.Value := Connection.Configuracion.FieldByName('sNombreCorto').AsString;

  oDetalle.Select;
    oGantt.Selection.HorizontalAlignment := xlLeft;
    oGantt.Selection.Font.Size := 24;
    oGantt.Selection.Font.Bold := True;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.WrapText := False;
    oGantt.Selection.Value := Connection.Configuracion.FieldByName('sNombreCorto').AsString;

  // Generar los encabezados de la gr�fica
  oHoja.Select;
    if xaCol > 8 then
      SetSelect(oHoja,'C2:' + Letra(rAncho + xaCol - 9) + '6')
    else
      SetSelect(oHoja,'C2:I6');

  oDetalle.Select;
    if xaCol > 8 then
      SetSelect(oDetalle,'C2:' + Letra(rAncho + xaCol - 9) + '6')
    else
      SetSelect(oDetalle,'C2:I6');

  oHoja.Select;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.HorizontalAlignment := xlLeft;
    oGantt.Selection.Font.Size := 14;
    oGantt.Selection.Font.Bold := True;
    oGantt.Selection.WrapText := True;

  oDetalle.Select;
    oGantt.Selection.MergeCells := True;
    oGantt.Selection.HorizontalAlignment := xlLeft;
    oGantt.Selection.Font.Size := 14;
    oGantt.Selection.Font.Bold := True;
    oGantt.Selection.WrapText := True;

  Cadena := Connection.contrato.FieldByName('mCliente').AsString + chr(10) + chr(10);
  Cadena := Cadena + Connection.contrato.FieldByName('sContrato').AsString + '  -  ' + Connection.contrato.FieldByName('mDescripcion').AsString + chr(10) + chr(10);
  Cadena := QuitarEnter(Cadena);

  SetSelect(oHoja,'C2');
  oGantt.Selection.Value := Cadena;

  SetSelect(oDetalle,'C2');
  oGantt.Selection.Value := Cadena;

  Cadena := 'DIAGRAMA DE GANTT GERENCIAL PERIODO DE GRAFICACI�N DEL ' + IntToStr(DayOfTheMonth(dInicio.Date)) + ' DE ' + nMes[MonthOfTheYear(dInicio.Date)] + ' DE ' + FormatDateTime('yyyy',dInicio.Date) + ' AL ' + IntToStr(DayOfTheMonth(dTermino.Date)) + ' DE ' + nMes[MonthOfTheYear(dTermino.Date)] + ' DE ' + FormatDateTime('yyyy',dTermino.Date);
  Cadena := QuitarEnter(Cadena);

  SetSelect(oHoja,'C7:' + Letra(13 + xaCol) + '7');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 14;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := False;
  oGantt.Selection.Value := Cadena;

  SetSelect(oDetalle,'C7:' + Letra(13 + xaCol) + '7');
  oGantt.Selection.MergeCells := True;
  oGantt.Selection.HorizontalAlignment := xlLeft;
  oGantt.Selection.VerticalAlignment := xlCenter;
  oGantt.Selection.Font.Size := 14;
  oGantt.Selection.Font.Bold := True;
  oGantt.Selection.WrapText := False;
  oGantt.Selection.Value := Cadena;

  // Formatear los bordes del encabezado
  EnmarcarArea(oHoja,'A8:' + Letra(CeldasAUsar + RAncho) + '11',xlMedium,xlThin, true);
  EnmarcarArea(oDetalle,'A8:' + Letra(CeldasAUsar + RAncho) + '11',xlMedium,xlThin, true);


  // Obtener la imagen del cliente desde la base de datos
  tmpNombre := GetTempFile('.~im');

  fs := Connection.configuracion.CreateBlobStream(Connection.configuracion.FieldByName('bImagen'), bmRead);
  try
    fs.Seek(0, soFromBeginning);
    with TFileStream.Create(tmpNombre, fmCreate) do
      try
        CopyFrom(fs, fs.Size)
      finally
        Free
      end;
  finally
    fs.Free
  end;

  // Agregar Imagen de la empresa
  oHoja.Select;
    oHoja.Cells[1,1].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombre).Select;

    // Determinar el tama�o real de la imagen
    Bloque := (oGantt.Rows[1].Height + oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height) + 20;   // * 0.7;
    oGantt.Selection.ShapeRange.ScaleWidth(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromBottomRight);
    oGantt.Selection.ShapeRange.ScaleHeight(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);

    oGantt.Selection.ShapeRange.Left := 24;  //Margen;    //Excel.Columns['A:A'].Width + Margen;
    oGantt.Selection.ShapeRange.Top := 0;

  // Agregar Imagen de la empresa
  oDetalle.Select;
    oDetalle.Cells[1,1].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombre).Select;

    // Determinar el tama�o real de la imagen
    Bloque := (oGantt.Rows[1].Height + oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height) + 20;   // * 0.7;
    oGantt.Selection.ShapeRange.ScaleWidth(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromBottomRight);
    oGantt.Selection.ShapeRange.ScaleHeight(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);

    oGantt.Selection.ShapeRange.Left := 24;  //Margen;    //Excel.Columns['A:A'].Width + Margen;
    oGantt.Selection.ShapeRange.Top := 0;

  // Obtener la imagen de la empresa desde la base de datos
  fs := Connection.Contrato.CreateBlobStream(Connection.Contrato.FieldByName('bImagen'), bmRead);
  try
    fs.Seek(0, soFromBeginning);
    with TFileStream.Create(tmpNombre, fmCreate) do
      try
        CopyFrom(fs, fs.Size)
      finally
        Free
    end;
  finally
    fs.Free
  end;

  // Agregar Imagen Cliente a la hoja de excel
  oHoja.Select;
    oHoja.Cells[2,18 + xaCol].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombre).Select;
    // Determinar el tama�o real de la imagen
    Bloque := (oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height);   // * 0.7;
    oGantt.Selection.ShapeRange.ScaleWidth(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft); //msoScaleFromBottomRight);
    oGantt.Selection.ShapeRange.ScaleHeight(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);

    oGantt.Selection.ShapeRange.IncrementLeft(oGantt.Selection.ShapeRange.Width * -1);
    oGantt.Selection.ShapeRange.Top := oGantt.Rows[1].Height;

  oDetalle.Select;
    oDetalle.Cells[2,18 + xaCol].Select;
    oGantt.ActiveSheet.Pictures.Insert(tmpNombre).Select;
    // Determinar el tama�o real de la imagen
    Bloque := (oGantt.Rows[2].Height + oGantt.Rows[3].Height + oGantt.Rows[4].Height + oGantt.Rows[5].Height);   // * 0.7;
    oGantt.Selection.ShapeRange.ScaleWidth(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft); //msoScaleFromBottomRight);
    oGantt.Selection.ShapeRange.ScaleHeight(Bloque / oGantt.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);

    oGantt.Selection.ShapeRange.IncrementLeft(oGantt.Selection.ShapeRange.Width * -1);
    oGantt.Selection.ShapeRange.Top := oGantt.Rows[1].Height;

  // Eliminar el archivo temporal
  Try
    DeleteFile(tmpNombre);
  Finally
    tmpNombre := '';
  End;

  Result := True;
end;

Function Ceiling(Valor: Extended): DWord;
Begin
  if Frac(Valor) > 0 then
    Ceiling := TRUNC(Valor) + 1
  else
    Ceiling := TRUNC(Valor);
End;

begin
  roqActividades := TZReadOnlyQuery.Create(Nil);
  roqActividades.Connection := Connection.ConnTrx;
  roqActividades.SQL.Text := 'select distinct b.swbscontrato, a.snumeroorden, a.swbs, a.snumeroactividad, a.dIdFecha ' +
                             'from bitacoradeactividades a inner join actividadesxorden b on a.scontrato=b.scontrato and a.swbs=b.swbs ' +
                             'where a.scontrato = :contrato and dIdFecha between :inicio and :final and a.sIdTipoMovimiento = "A" ' +
                             'order by dIdFecha';

  Inicio_Semana := 1;  // Para esta variable, el 1 representa el Domingo
  Tipo_Programa := 2;
  tColor := 1;

  InicioG := dInicio.Date;
  TerminoG := dTermino.Date;

  // Abrir los datos
  roqAnexo.Close;
  roqAnexo.ParamByName('contrato').AsString := global_contrato;
  roqAnexo.ParamByName('convenio').AsString := global_convenio;
  roqAnexo.ParamByName('inicio').AsDate := dInicio.Date;
  roqAnexo.ParamByName('final').AsDate := dTermino.Date;
  roqAnexo.Open;

  roqOrdenAntes.Close;
  roqOrdenAntes.ParamByName('contrato').AsString := global_contrato;
  roqOrdenAntes.ParamByName('convenio').AsString := global_convenio;
  roqOrdenAntes.ParamByName('inicio').AsFloat := 0;
  roqOrdenAntes.ParamByName('final').AsDate := IncDay(dInicio.Date, -1);
  roqOrdenAntes.Open;

  roqOrdenHoy.Close;
  roqOrdenHoy.ParamByName('contrato').AsString := global_contrato;
  roqOrdenHoy.ParamByName('convenio').AsString := global_convenio;
  roqOrdenHoy.ParamByName('inicio').AsDate := dInicio.Date;
  roqOrdenHoy.ParamByName('final').AsDate := dTermino.Date;
  roqOrdenHoy.Open;
  roqOrdenHoy.first;

  roqActividades.Close;
  roqActividades.ParamByName('contrato').AsString := global_contrato;
  roqActividades.ParamByName('inicio').AsDate := dInicio.Date;
  roqActividades.ParamByName('final').AsDate := dTermino.Date;
  roqActividades.Open;

  //  Establecer por medio del uso de un usuario el nombre y localizaci�n del archivo que se ha de generar
  OpenDialog := TOpenDialog.Create(self);
  OpenDialog.Filter := 'Archivos de MS EXCEL|*.xls';
  OpenDialog.FilterIndex := 1;

  // Determinar el nombre propuesto para la grafica de gantt
  OpenDialog.FileName := 'Gr�fica de GANTT de seguimiento - ' + global_contrato + '.xls';

  // Verificar la extenci�n en el nombre de archivo
  If openDialog.Execute Then
    fName := OpenDialog.FileName
  Else
  Begin
    Close;
    Exit;
  End;

  if Length(fName) > 3 then
  Begin
    Ext := Upcase(fName[Length(fName)- 3]) + Upcase(fName[Length(fName)- 2]) + Upcase(fName[Length(fName)- 1]) + Upcase(fName[Length(fName)]);
    if Ext <> '.XLS' then
      fName := fName + '.xls';
  End
  Else
    fName := fName + '.xls';

  // Generar el ambiente de excel para vaciar los datos de la gr�fica
  Try
    oGantt := CreateOleObject('Excel.Application');
  except
    oGantt.quit;
    showmessage('No se pudo crear el objeto EXCEL o se produjo alg�n error.');
    Exit;
  End;

  version := oGantt.version;

  oGantt.Visible := False;
  oGantt.DisplayAlerts := False;
 // oGantt.ScreenUpdating := False;
  oLibro := oGantt.Workbooks.Add;    // Crear el libro sobre el que se ha de trabajar

  // Analizar el n�mero de hojas que se requieren para la generaci�n
  while OGantt.Sheets.Count <> 2 do //Objeto.RecordCount + 2 do
  begin
    if OGantt.Sheets.Count > 1 then //Objeto.RecordCount + 1 then
    begin
      OGantt.Sheets[1].Select;
      OGantt.ActiveSheet.Delete;
    end
    else
      OGantt.Sheets.Add;
  end;

  oHoja := oLibro.Sheets[1];        // Determinar la hoja de trabajo
  oHoja.Name := 'ANEXO CONDENSADO';
  maxcol := oHoja.Columns.count;
  //maxcol := 256;
  maxrow := oHoja.Rows.count;
  TipoReporte := cbTipo.ItemIndex;

  // Determinar de acuerdo a las fechas cuantas columnas van a ser necesarias
  repeat
    case TipoReporte of
      0: begin   // Dias
            CeldasAUsar := TRUNC(dTermino.Date - dInicio.Date) + 1;
            Mensaje := 'Diario'
         end;
      1: begin  // Semanas
            CeldasAUsar := Ceiling((TRUNC(dTermino.Date - dInicio.Date) + 1)/7) + 1;
            Mensaje := 'Semanal'
         end;
      2: begin  // Meses
            if yearof(dInicio.Date)=yearof(dTermino.Date) then
                CeldasAUsar := monthof(dTermino.Date) - monthof(dInicio.Date) + 1
            else
                CeldasAUsar := 12*(yearof(dTermino.Date) - yearof(dInicio.Date) - 1) + (13 - monthof(dInicio.Date)) + monthof(dTermino.Date);
            Mensaje := 'Mensual'
        end;
      3: begin  // A�os
            CeldasAUsar := yearof(dTermino.Date) - yearof(dInicio.Date) + 1;
            Mensaje := 'Anual'
         end;
    end;

    if CeldasAUsar + RAncho > maxcol then
      TipoReporte := TipoReporte + 1


  until (CeldasAUsar + RAncho < maxcol) OR (TipoReporte>3) ;

  if TipoReporte=4 then   // Si no se puede sacar ning�n Tipo de Reporte posible, salir
  begin
    // Cerrar excel
    oLibro.close(False);
    oGantt.Quit;
    ShowMessage('El reporte no puede ser generado.'+chr(10)+'Se ha sobrepasado el l�mite de columnas');
    exit;
  end;


  if TipoReporte<>cbTipo.ItemIndex then
      if messagedlg('La gr�fica ocupa m�s columnas que las que soporta Excel.'+chr(10)+
      '�Intentar sacar el reporte ' + Mensaje + '?',mtConfirmation, mbYesNo, 0)=idNo then
      begin
        // Cerrar excel
        oLibro.close(False);
        oGantt.Quit;
        exit;
      end;

  oGantt.Visible := CheckBox1.Checked;   // Hasta aqu� mostrar de ser necesario la hoja de excel

    // Agregar las columnas de los titulos de las actividades/paquetes
  oDetalle := oLibro.Sheets[2];
  oDetalle.Name := 'ANEXO DETALLADO';

  If Not xEncabezado then   // Colocar el encabezado del diagrama de GANTT
      begin
        // Cerrar excel
        oLibro.close(False);
        oGantt.Quit;
        ShowMessage('El reporte no puede ser generado.');
        exit;
      end;

  ProgressBar1.Min := 0;
  ProgressBar1.Max := roqAnexo.RecordCount;
  ProgressBar1.Visible := True;
  ProgressBar1.Step := 1;

  Ren := rAlto + 1;       // Aqu� inicia el proceso de datos
  Ren1 := rAlto + 1;
  cNivel := -1;

  Busca := TZReadOnlyQuery.Create(Nil);
  Busca.Connection := roqAnexo.Connection;

  roqAnexo.First;
//  while (Not roqAnexo.Eof) and (ren<=500) do
  while (Not roqAnexo.Eof) do
  begin
    ProgressBar1.Position := roqAnexo.RecNo;

    Procede := true;

    if Procede then
    begin
      if oGantt.Visible then
        oGantt.ScreenUpdating := False;

        AcumAnterior := 0;
        AcumAvance := 0;
        OpcAnexo := (roqAnexo.FieldByName('sTipoActividad').AsString = 'Paquete');

        // Filtrar los datos para este paquete para obtener el avance anterior
        roqOrdenAntes.Filtered := False;
        roqOrdenAntes.Filtered := True;
        roqOrdenAntes.First;
        while Not roqOrdenAntes.Eof do
        begin
          AcumAnterior := AcumAnterior + roqOrdenAntes.FieldByName('dAvanceAnexoPondDirecto').AsFloat;

          roqOrdenAntes.Next;
        end;


        roqOrdenHoy.Filtered := False;
        roqOrdenHoy.Filtered := True;
        roqOrdenHoy.First;
        while Not roqOrdenHoy.Eof do
        begin
          AcumAvance := AcumAvance + roqOrdenHoy.FieldByName('dAvanceAnexoPondDirecto').AsFloat;

          roqOrdenHoy.Next;
        end;



        if (Tipo_Programa=0) OR
           ((Tipo_Programa=1) AND (AcumAvance>0)) OR
           ((Tipo_Programa=2) AND (AcumAvance>AcumAnterior)) then
        begin




        //*****************************************************************************************************************************
        //* Inicia proceso de datos generales de partidas
        // Anexo
        SetSelect(oHoja,'A' + Trim(Inttostr(Ren)) + ':A' + Trim(Inttostr(Ren + 3)));
        oGantt.Selection.MergeCells := True;
        oGantt.Selection.Value := roqAnexo.FieldByName('sNumeroActividad').asString;

        SetSelect(oHoja,'B' + Trim(Inttostr(Ren)) + ':B' + Trim(Inttostr(Ren + 3)));
        oGantt.Selection.Value := roqAnexo.FieldByName('mDescripcion').asString;
        oGantt.Selection.MergeCells := True;

        SetSelect(oHoja,'C' + Trim(Inttostr(Ren)) + ':C' + Trim(Inttostr(Ren + 3)));
        oGantt.Selection.Value := ((roqAnexo.FieldByName('dFechaFinal').AsFloat + 1)  - roqAnexo.FieldByName('dFechaInicio').AsFloat) * 24;
        oGantt.Selection.MergeCells := True;

        SetSelect(oHoja,'D' + Trim(Inttostr(Ren)) + ':D' + Trim(Inttostr(Ren + 3)));
        //oGantt.Selection.Value := '';
        oGantt.Selection.MergeCells := True;

      // Detalle
        SetSelect(oDetalle,'A' + Trim(Inttostr(Ren1)) + ':A' + Trim(Inttostr(Ren1 + 3)));
        oGantt.Selection.MergeCells := True;
        oGantt.Selection.Value := roqAnexo.FieldByName('sNumeroActividad').asString;

        SetSelect(oDetalle,'B' + Trim(Inttostr(Ren1)) + ':B' + Trim(Inttostr(Ren1 + 3)));
        oGantt.Selection.Value := roqAnexo.FieldByName('mDescripcion').asString;
        oGantt.Selection.MergeCells := True;

        SetSelect(oDetalle,'C' + Trim(Inttostr(Ren1)) + ':C' + Trim(Inttostr(Ren1 + 3)));
        oGantt.Selection.Value := ((roqAnexo.FieldByName('dFechaFinal').AsFloat + 1)  - roqAnexo.FieldByName('dFechaInicio').AsFloat) * 24;
        oGantt.Selection.MergeCells := True;

        SetSelect(oDetalle,'D' + Trim(Inttostr(Ren1)) + ':D' + Trim(Inttostr(Ren1 + 3)));
        //oGantt.Selection.Value := '';
        oGantt.Selection.MergeCells := True;

      if roqAnexo.FieldByName('sTipoActividad').AsString <> 'Paquete' then
      Begin
        // Colocar formato condicional para actividades
        oHoja.Select;
          oGantt.Selection.FormatConditions.Delete;
          oGantt.Selection.FormatConditions.Add(xlCellValue, xlEqual, '1');
          oGantt.Selection.FormatConditions[1].Font.Bold := True;
          oGantt.Selection.FormatConditions[1].Interior.ColorIndex := 8;

        oDetalle.Select;
          oGantt.Selection.FormatConditions.Delete;
          oGantt.Selection.FormatConditions.Add(xlCellValue, xlEqual, '1');
          oGantt.Selection.FormatConditions[1].Font.Bold := True;
          oGantt.Selection.FormatConditions[1].Interior.ColorIndex := 8;
      End;

      SetSelect(oHoja,'N' + Trim(Inttostr(Ren)) + ':N' + Trim(Inttostr(Ren + 3)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.Value := FechaCadena(roqAnexo.FieldByName('dFechaInicio').AsDateTime,0);

      SetSelect(oHoja,'O' + Trim(Inttostr(Ren)) + ':O' + Trim(Inttostr(Ren + 3)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.Value := FechaCadena(roqAnexo.FieldByName('dFechaFinal').AsDateTime,1);

      SetSelect(oDetalle,'N' + Trim(Inttostr(Ren1)) + ':N' + Trim(Inttostr(Ren1 + 3)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.Value := FechaCadena(roqAnexo.FieldByName('dFechaInicio').AsDateTime,0);

      SetSelect(oDetalle,'O' + Trim(Inttostr(Ren1)) + ':O' + Trim(Inttostr(Ren1 + 3)));
      oGantt.Selection.MergeCells := True;
      oGantt.Selection.Value := FechaCadena(roqAnexo.FieldByName('dFechaFinal').AsDateTime,1);

      Inc(Ren);
      Inc(Ren1);
      //* Terminar proceso de datos generales de partidas
      //*****************************************************************************************************************************

      //* Terminar proceso de datos generales de partidas
      //*****************************************************************************************************************************

      //*****************************************************************************************************************************
      //* Inicia proceso de avance programado
      myInicio := roqAnexo.FieldByName('dFechaInicio').AsFloat;
      myTermino := roqAnexo.FieldByName('dFechaFinal').AsFloat;
      if myInicio < InicioG then
        myInicio := InicioG;
      if myTermino > TerminoG then
        myTermino := TerminoG;


      ColBarra := CalculaPeriodo(InicioGrafica,myInicio,TipoReporte) - 1;
      myWidth := oHoja.Cells[Ren - 1, rAncho + ColBarra + 1].Width;
      desplaze := CalcularDesfase(myInicio,TipoReporte,Inicio_Semana) * myWidth;
      //ColBarra := Trunc(pInicioP - myInicio);
      //ColBarra := Trunc(myInicio - pInicioP);
      case TipoReporte of
        0:  rt1 := (myTermino - myInicio + 1);
        1:  rt1 := (myTermino - myInicio + 1)/7;
        2:  Begin
              if monthof(myInicio)=monthof(myTermino) then
                rt1 := (myTermino - myInicio + 1)/DaysInAMonth(YearOf(myInicio),MonthOf(myInicio))
              else
                rt1 := (MesesEntre(myInicio,myTermino) - 1) +
                       (1 + DaysInAMonth(YearOf(myInicio),MonthOf(myInicio)) - DayOfTheMonth(myInicio))/DaysInAMonth(YearOf(myInicio),MonthOf(myInicio)) +
                       DayOfTheMonth(myTermino)/DaysInAMonth(YearOf(myTermino),MonthOf(myTermino));
            End;
        3:  Begin
              if YearOf(myInicio)=YearOf(myTermino) then
                rt1 := (DayOfTheYear(myTermino) - DayOfTheYear(myInicio) + 1)/DaysInAYear(YearOf(myInicio))
              else
                rt1 := (YearOf(myTermino) - YearOf(myInicio) - 1) +
                       (1 + DaysInAYear(YearOf(myInicio)) - DayOfTheYear(myInicio))/DaysInAYear(YearOf(myInicio))
                        + DayOfTheYear(myTermino)/DaysInAYear(YearOf(myTermino));
            End;
      end;

      //DurBarra := ((myTermino - myInicio) / (pTerminoP - pInicioP) * wCol);
      // if DurBarra < 0 then DurBarra := 0;

      GenerarBarra(oHoja,Ren);
      GenerarBarra(oDetalle,Ren1);
      EnmarcarArea(oHoja,'A'+Trim(IntToStr(Ren-1))+':' + Letra(CeldasAUsar + RAncho) + Trim(IntToStr(Ren+2)),xlMedium,xlHairline, True);
      EnmarcarArea(oDetalle,'A'+Trim(IntToStr(Ren1-1))+':' + Letra(CeldasAUsar + RAncho) + Trim(IntToStr(Ren1+2)),xlMedium,xlHairline, True);

      // Filtrar las actividades de la bitacora
      roqActividades.Filtered := false;
      roqActividades.Filter := 'swbscontrato=' + QuotedStr(roqAnexo.FieldByName('swbs').AsString);
      roqActividades.Filtered := true;

      GenerarBarraReal(oHoja,Ren+1,TipoReporte,17);
      GenerarBarraReal(oDetalle,Ren1+1,TipoReporte,17);

      Inc(Ren);
      Inc(Ren1);
      //* Terminar proceso de avances programados
      //*****************************************************************************************************************************

      Inc(Ren);
      Inc(Ren1);

      SetSelect(oHoja,'E' + Trim(Inttostr(Ren - 3)) + ':E' + Trim(Inttostr(Ren)));
      oGantt.Selection.Value := roqAnexo.FieldValues['dPonderado'] / 100;
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'E' + Trim(Inttostr(Ren1 - 3)) + ':E' + Trim(Inttostr(Ren1)));
      oGantt.Selection.Value := roqAnexo.FieldValues['dPonderado'] / 100;
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'F' + Trim(Inttostr(Ren - 3)) + ':F' + Trim(Inttostr(Ren)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'F' + Trim(Inttostr(Ren1 - 3)) + ':F' + Trim(Inttostr(Ren1)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'G' + Trim(Inttostr(Ren - 3)) + ':G' + Trim(Inttostr(Ren)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'G' + Trim(Inttostr(Ren1 - 3)) + ':G' + Trim(Inttostr(Ren1)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'H' + Trim(Inttostr(Ren - 3)) + ':H' + Trim(Inttostr(Ren)));
      if roqAnexo.FieldByName('sTipoActividad').AsString <> 'Paquete' then
        oGantt.Selection.FormulaR1C1 := '=+RC[5]*RC[-3]';
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'H' + Trim(Inttostr(Ren1 - 3)) + ':H' + Trim(Inttostr(Ren1)));
      if roqAnexo.FieldByName('sTipoActividad').AsString <> 'Paquete' then
        oGantt.Selection.FormulaR1C1 := '=+RC[5]*RC[-3]';
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'I' + Trim(Inttostr(Ren - 3)) + ':I' + Trim(Inttostr(Ren)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'I' + Trim(Inttostr(Ren1 - 3)) + ':I' + Trim(Inttostr(Ren1)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'J' + Trim(Inttostr(Ren - 3)) + ':J' + Trim(Inttostr(Ren)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'J' + Trim(Inttostr(Ren1 - 3)) + ':J' + Trim(Inttostr(Ren1)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'K' + Trim(Inttostr(Ren - 3)) + ':K' + Trim(Inttostr(Ren)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'K' + Trim(Inttostr(Ren1 - 3)) + ':K' + Trim(Inttostr(Ren1)));
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'L' + Trim(Inttostr(Ren - 3)) + ':L' + Trim(Inttostr(Ren)));
      if roqAnexo.FieldByName('sTipoActividad').AsString <> 'Paquete' then
        oGantt.Selection.FormulaR1C1 := '=+SUM(R[3]C[4]:R[3]C[' + IntToStr(aCol + 6) + '])';
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'L' + Trim(Inttostr(Ren1 - 3)) + ':L' + Trim(Inttostr(Ren1)));
      if roqAnexo.FieldByName('sTipoActividad').AsString <> 'Paquete' then
        oGantt.Selection.FormulaR1C1 := '=+SUM(R[3]C[4]:R[3]C[' + IntToStr(aCol + 6) + '])';
      oGantt.Selection.MergeCells := True;

      SetSelect(oHoja,'M' + Trim(Inttostr(Ren - 3)) + ':M' + Trim(Inttostr(Ren)));
      if roqAnexo.FieldByName('sTipoActividad').AsString <> 'Paquete' then
        oGantt.Selection.FormulaR1C1 := '=+SUM(RC[-2]:R[3]C[-1])';
      oGantt.Selection.MergeCells := True;

      SetSelect(oDetalle,'M' + Trim(Inttostr(Ren1 - 3)) + ':M' + Trim(Inttostr(Ren1)));
      if roqAnexo.FieldByName('sTipoActividad').AsString <> 'Paquete' then
        oGantt.Selection.FormulaR1C1 := '=+SUM(RC[-2]:R[3]C[-1])';
      oGantt.Selection.MergeCells := True;

      if roqAnexo.FieldByName('sTipoActividad').AsString = 'Paquete' then
      Begin
        {*************************************************************
          Inicia proceso de paquete de actividades
         *************************************************************}

        // Obtener los avances de la actividad por anexo
        AcumAnterior := 0;
        AcumAvance := 0;
        OpcAnexo := True;

        //if roqAnexo.FieldByName('swbs').AsString = 'A.9.9.1' then
        //  showmessage('rangel');

        // Filtrar los datos para este paquete para obtener el avance anterior
        roqOrdenAntes.Filtered := False;
        roqOrdenAntes.Filtered := True;
        roqOrdenAntes.First;
        while Not roqOrdenAntes.Eof do
        begin
          AcumAnterior := AcumAnterior + roqOrdenAntes.FieldByName('dAvanceAnexoPondDirecto').AsFloat;

          roqOrdenAntes.Next;
        end;

        roqOrdenHoy.Filtered := False;
        roqOrdenHoy.Filtered := True;
        roqOrdenHoy.First;
        while Not roqOrdenHoy.Eof do
        begin
          AcumAvance := AcumAvance + roqOrdenHoy.FieldByName('dAvanceAnexoPondDirecto').AsFloat;

          roqOrdenHoy.Next;
        end;

        // Colocar el avance general
        SetSelect(oHoja,'H' + Trim(Inttostr(Ren - 3)) + ':' + 'H' + Trim(Inttostr(Ren)));
        oGantt.Selection.FormulaR1C1 := '=+RC[+5]*RC[-3]';

        SetSelect(oDetalle,'H' + Trim(Inttostr(Ren1 - 3)) + ':' + 'H' + Trim(Inttostr(Ren1)));
        oGantt.Selection.FormulaR1C1 := '=+RC[+5]*RC[-3]';

        SetSelect(oHoja,'K' + Trim(Inttostr(Ren)) + ':K' + Trim(Inttostr(Ren)));
        oGantt.Selection.Value := AcumAnterior / 100;

        SetSelect(oDetalle,'K' + Trim(Inttostr(Ren1)) + ':K' + Trim(Inttostr(Ren1)));
        oGantt.Selection.Value := AcumAnterior / 100;

        SetSelect(oHoja,'L' + Trim(Inttostr(Ren)) + ':L' + Trim(Inttostr(Ren)));
        oGantt.Selection.FormulaR1C1 := '=+RC[+1]-RC[-1]';

        SetSelect(oDetalle,'L' + Trim(Inttostr(Ren1)) + ':L' + Trim(Inttostr(Ren1)));
        oGantt.Selection.FormulaR1C1 := '=+RC[+1]-RC[-1]';

        // Colocar avance general
        SetSelect(oHoja,'M' + Trim(Inttostr(Ren)) + ':M' + Trim(Inttostr(Ren)));
        oGantt.Selection.Value := AcumAvance / 100;
        SetSelect(oDetalle,'M' + Trim(Inttostr(Ren1)) + ':M' + Trim(Inttostr(Ren1)));
        oGantt.Selection.Value := AcumAvance / 100;
      End
      Else
      Begin
        {***********************************************************************
         Inicia proceso de actividades
         ***********************************************************************}

        // Filtrar los datos para este paquete para obtener el avance anterior
        AcumAnterior := 0;
        AcumAvance := 0;
        OpcAnexo := False;
        Indice := 0;

        // Filtrar los datos para este paquete para obtener el avance anterior
        roqOrdenAntes.Filtered := False;
        roqOrdenAntes.Filtered := True;
        roqOrdenAntes.First;
        wbsxorden := nil;
        while Not roqOrdenAntes.Eof do
        begin
          AcumAnterior := AcumAnterior + roqOrdenAntes.FieldByName('dAvanceAnexoActividad').AsFloat;
          SetLength(wbsxorden,Indice + 1);
          wbsxorden[Indice].wbs := roqOrdenAntes.FieldByName('swbs').AsString;
          wbsxorden[Indice].valor  := roqOrdenAntes.FieldByName('dAvanceAnexoActividad').AsFloat;
          roqOrdenAntes.Next;
          Inc(Indice);
        end;

        primero := true;
        RenAnt := Ren1;
        roqOrdenHoy.Filtered := False;
        roqOrdenHoy.Filtered := True;
        roqOrdenHoy.First;
        while Not roqOrdenHoy.Eof do
        begin
          AcumAvance := AcumAvance + roqOrdenHoy.FieldByName('dAvanceAnexoActividad').AsFloat;

          // Filtrar las actividades de la bitacora
          roqActividades.Filtered := false;
          roqActividades.Filter := 'swbscontrato=' + QuotedStr(roqOrdenHoy.FieldByName('swbscontrato').AsString)
                      + ' and snumeroorden=' + QuotedStr(roqOrdenHoy.FieldByName('snumeroorden').AsString)
                      + ' and swbs=' + QuotedStr(roqOrdenHoy.FieldByName('swbs').AsString);
          roqActividades.Filtered := true;

          if roqActividades.RecordCount > 0 then
          begin
            Inc(Ren1);
            SetSelect(oDetalle,'B' + Trim(IntToStr(Ren1)) + ':' + 'B' + Trim(IntToStr(Ren1)));
            oGantt.Selection.Value := FORMAT('%s (%s)',
                                    [roqOrdenHoy.FieldByName('snumeroorden').AsString,
                                     roqOrdenHoy.FieldByName('swbs').AsString]);
            SetSelect(oDetalle,'M' + Trim(IntToStr(Ren1)) + ':' + 'M' + Trim(IntToStr(Ren1)));
            oGantt.Selection.value := roqOrdenHoy.FieldByName('dAvanceOrdenActividad').AsFloat/100;

            // Poner por default a cero
            SetSelect(oDetalle,'K' + Trim(IntToStr(Ren1)) + ':' + 'K' + Trim(IntToStr(Ren1)));
            oGantt.Selection.value := 0;
            // Buscar el avance anterior al wbs correspondiente
            for Cont := 0 to Indice - 1 do
              if wbsxorden[Cont].wbs = roqOrdenHoy.FieldByName('swbs').AsString then
                begin
                  oGantt.Selection.value := wbsxorden[Cont].valor/100;
                end;

            SetSelect(oDetalle,'L' + Trim(Inttostr(Ren1)) + ':L' + Trim(Inttostr(Ren1)));
            oGantt.Selection.FormulaR1C1 := '=+RC[+1]-RC[-1]';


            SetSelect(oDetalle,'A' + Trim(Inttostr(Ren1)) + ':O' + Trim(Inttostr(Ren1)));
            // Poner Color
            if Primero then
              oGantt.Selection.Interior.colorIndex := 35
            else
              oGantt.Selection.Interior.colorIndex := 36;

            oGantt.Selection.Interior.Pattern := xlSolid;

            GenerarBarraReal(oDetalle,Ren1+1,TipoReporte,11);
            EnmarcarArea(oDetalle,'A'+Trim(IntToStr(Ren1))+':' + Letra(CeldasAUsar + RAncho) + Trim(IntToStr(Ren1)),xlMedium,xlHairline, False);
          end;

          //Primero := not Primero;
          roqOrdenHoy.Next;
        end;
        // Colocar el avance general
        SetSelect(oHoja,'H' + Trim(Inttostr(Ren - 3)) + ':' + 'H' + Trim(Inttostr(Ren)));
        oGantt.Selection.FormulaR1C1 := '=+RC[+5]*RC[-3]';

        SetSelect(oHoja,'K' + Trim(Inttostr(Ren)) + ':K' + Trim(Inttostr(Ren)));
        oGantt.Selection.Value := AcumAnterior / 100;

        SetSelect(oDetalle,'K' + Trim(Inttostr(RenAnt)) + ':K' + Trim(Inttostr(RenAnt)));
        oGantt.Selection.Value := AcumAnterior / 100;

        SetSelect(oHoja,'L' + Trim(Inttostr(Ren)) + ':L' + Trim(Inttostr(Ren)));
        oGantt.Selection.FormulaR1C1 := '=+RC[+1]-RC[-1]';

        SetSelect(oDetalle,'L' + Trim(Inttostr(RenAnt)) + ':L' + Trim(Inttostr(RenAnt)));
        oGantt.Selection.FormulaR1C1 := '=+RC[+1]-RC[-1]';

        // Colocar avance general
        SetSelect(oHoja,'M' + Trim(Inttostr(Ren)) + ':M' + Trim(Inttostr(Ren)));
        oGantt.Selection.Value := AcumAvance / 100;

        SetSelect(oDetalle,'M' + Trim(Inttostr(RenAnt)) + ':M' + Trim(Inttostr(RenAnt)));
        oGantt.Selection.Value := AcumAvance / 100;
      End;
      //* Termina proceso de avances reales
      //*****************************************************************************************************************************

      if oGantt.Visible then
      begin
        oGantt.ScreenUpdating := True;
        SetSelect(oHoja,'A' + IntToStr(Ren) + ':A' + IntToStr(Ren));
      end;

      Inc(Ren);
      Inc(Ren1);
    end;
    end;
    roqAnexo.Next;
  end;

  oLibro.SaveAs(fName);
  oGantt.Visible := True;
  oGantt.ScreenUpdating := True;
  oGantt.DisplayAlerts := True;
  Close;
end;

end.
