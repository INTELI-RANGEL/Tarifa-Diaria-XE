unit frm_importaciondedatos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, global, ComCtrls, ZDataset, Math, UnitExcel,
  StdCtrls, ExtCtrls, DBCtrls, db, Menus, OleCtrls, ZAbstractRODataset, 
  frxClass, frxDBSet, Buttons, OleServer, ExcelXP, ComObj, Excel2000, Utilerias,
  masUtilerias, RxMemDS, UnitExcepciones, frm_PopUpImportacion, frm_PopUpImportacionC, DateUtils,
  ZAbstractDataset, StrUtils;

type
  TfrmImportaciondeDatos = class(TForm)
    btnResumido: TBitBtn;
    btnSalir: TBitBtn;
    Label1: TLabel;
    OpenXLS: TOpenDialog;
    tsArchivo: TEdit;
    ExcelApplication1: TExcelApplication;
    ExcelWorkbook1: TExcelWorkbook;
    ExcelWorksheet1: TExcelWorksheet;
    GroupBox1: TGroupBox;
    rPrograma: TRadioButton;
    btnFiles: TBitBtn;
    chkBorrar: TCheckBox;
    ProgressBar1: TProgressBar;
    rbInsumos: TRadioButton;
    rAnexoPersonal: TRadioButton;
    rAnexoEquipo: TRadioButton;
    PopupPrincipal: TPopupMenu;
    Salir1: TMenuItem;
    SaveDialog1: TSaveDialog;
    RxMDValida: TRxMemoryData;
    frxDBValida: TfrxDBDataset;
    RxMDValidasNumeroActividad: TStringField;
    RxMDValidasWbs: TStringField;
    RxMDValidadCantidad: TStringField;
    RxMDValidasuma: TStringField;
    RxMDValidaaMN: TStringField;
    RxMDValidaaDLL: TStringField;
    RxMDValidabMN: TStringField;
    RxMDValidabDLL: TStringField;
    RxMDValidadCantidadAnexo: TStringField;
    frxReporte: TfrxReport;
    RxMDValidadescripcion: TStringField;
    RxMDValidamensaje: TStringField;
    RxMDValidasNumeroOrden: TStringField;
    RxMDValidasWbs2: TStringField;
    rAvances: TRadioButton;
    rAnexoDT: TRadioButton;
    rAnexoDTStruct: TRadioButton;
    rAnexoDE: TRadioButton;
    rAnexoDEDLL: TRadioButton;
    rAnexoDTCia: TRadioButton;
    rAnexoDTOrdenCia: TRadioButton;
    rAnexoHerr: TRadioButton;
    rAnexoBasicos: TRadioButton;
    rConstruye: TRadioButton;
    rbPersonalxP: TRadioButton;
    rbEquipoxP: TRadioButton;
    rHerrxPartida: TRadioButton;
    rBasicosxPart: TRadioButton;
    rbInsumosxP: TRadioButton;
    rbAlcances: TRadioButton;
    rAnexoA: TRadioButton;
    rAnexoDMA: TRadioButton;
    rAnexoDMO: TRadioButton;
    rAnexoDME: TRadioButton;
    RadioButton1: TRadioButton;
    rbDetalleDeActividades: TRadioButton;
    rbEsPer: TRadioButton;
    zq_listadoper: TZQuery;
    zq_Esp: TZQuery;
    zq_compania: TZQuery;
    rbtnPrograma: TRadioButton;
    ArchivoMsP: TFileOpenDialog;
    rAnexoC: TRadioButton;
    rbControlProgramado: TRadioButton;
    rAvancesCedula: TRadioButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnFilesClick(Sender: TObject);
    procedure tnResumidoClick(Sender: TObject);
    procedure btnSalirClick(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure formatoEncabezado();
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure tsArchivoEnter(Sender: TObject);
    procedure tsArchivoExit(Sender: TObject);
    function ValidaAnexosDT(dParamTipo: string): boolean;
    function ValidaAnexosDME(dParamTipo, dParamTabla, dParamId: string): boolean;
    function ValidaAnexosInsumo(): boolean;
    function ValidaAnexosPE(dParamTipo: string): boolean;
    function ValidaAnexosBasicos(dParamTipo: string): boolean;
    function ValidaAnexosPERxP(dParamTipo, sParamTabla, dParamCampo: string): boolean;
    function ValidaAnexosA(): boolean;
    function ValidaAnexosC(dParamAnexo: string): boolean;
    function ValidaAvancesProgramados(): boolean;
    function ValidaDeleteAnexosP(dParamTabla, dParamId, dParamBuscaTabla, dParamBuscaTabla2: string): boolean;
    procedure ColoresErrorExcel(sFila, sColumna: string; iTipo: integer; sMensaje: string);
    procedure ValidaCampo(sTipo, Columna: string; fila: integer; Campo: string; lFechas: boolean; sColAnt: string);
    procedure CuadroColores(sCodigoC1, sCodigoC2, sErrorC1, sErrorC2: string);
    procedure EliminaCuadro(sPosicion: string; iIndice: integer);
    procedure ConstruyeExplosion();
    function PartidasRepetidas(sParamTipo: string): boolean;
    function ValidaAnexoDTCia(): boolean;
    function ValidaAnexoDTOrdenCia(): boolean;
    procedure rAnexoPersonalClick(Sender: TObject);
    procedure importarControlProgramado;
  private
    Procedure ImportOTProject(sFileProject:TFileName;PgbAvance: TProgressBar=nil);

  public
    { Public declarations }
    {funciones para Validacion de Datos}
    procedure CalcDiferenciasOT(lista: TStringList);
    procedure ventasDiferentes(sWBSContrato, suma: string);
    function cantidadesDiferentes(sWBSContrato: string): string;
    procedure acumularDiferencia(suma, sMensaje: string);

  end;

var
  frmImportaciondeDatos: TfrmImportaciondeDatos;
  Excel, Libro, Hoja: Variant;
  columnas: array[1..260] of string;
  MensajePartidas, sDatoNivel: string;
  lContratoActual: boolean;
  flcid: integer;
  CodigoColor: array[1..4] of string;

  recursos: array[1..25000, 1..6] of string;
  registro: integer;

implementation

uses Frm_PopUpImportacionPP,UFunctionsGHH;

{$R *.dfm}


Procedure TfrmImportaciondedatos.ImportOTProject(sFileProject:TFileName;PgbAvance: TProgressBar=nil);
(***********************************************************************************
* Programador : Gamael Hdez Huerta                                                 *
* Proposito   : Funcion para Importar un Programa de Trabajo hecho en Ms Project.  *
*                                                                                  *
* Información Tecnica:                                                             *
*                                                                                  *
*            Task.Summary  Este es el resumen, me indica si es paquete o Actividad *
*            Task.Name     Es la Descripcion de la Actividad                       *
*            Task.start    Me da la fecha de Inicio de la Actividad                *
*            Task.Finish   Me da la fecha de Termino de la Actividad               *
*            Task.wbs      Este trae el Valor Wbs del Project, editable en el mismo*
*            Task.outlinelevel  El Nivel de la Actividad                           *
*            Task.outlineNumber Este es similar al wbs solo que no es editable     *
*      En el Task.text(x) se define el No. de Partida y el Ponderado               *
*                                                                                  *
*      Se Utiliza un Nuevo algoritmo de Ordenamiento.                              *
*      GetITemOrden   de la Unidad  UFunctionsGHH.pas                              *
*                                                                                  *
* Fecha: 01-Sept-2014                                                              *
*                                                                                  *
*                                                                                  *
*                                                                                  *
************************************************************************************)
var
  MsProject:Variant;
  ActProject:Variant;
  Task:Variant;
  error:Boolean;
  reng,Col,x:Integer;
  Resumen:String;
  Descripcion:String;
  wbs:String;
  CadTmp:String;
  iPos:Integer;
  Existir: TrxMemoryData;

  ImpsContrato,
  ImpsNumeroOrden,
  ImpsNumeroActividad,
  ImpdPonderado,
  ImpsMedida,
  sTipo,
  ImpsTipo,
  ImpsWbsAnterior,
  ImpsWbsContrato,
  ImpdVentaMN,
  ImpdVentaDLL,
  sWbs,
  CodErr1,
  CodErr2,
  sSQL,
  ImpiItemOrden,
  ImpsActAnterior,
  ValidaMat,
  MaterialAuto:String;
  iNivel: Byte;

  ImpmDescripcion: WideString;
  ImpdCantidadAnexo:Integer;

  ImpdFechaInicio,
  ImpdFechaFinal:TDateTime;

  SobreTodos: Boolean;
  Resp,
  BotonSelec: Integer;

  ListaActividades: TStringList;
  TmpiItemOrden:String;
  PosPda,PosPonderado:Byte;
begin
  Error:=false;
  if FileExists(sFileProject) then
  begin
    if AnsiEndsText('.mpp',sFileProject) then
    begin
      Existir := TrxMemoryData.Create(nil);
      ListaActividades:=TstringList.Create;
      try
        Existir.FieldDefs.Clear;
        try
          try
            MsProject:=GetActiveOLEObject ('MSProject.Application');
          except
            MsProject:=CreateOleObject ('MSProject.Application');
          end;
        except
          Error:=true;
        end;

        if not error then
        begin
                    // Generar una lista de registros que deben existir
          Existir.Close;
          Existir.FieldDefs.Add('sContrato', ftString, 15);
          Existir.FieldDefs.Add('sIdConvenio', ftString, 5);
          Existir.FieldDefs.Add('sNumeroOrden', ftString, 35);
          Existir.FieldDefs.Add('sWbs', ftString, 100);
          Existir.FieldDefs.Add('sPaquete', ftString, 10);
          Existir.FieldDefs.Add('sNumeroActividad', ftString, 20);
          Existir.FieldDefs.Add('sTipoActividad', ftString, 15);
          Existir.Open;
          Existir.EmptyTable;

          Application.CreateForm(TFrmPopUpImportacionPP,FrmPopUpImportacionPP);
          FrmPopUpImportacionPP.Left := trunc((Screen.Width) / 2) - trunc((FrmPopUpImportacionPP.Width) / 2);
          FrmPopUpImportacionPP.Top := trunc((screen.Height) / 2) - trunc((FrmPopUpImportacionPP.Height) / 2);
          FrmPopUpImportacionPP.Caption := ' Importacion Puntos de Programa (Ms Project)';
          try
            if FrmPopUpImportacionPP.ShowModal=mrOK then
            begin
              ImpsContrato := FrmPopUpImportacionPP.QrContratos.FieldByName('sContrato').AsString;
              ImpsNumeroOrden:=FrmPopUpImportacionPP.QrFolios.FieldByName('sNumeroOrden').AsString;
              PosPda:=FrmPopUpImportacionPP.lCmbPartida.ItemIndex;
              PosPonderado:=FrmPopUpImportacionPP.lCmbPonderado.ItemIndex;
            end
            else
              error:=true;
          finally
            FrmPopUpImportacionPP.Destroy;
          end;


          if not error then
          begin
            MsProject.visible:=true;
            MsProject.FileOpen(ArchivoMsP.FileName);
            ActProject:=MsProject.ActiveProject;
            TmpiItemOrden:='';

            if PgbAvance<>nil then
            begin
              PgbAvance.Max := ActProject.Tasks.Count * 2;
              PgbAvance.Position := 0;
            end;

            for Reng:=0 to ActProject.Tasks.Count-1 do
            begin
              Task:=ActProject.Tasks.item[Reng+1];

              iNivel := Task.outlinelevel;
              iNivel:=(iNivel-1);
              ImpsNumeroActividad :=GetProjectText(Task,PosPda+1);

              ImpmDescripcion := Task.Name;
              ImpdCantidadAnexo :=1;
              if PosPonderado>0 then
                ImpdPonderado := GetProjectText(Task,PosPonderado)
              else
                ImpdPonderado :='0';
              ImpdFechaInicio := Task.start;
              ImpdFechaFinal := Task.Finish;

              if not Task.Summary then
              begin
                ImpsMedida:='Actividad';
                ImpsTipo:='ADM';
                sTipo:='Actividad';
              end
              else
              begin
                ImpsMedida:='';
                ImpsTipo:='ADM';
                sTipo:='Paquete';
              end;
              sWbs:=Task.wbs;
              iPos:=LastDelimiter('.',sWbs);
              CadTmp:='';
              if iPos>0 then
                CadTmp:=AnsiMidStr(sWbs,1,iPos-1);

              ImpsWbsAnterior:=CadTmp;
              ImpsWbsContrato := '';
              ImpdVentaMN := '0.00';
              ImpdVentaDLL := '0.00';
              ImpsActAnterior:='';
              ImpiItemOrden :=GetITemOrden(TmpiItemOrden,'',iNivel);
              TmpiItemOrden:=ImpiItemOrden;// Esto es solo Si el Swbs Posterior no existe,
                                           // como es una importacion desde cero, se da por hecho
                                           // que no hay un registro posterior.
              try
                      // Inserto Datos a la Tabla .....
                CodErr1 := 'Al importar información del programa de trabajo desde EXCEL';
                CodErr2 := 'Al insertar registros de actividadesxorden';

                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                sSQL := 'INSERT INTO actividadesxorden ( sContrato , sNumeroOrden, sIdConvenio, sTipoActividad, sWbsAnterior, ' +
                  'sWbs, sNumeroActividad, iItemOrden , mDescripcion, dFechaInicio, dDuracion, dFechaFinal, ' +
                  'dVentaMN, dVentaDLL, sMedida, dCantidad, dPonderado, iColor, lGerencial, iNivel, mComentarios, ' +
                  'sTipoAnexo, sWbsContrato, sAnexo, sActividadAnterior, lExtraordinario ) ' +
                  'VALUES (:contrato, :orden, :convenio, :tipo, :anterior, :wbs, :actividad, :Item, :Descripcion, ' +
                  ':Inicio, :Duracion, :Final, :MN, :DLL, :Medida, :CantidadAnexo, :Ponderado, :color, :Gerencial, ' +
                  ':Nivel, :Comentarios, :TipoA, :WbsContrato, :Anexo, :ActividadAnterior, :Extraordinario)';
                connection.zCommand.SQL.Add(sSQL);
                Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                Connection.zCommand.Params.ParamByName('orden').DataType := ftString;
                Connection.zCommand.Params.ParamByName('orden').value := ImpsNumeroOrden;
                Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                Connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('tipo').value := sTipo
                else
                  Connection.zCommand.Params.ParamByName('tipo').value := 'Actividad';
                Connection.zCommand.Params.ParamByName('anterior').DataType := ftString;
                Connection.zCommand.Params.ParamByName('anterior').value := Trim(ImpsWbsAnterior);
                Connection.zCommand.Params.ParamByName('wbs').DataType := ftString;
                if Trim(sWbs) <> '' then
                  Connection.zCommand.Params.ParamByName('wbs').value := sWbs
                else
                  Connection.zCommand.Params.ParamByName('wbs').value := Trim(ImpsNumeroActividad);
                Connection.zCommand.Params.ParamByName('WbsContrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('WbsContrato').value := ImpsWbsContrato;
                Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                Connection.zCommand.Params.ParamByName('actividad').value := Trim(ImpsNumeroActividad);
                Connection.zCommand.Params.ParamByName('Item').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Item').value := ImpiItemOrden;
                Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                Connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
                Connection.zCommand.Params.ParamByName('Inicio').AsDateTime := ImpdFechaInicio;

                Connection.zCommand.Params.ParamByName('Duracion').value := ((ImpdFechaFinal) - (ImpdFechaInicio)) + 1; // DaysBetween(StrToDate(ImpdFechaInicio),StrToDate(ImpdFechaFinal) )+1;

                Connection.zCommand.Params.ParamByName('Final').AsDateTime := ImpdFechaFinal;
                Connection.zCommand.Params.ParamByName('MN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('MN').value := ImpdVentaMN;
                Connection.zCommand.Params.ParamByName('DLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('DLL').value := ImpdVentaDLL;
                Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medida').value := Trim(ImpsMedida);
                Connection.zCommand.Params.ParamByName('CantidadAnexo').DataType := ftFloat;
                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('CantidadAnexo').value := 1
                else
                  Connection.zCommand.Params.ParamByName('CantidadAnexo').value := ImpdCantidadAnexo;
                Connection.zCommand.Params.ParamByName('Ponderado').DataType := ftFloat;

                ipos:=AnsiPos('%',ImpdPonderado);
                if iPos>0 then
                  Delete(ImpdPonderado,iPos,Length(ImpdPonderado)-(Ipos-1));

                Connection.zCommand.Params.ParamByName('Ponderado').value := strToFloatDef(ImpdPonderado,0);


                Connection.zCommand.Params.ParamByName('Color').DataType := ftInteger;

                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('Color').value := 12
                else
                  Connection.zCommand.Params.ParamByName('Color').value := 0;
                Connection.zCommand.Params.ParamByName('Gerencial').DataType := ftString;

                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('Gerencial').value := 'Si'
                else
                  Connection.zCommand.Params.ParamByName('Gerencial').value := 'No';

                Connection.zCommand.Params.ParamByName('Nivel').DataType := ftInteger;
                Connection.zCommand.Params.ParamByName('Nivel').value := iNivel;
                Connection.zCommand.Params.ParamByName('Comentarios').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Comentarios').value := '*';
                Connection.zCommand.Params.ParamByName('TipoA').DataType := ftString;
                Connection.zCommand.Params.ParamByName('TipoA').value := 'ADM';
                Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Anexo').value := 'C';
                Connection.zCommand.Params.ParamByName('ActividadAnterior').DataType := ftString;
                Connection.zCommand.Params.ParamByName('ActividadAnterior').value := ImpsActAnterior;
                Connection.zCommand.Params.ParamByName('Extraordinario').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Extraordinario').value := 'No';
                connection.zCommand.ExecSQL;
              except
                on e: exception do
                begin
                      // Verificar si se encontró una duplicidad de registros
                  if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                  begin
                          // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro

                    if not SobreTodos then
                      Resp := MessageDlg('El punto de programa de EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsNumeroOrden + ' - ' + Global_Convenio + ' - ' + sWbs + ' - ' + ImpsNumeroActividad + #10 + #10 +
                        '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);
                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('UPDATE actividadesxorden SET sWbsAnterior = :anterior, iItemOrden = :Item, mDescripcion = :Descripcion, ' +
                        'dFechaInicio = :Inicio, dDuracion = :Duracion, dFechaFinal = :Final, dVentaMN = :MN, dVentaDLL = :DLL, ' +
                        'sMedida = :Medida, dCantidad = :CantidadAnexo, dPonderado = :Ponderado, iColor = :color, lGerencial = :Gerencial, ' +
                        'iNivel = :Nivel, mComentarios = :Comentarios, sTipoAnexo = :TipoA, sWbsContrato = :WbsContrato, sAnexo = :Anexo, ' +
                        'sActividadAnterior = :ActividadAnterior, lExtraordinario =:Extraordinario ' +
                        'WHERE sContrato = :Contrato and sIdConvenio = :Convenio and sNumeroOrden = :Orden and sWbs = :Wbs and sNumeroActividad = :Actividad and sTipoActividad = :Tipo');

                      Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                      Connection.zCommand.Params.ParamByName('orden').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('orden').value := ImpsNumeroOrden;
                      Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                      Connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                      if sTipo = 'Paquete' then
                        Connection.zCommand.Params.ParamByName('tipo').value := sTipo
                      else
                        Connection.zCommand.Params.ParamByName('tipo').value := 'Actividad';
                      Connection.zCommand.Params.ParamByName('anterior').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('anterior').value := Trim(ImpsWbsAnterior);
                      Connection.zCommand.Params.ParamByName('wbs').DataType := ftString;
                      if Trim(sWbs) <> '' then
                        Connection.zCommand.Params.ParamByName('wbs').value := sWbs
                      else
                        Connection.zCommand.Params.ParamByName('wbs').value := Trim(ImpsNumeroActividad);
                      Connection.zCommand.Params.ParamByName('WbsContrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('WbsContrato').value := ImpsWbsContrato;
                      Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('actividad').value := Trim(ImpsNumeroActividad);
                      Connection.zCommand.Params.ParamByName('Item').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Item').value := ImpiItemOrden;
                      Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                      Connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
                      Connection.zCommand.Params.ParamByName('Inicio').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('Inicio').value := (ImpdFechaInicio);
                      Connection.zCommand.Params.ParamByName('Duracion').DataType := ftInteger;
                      Connection.zCommand.Params.ParamByName('Duracion').value := ((ImpdFechaFinal) - (ImpdFechaInicio)) + 1;
                      Connection.zCommand.Params.ParamByName('Final').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('Final').value := (ImpdFechaFinal);
                      Connection.zCommand.Params.ParamByName('MN').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('MN').value := ImpdVentaMN;
                      Connection.zCommand.Params.ParamByName('DLL').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DLL').value := ImpdVentaDLL;
                      Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Medida').value := Trim(ImpsMedida);
                      Connection.zCommand.Params.ParamByName('CantidadAnexo').DataType := ftFloat;
                      if sTipo = 'Paquete' then
                        Connection.zCommand.Params.ParamByName('CantidadAnexo').value := 1
                      else
                        Connection.zCommand.Params.ParamByName('CantidadAnexo').value := ImpdCantidadAnexo;
                      Connection.zCommand.Params.ParamByName('Ponderado').DataType := ftFloat;


                      ipos:=AnsiPos('%',ImpdPonderado);
                      if iPos>0 then
                        Delete(ImpdPonderado,iPos,Length(ImpdPonderado)-(Ipos-1));

                      Connection.zCommand.Params.ParamByName('Ponderado').value := sTrToFloatDef(ImpdPonderado,0);
                      Connection.zCommand.Params.ParamByName('Color').DataType := ftInteger;

                      if sTipo = 'Paquete' then
                        Connection.zCommand.Params.ParamByName('Color').value := 12
                      else
                        Connection.zCommand.Params.ParamByName('Color').value := 0;
                      Connection.zCommand.Params.ParamByName('Gerencial').DataType := ftString;

                      if sTipo = 'Paquete' then
                        Connection.zCommand.Params.ParamByName('Gerencial').value := 'Si'
                      else
                        Connection.zCommand.Params.ParamByName('Gerencial').value := 'No';

                      Connection.zCommand.Params.ParamByName('Nivel').DataType := ftInteger;
                      Connection.zCommand.Params.ParamByName('Nivel').value := iNivel;
                      Connection.zCommand.Params.ParamByName('Comentarios').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Comentarios').value := '*';
                      Connection.zCommand.Params.ParamByName('TipoA').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('TipoA').value := ImpsTipo;
                      Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Anexo').value := 'C';
                      Connection.zCommand.Params.ParamByName('ActividadAnterior').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('ActividadAnterior').value := ImpsActAnterior;
                      Connection.zCommand.Params.ParamByName('Extraordinario').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Extraordinario').value := 'No';
                      Connection.zCommand.ExecSQL;
                    end; //termina duplicate
                  end
                  else
                    raise;
                end;
              end;

                  // Cargar la lista de registros procesados
              Existir.Append;
              Existir.FieldByName('sContrato').AsString := global_contrato;
              Existir.FieldByName('sIdConvenio').AsString := global_convenio;
              Existir.FieldByName('sNumeroOrden').AsString := ImpsNumeroOrden;
              Existir.FieldByName('sWbs').AsString := sWbs;
              Existir.FieldByName('sPaquete').AsString := '0';
              Existir.FieldByName('sNumeroActividad').AsString := ImpsNumeroActividad;
              Existir.FieldByName('sTipoActividad').AsString := sTipo;
              Existir.Post;


                  {Se crea lista con actividades...}
              if ListaActividades.IndexOf(Trim(ImpsWbsContrato)) < 0 then
                ListaActividades.Add(Trim(ImpsWbsContrato));

                  //Agregar las unidades de medida a la configuracion en automatico...
              CodErr1 := 'Al registrar datos de programa de trabajo';
              CodErr2 := 'Al actualizar información de configuracion';
              x := pos(ImpsMedida, ValidaMat);
              if (x < 1) and (trim(ImpsMedida) <> '') then
              begin
                ValidaMat := ValidaMat + ImpsMedida + '|';
                Connection.zCommand.Active := False;
                Connection.zCommand.SQL.Clear;
                Connection.zCommand.SQL.Add('Update configuracion set txtValidaMaterial = :Medidas where sContrato = :Contrato');
                Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
                Connection.zCommand.Params.ParamByName('Medidas').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medidas').Value := ValidaMat;
                Connection.zCommand.ExecSQL;
              end;

              x := pos(ImpsMedida, MaterialAuto);
              if (x < 1) and (trim(ImpsMedida) <> '') then
              begin
                MaterialAuto := MaterialAuto + ImpsMedida + '|';
                Connection.zCommand.Active := False;
                Connection.zCommand.SQL.Clear;
                Connection.zCommand.SQL.Add('Update configuracion set txtMaterialAutomatico = :Medidas where sContrato = :Contrato');
                Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
                Connection.zCommand.Params.ParamByName('Medidas').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medidas').Value := MaterialAuto;
                Connection.zCommand.ExecSQL;
              end;

              ImpdVentaMN := '0';
              ImpdVentaDLL := '0';

              if PgbAvance<>nil then
              begin
                //PgbAvance.Max := ActProject.Tasks.Count * 2;
                PgbAvance.Position := PgbAvance.Position + 1;
              end;
              //ProgressBar1.Max := ProgressBar1.Max + 1;
              //ProgressBar1.Position := ProgressBar1.Position + 1;
            end; // Termino del While

              // Verificar los registros que debería ser eliminados
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Select * from actividadesxorden where sContrato = :Contrato and sIdConvenio = :Convenio');
            Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
            Connection.zCommand.ParamByName('Convenio').AsString := Global_Convenio;
            Connection.zCommand.Open;


            if BotonSelec = mrYes then
            begin
              try
                Kardex('Importacion de Datos', 'Termina Proceso', 'Frente de Trabajo', '', '', '', '','Tarifa Diaria','Importacion de Datos');
              except
                on e: exception do
                begin
                        // Aquí si se debe dejar independiente esta excepción debido a que si no se puede registrar el kardex tampoco se quiere que se cancele todo el proceso.
                  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al registrar en kardex Importacion de Frente de Trabajo', 0);
                end;
              end;
            end;

              {Se manda mensaje al usuario para informar sobre las diferencias..}
            if Assigned(ListaActividades) then
            begin
              CalcDiferenciasOT(ListaActividades);
              if RxMDValida.RecordCount > 0 then
              begin
                MessageDlg('Existen diferencias. Oprima aceptar para ver el reporte.', mtInformation, [mbOk], 0);
                frxReporte.LoadFromFile(global_files + 'validaActOrden.fr3');
                frxReporte.PreviewOptions.MDIChild := True;
                frxReporte.PreviewOptions.Modal := False;
                try
                  frxReporte.PreviewOptions.Maximized := lCheckMaximized;
                except
                  on e: exception do
                    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al verificar reporte al importar programa de trabajo', 0);
                end;
                frxReporte.PreviewOptions.ShowCaptions := False;
                frxReporte.Previewoptions.ZoomMode := zmPageWidth;
                frxReporte.ShowReport;
              end;
            end;
          end;
        end;
      finally
        ListaActividades.Destroy;
        Existir.Destroy;
      end;
    end
    else
      MessageDlg('El Archivo: '+ #13 + #10 + Quotedstr(sFileProject) + ' NO es Valido.',
                  MtError,[MbOk],0);
  end;
end;



procedure TfrmImportaciondeDatos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree;

  try
    connection.zCommand.Active := False;
    Connection.zCommand.SQL.Clear;
    Connection.zCommand.SQL.Add('drop table actividadesxanexo_temp ');
    connection.zCommand.ExecSQL;
  Except

  end;



end;

procedure TfrmImportaciondeDatos.importarControlProgramado;
var
  Cursor: TCursor;
  excel, libro, hoja : Variant;
  partida, ano, mes, dia, diaAnterior : string;
  fila, columna, mesActual, ultimoDia, ultimoMes, barraProgresoMax, columnaPintada : Integer;
  cantidad : Real;
  zqryInsertarControl, sWbs, zqryPartida, zqryActualizarControl, zqryCantidad : TZQuery;
  filaPintada : Boolean;
begin
  //Importa el control Programado de un archivo de excel
  if (tsArchivo.Text = '') or ((tsArchivo.Text <> '') and not FileExists(tsArchivo.Text)) then
  begin
    MessageDlg('El archivo seleccionado es incorrecto, verifique esto e intente de nuevo.', mtWarning, [mbOk], 0);
    Exit;
  end;

  //excel variante
  try
    excel := CreateOleObject('Excel.Application');
    excel.Visible := False;
    excel.displayalerts := False;
    excel.workbooks.open(Trim( tsArchivo.Text ) );
    libro := excel.workbooks[1];
    hoja := libro.sheets[2];
    hoja.select;
  except
    on e:Exception do
    begin
      MessageDlg('No se puede continuar, verifique tener instalada la aplicación Microsoft Excel ', mtInformation, [mbok], 0);
      Exit;
    end;
  end;
  Cursor := Screen.Cursor;
  Fila := 6;
  barraProgresoMax := 0;
  partida := excel.Range['B' + Trim(IntToStr(Fila))].value;

  while partida <> '' do
  begin
    barraProgresoMax := barraProgresoMax + 1;
    partida := excel.Range['B' + Trim(IntToStr(Fila))].value;
    fila := Fila + 1;
  end;

  zqryCantidad := TZQuery.Create(Self);
  zqryCantidad.Connection := connection.zConnection;

  zqryActualizarControl := TZQuery.Create(Self);
  zqryActualizarControl.Connection := connection.zConnection;

  zqryPartida := TZQuery.Create(Self);
  zqryPartida.Connection := connection.zConnection;

  sWbs := TZQuery.Create(Self);
  sWbs.Connection := connection.zConnection;

  zqryInsertarControl := TZQuery.Create(Self);
  zqryInsertarControl.Connection := connection.zConnection;

  //Obtener la partida del archivo de excel
  filaPintada := False;
  Fila := 6;
  columna := 4;

  partida := excel.Range['B' + Trim(IntToStr(Fila))].value;

  with connection.QryBusca do
  begin
    Active := False;
    SQL.Clear;
    SQL.Add('select ' +
            'DAY(LAST_DAY(dFiProgramado)) as ultimoDia, ' +
            'YEAR(dFiProgramado) as anoInicio, ' +
            'MONTH(dFiProgramado) as mesInicio ' +
            'from ordenesdetrabajo ' +
            'where sContrato = :sContrato ');
    Params.ParamByName('sContrato').AsString := global_contrato;
    Open;

    ultimoDia := FieldByName('ultimoDia').AsInteger;
    mesActual := FieldByName('mesInicio').AsInteger;
    ano := FieldByName('anoInicio').AsString;
  end;

  ProgressBar1.Max := barraProgresoMax;

  while partida <> '' do
  begin
    Screen.Cursor := crAppStart;

    zqryPartida.Active := False;
    zqryPartida.SQL.Clear;

    zqryPartida.SQL.Add('select sNumeroActividad from actividadesxanexo where ' +
                        'sNumeroActividad = :sNumeroActividad limit 1');
    zqryPartida.Params.ParamByName('sNumeroActividad').AsString := partida;
    zqryPartida.Open;

    //Si la partida existe en actividadesxanexo, que haga las acciones necesarias
    if zqryPartida.RecordCount <> 0 then
    begin
      dia := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(5))].value;

      sWbs.Active := False;
      sWbs.SQL.Clear;

      sWbs.SQL.Add('select sWbs from actividadesxanexo where sNumeroActividad = :sNumeroActividad limit 1');
      sWbs.Params.ParamByName('sNumeroActividad').AsString := partida;
      sWbs.Open;

      zqryPartida.Active := False;
      zqryPartida.SQL.Clear;

      zqryPartida.SQL.Add('select dCantidad from distribuciondeanexo where ' +
                          'sNumeroActividad = :sNumeroActividad limit 1');
      zqryPartida.Params.ParamByName('sNumeroActividad').AsString := partida;
      zqryPartida.Open;

      dia := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(5))].value;
      cantidad := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(Fila))].value;

      //**************************************************************************************************
      if zqryPartida.RecordCount = 0 then
      begin

        while dia <> '' do
        begin
          diaAnterior := excel.Range[ColumnaNombre(Columna - 1) + Trim(IntToStr(5))].value;

          if (StrToInt(dia) = 1) and (diaAnterior = IntToStr(ultimoDia)) then
          begin
            mesActual := mesActual + 1;
            if mesActual = 13 then
            begin
              mesActual := 1;
              ano := IntToStr( strtoInt(ano) + 1 );
            end;

            with connection.zCommand do
            begin
              Active := False;
              SQL.Clear;
              SQL.Add('select ' +
                      'DAY(LAST_DAY(:fecha)) as ultimoDia');
              Params.ParamByName('fecha').AsString := ano + '/' + IntToStr(mesActual) + '/' + dia;

              Open;
              ultimoDia := FieldByName('ultimoDia').AsInteger;
            end;
          end;

          if (sWbs.RecordCount = 1) and (cantidad <> 0) then
          begin
            //Insertar los datos a la tabla distribuciondeanexo
            try
              zqryInsertarControl.Active := False;
              zqryInsertarControl.SQL.Clear;
              zqryInsertarControl.SQL.Add('insert into distribuciondeanexo(sContrato, sIdConvenio, dIdFecha, sWbs, sNumeroActividad, dCantidad) ' +
                                          'values(:sContrato, :sIdConvenio, :dIdFecha, :sWbs, :sNumeroActividad, :dCantidad)');
              zqryInsertarControl.Params.ParamByName('sContrato').AsString := global_contrato;
              zqryInsertarControl.Params.ParamByName('sIdConvenio').AsString := '' + global_convenio;
              zqryInsertarControl.Params.ParamByName('dIdFecha').AsString := ano + '-' + IntToStr(mesActual) + '-' + dia;
              zqryInsertarControl.Params.ParamByName('sWbs').AsString := sWbs.FieldByName('sWbs').AsString;
              zqryInsertarControl.Params.ParamByName('sNumeroActividad').AsString := partida;
              zqryInsertarControl.Params.ParamByName('dCantidad').AsFloat := cantidad;
              zqryInsertarControl.ExecSQL;
            except
              on e : Exception do
              begin
                showMessage(e.Message);
              end;
            end;
          end;

          columna := columna + 1;
          dia := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(5))].value;
          cantidad := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(Fila))].value;
        end;
      end

      else
      begin
        while dia <> '' do
        begin
          diaAnterior := excel.Range[ColumnaNombre(Columna - 1) + Trim(IntToStr(5))].value;

          if (StrToInt(dia) = 1) and (diaAnterior = IntToStr(ultimoDia)) then
          begin
            mesActual := mesActual + 1;
            if mesActual = 13 then
            begin
              mesActual := 1;
              ano := IntToStr( strtoInt(ano) + 1 );
            end;

            with connection.zCommand do
            begin
              Active := False;
              SQL.Clear;
              SQL.Add('select ' +
                      'DAY(LAST_DAY(:fecha)) as ultimoDia');
              Params.ParamByName('fecha').AsString := ano + '/' + IntToStr(mesActual) + '/' + dia;

              Open;
              ultimoDia := FieldByName('ultimoDia').AsInteger;
            end;
          end;

          zqryCantidad.Active := False;
          zqryCantidad.SQL.Clear;
          zqryCantidad.SQL.Add('select dCantidad from distribuciondeanexo where ' +
                              'sContrato = :sContrato and sIdConvenio = :sIdConvenio and ' +
                              'sNumeroActividad = :sNumeroActividad ' +
                              'and dIdFecha = :dIdFecha');
          zqryCantidad.Params.ParamByName('sContrato').AsString := global_contrato;
          zqryCantidad.Params.ParamByName('sIdConvenio').AsString := '' + global_convenio;
          zqryCantidad.Params.ParamByName('sNumeroActividad').AsString := partida;
          zqryCantidad.Params.ParamByName('dIdFecha').AsString := ano + '-' + IntToStr(mesActual) + '-' + dia;
          zqryCantidad.Open;

          if (zqryCantidad.RecordCount = 1) then
          begin
            //actualizar cantidad
            try
              zqryActualizarControl.Active := False;
              zqryActualizarControl.SQL.Clear;
              zqryActualizarControl.SQL.Add('update distribuciondeanexo SET dCantidad = :dCantidad ' +
                                          'where sContrato = :sContrato and sIdConvenio = :sIdConvenio and ' +
                                          'dIdFecha = :dIdFecha and sWbs = :sWbs and sNumeroActividad = :sNumeroActividad');
              zqryActualizarControl.Params.ParamByName('dCantidad').AsFloat := cantidad + zqryCantidad.FieldByName('dCantidad').AsFloat;
              zqryActualizarControl.Params.ParamByName('sContrato').AsString := global_contrato;
              zqryActualizarControl.Params.ParamByName('sIdConvenio').AsString := '' + global_convenio;
              zqryActualizarControl.Params.ParamByName('dIdFecha').AsString := ano + '-' + IntToStr(mesActual) + '-' + dia;
              zqryActualizarControl.Params.ParamByName('sWbs').AsString := sWbs.FieldByName('sWbs').AsString;
              zqryActualizarControl.Params.ParamByName('sNumeroActividad').AsString := partida;

              zqryActualizarControl.ExecSQL;
            except
              on e : Exception do
              begin
                showMessage(e.Message);
              end;
            end;
          end

          else if cantidad <> 0 then  //Si ese registro no existe xq anteriormente tenia  de cantidad 0, que lo inserte individualmente
          begin
            zqryInsertarControl.Active := False;
            zqryInsertarControl.SQL.Clear;
            zqryInsertarControl.SQL.Add('insert into distribuciondeanexo(sContrato, sIdConvenio, dIdFecha, sWbs, sNumeroActividad, dCantidad) ' +
                                        'values(:sContrato, :sIdConvenio, :dIdFecha, :sWbs, :sNumeroActividad, :dCantidad)');
            zqryInsertarControl.Params.ParamByName('sContrato').AsString := global_contrato;
            zqryInsertarControl.Params.ParamByName('sIdConvenio').AsString := '' + global_convenio;
            zqryInsertarControl.Params.ParamByName('dIdFecha').AsString := ano + '-' + IntToStr(mesActual) + '-' + dia;
            zqryInsertarControl.Params.ParamByName('sWbs').AsString := sWbs.FieldByName('sWbs').AsString;
            zqryInsertarControl.Params.ParamByName('sNumeroActividad').AsString := partida;
            zqryInsertarControl.Params.ParamByName('dCantidad').AsFloat := cantidad;
            zqryInsertarControl.ExecSQL;
          end;

          columna := columna + 1;
          dia := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(5))].value;
          cantidad := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(Fila))].value;
        end;
      end;
    end

    else //Pintar la partida que no exista en actividadesxanexo
    begin
      excel.Range['B' + Trim(IntToStr(Fila))].Interior.ColorIndex := 3;
      filaPintada := True;
    end;

      //**************************************************************************************************

      ultimoDia := connection.QryBusca.FieldByName('ultimoDia').AsInteger;
      ano := IntToStr(connection.QryBusca.FieldByName('anoInicio').AsInteger);
      mesActual := connection.QryBusca.FieldByName('mesInicio').AsInteger;
      Fila := Fila + 1;
      columna := 4;
      partida := excel.Range['B' + Trim(IntToStr(Fila))].value;
      cantidad := excel.Range[ColumnaNombre(Columna) + Trim(IntToStr(Fila))].value;

      ProgressBar1.Position := ProgressBar1.Position + 1;
  end;
  Screen.Cursor := Cursor;

  if filaPintada then
  begin
    excel.Visible := True;
    excel.Range['B' + Trim(IntToStr(Fila + 3))].Interior.ColorIndex := 3;
    excel.Range['C' + Trim(IntToStr(Fila + 3))].Value := 'La partida no existe en la BD';
    excel.Range['C' + Trim(IntToStr(Fila + 3))].Select;
  end;
end;

procedure TfrmImportaciondeDatos.Salir1Click(Sender: TObject);
const
  NombreCols: Array[1..11] of String = ('No. Ficha', 'Nombre', 'Apellido P', 'Apellido M', 'Especialidad', 'Id Personal', 'Rfc', 'Compañia', 'Id Compañia', 'Libreta de Mar', 'Vigencia de la Libreta');
  NombreFields: Array[1..11] of string = ('sIdTripulacion', 'sNombre', 'sApellidoP', 'sApellidoM', '', 'sIdPersonal', 'sRfc', '', 'sIdCompania', 'sLibretadeMar', 'dVigencia');
var
  CadError, OrdenVigencia: string;
  ExcepcionesFields: TStringList;
//////////////////////////////////// PLANTILAS DE IMPORTACION /////////////////////////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    Procedure createComboExcel(Var Hoja: Variant; NombreHoja: String; PosCol: String; ListaDatos:string);
    begin
  
      Hoja.Sheets[NombreHoja].Select;
      Hoja.Range[PosCol].Select;
      hoja.Selection.Validation.Delete;
      hoja.Selection.Validation.add(xlValidateList,AlertStyle := xlValidAlertStop,Operator := xlBetween, Formula1:=ListaDatos);
      hoja.Selection.Validation.IgnoreBlank := True;
      hoja.Selection.Validation.InCellDropdown := True;
      hoja.Selection.Validation.InputTitle := '';
      hoja.Selection.Validation.ErrorTitle := '';
      hoja.Selection.Validation.ErrorMessage := '';
      hoja.Selection.Validation.ShowInput := True;
      hoja.Selection.Validation.ShowError := True;
    end;
    Procedure AsignaFormulas(var Hoja: Variant; NombreHoja: string; Celda: String; Formula: String; Rango_AutoFill: string;
                             sLocked: Boolean; sFormulaOculta: Boolean; OcultarColumna: String);
    begin
      Hoja.Sheets[NombreHoja].Select;
      Hoja.Range[Celda].Select;
      Hoja.Selection.FormulaR1C1 := Formula;
      if Length(Trim(Rango_AutoFill)) > 0 then
        Hoja.Selection.Autofill(Hoja.range[Rango_AutoFill], xlFillDefault);
      if Length(Trim(OcultarColumna)) > 0 then
      begin
        Hoja.Columns[OcultarColumna].select;
        Hoja.Selection.Locked := sLocked;
        Hoja.Selection.FormulaHidden := sFormulaOculta;
        Hoja.Selection.EntireColumn.hidden := sFormulaOculta;
      end;
    end;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena, sColIDName: string;
      fs: tStream;
      i, x, n, j: integer;
      Fecha: TDate;
      Acumulado: Double;
    begin
      // Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 85;
      //Genera Plantilla para Anexo A...
      x := 1;
     {$region 'Plantilla Especialidad P.'}
   //Especialidad del Personal
      if rbEsPer.Checked then
      begin
        try
          i := 0;
          Cursor := Screen.Cursor;
          Screen.Cursor := crAppStart;

          try
            //Columnas descartadas
            ExcepcionesFields := TStringList.Create;
            ExcepcionesFields.Add('5');
            ExcepcionesFields.Add('8');

            zq_listadoper.Active:=False;
            zq_listadoper.Open;

            zq_Esp.Active:=False;
            zq_Esp.ParamByName('Contrato').AsString:=global_contrato_barco;
            zq_Esp.ParamByName('TipoPer').AsString:='-1';
            zq_Esp.ParamByName('Per').AsString:='-1';
            zq_Esp.Open;

            zq_compania.Active:=False;
            zq_compania.Open;

            Excel.ActiveSheet.Name := leftStr('Compañias', 31);

            //Columnas Excel LookUpComboExcel Companias
            Excel.Sheets['Compañias'].Select;
            zq_compania.First;
            while Not zq_compania.Eof do
            begin
              Excel.Cells[zq_compania.RecNo,2] := zq_compania.FieldByName('sIdCompania').AsString;
              Excel.Cells[zq_compania.RecNo,1] := zq_compania.FieldByName('sDescripcion').AsString;
              zq_compania.Next;
            end;

             //Columnas Excel LookUpComboExcel Companias
            Excel.sheets.add;
            Excel.ActiveSheet.Name := LeftStr('Especialidad',31);
            Excel.Sheets['Especialidad'].Select;
            zq_Esp.First;
            while Not zq_Esp.Eof do
            begin
              Excel.Cells[zq_Esp.RecNo,2] := zq_Esp.FieldByName('sIdPersonal').AsString;
              Excel.Cells[zq_Esp.RecNo,1] := zq_Esp.FieldByName('esp').AsString;
              zq_Esp.Next;
            end;

            //Llenar el La plantilla de Excel
            Excel.sheets.add;
            Excel.ActiveSheet.Name := LeftStr('Plantilla',31);
            Excel.Sheets['Plantilla'].Select;

            for i  := 1 to 11 do
              Excel.Cells[1, i] := nombreCols[i];

            if (zq_listadoper.Active) and (zq_listadoper.RecordCount > 0) then
            begin
              zq_listadoper.First;
              while not zq_listadoper.Eof do
              begin
                for i := 1 to 11 do
                  for j := 0 to ExcepcionesFields.Count - 1 do
                    if ExcepcionesFields[j] <> IntToStr(i) then
                      if Length(trim(NombreFields[i])) > 0 then
                        Excel.Cells[zq_listadoper.RecNo + 1, i] := zq_listadoper.FieldByName(NombreFields[i]).AsString
                      else
                        Excel.Cells[zq_listadoper.RecNo + 1, i] := '';
                createComboExcel(Excel, 'Plantilla', 'E' + IntToStr(zq_listadoper.RecNo + 1), '=Especialidad!$A:$A');
                //Excel.Cells[zq_listadoper.RecNo + 1, 5] := zq_listadoper.FieldByName(NombreFields[6]).AsString;
                createComboExcel(Excel, 'Plantilla', 'H' + IntToStr(zq_listadoper.RecNo + 1), '=Compañias!$A:$A');
                //Excel.Cells[zq_listadoper.RecNo + 1, 8] := zq_listadoper.FieldByName(NombreFields[9]).AsString;

                //Comentarios
                Excel.Cells[zq_listadoper.RecNo + 1, 6].AddComment;
                Excel.Cells[zq_listadoper.RecNo + 1, 6].Comment.Visible := False;
                Excel.Cells[zq_listadoper.RecNo + 1, 6].Comment.Text('No tipiar en esta columna, ya que se llena en base a la Especialidad que elija');

                Excel.Cells[zq_listadoper.RecNo + 1, 9].AddComment;
                Excel.Cells[zq_listadoper.RecNo + 1, 9].Comment.Visible := False;
                Excel.Cells[zq_listadoper.RecNo + 1, 9].Comment.Text('No tipiar en esta columna, ya que se llena en base a la Compañía que elija');

                zq_listadoper.Next;
              end;
            end;
            AsignaFormulas(Excel, 'Plantilla', 'F2', '=VLOOKUP(RC[-1], Especialidad!C[-5]:C[-4], 2, FALSE)', 'F2:' + 'F' + IntToStr(zq_listadoper.RecordCount + 1), false, False, 'F:F');
            AsignaFormulas(Excel, 'Plantilla', 'I2', '=VLOOKUP(RC[-1],Compañias!C[-8]:C[-7],2,FALSE)', 'I2:' + 'I' + IntToStr(zq_listadoper.RecordCount + 1), false, False, 'I:I');

            //Encabezado
            Excel.Range['F:F'].Style := 'Énfasis2';
            Excel.Range['I:I'].Style := 'Énfasis2';
            Excel.Range['A1:K1'].select;
            Excel.Range['A1:K1'].Style := 'Énfasis1';
            Excel.Range['A1:K1'].Font.Bold := true;
            Excel.Range['A1:K1'].HorizontalAlignment := xlCenter;
            Excel.Range['A:K'].columns.autofit;

          finally
            Screen.Cursor := Cursor;
          end;
        except
          on e: Exception do
          begin
            ShowMessage(e.Message);
          end;
        end;
      end;
      {$endregion}
     {$region 'Plantilla Anexo'}
      if rAnexoC.Checked then
      begin
        Excel.Columns['A:A'].ColumnWidth := 15;
        Excel.Columns['B:B'].ColumnWidth := 7.29;
        Excel.Columns['C:C'].ColumnWidth := 12.86;
        Excel.Columns['D:D'].ColumnWidth := 10;
        Excel.Columns['E:E'].ColumnWidth := 38;
        Excel.Columns['F:F'].ColumnWidth := 11;
        Excel.Columns['G:O'].ColumnWidth := 12;

      // Colocar los encabezados de la plantilla...
        Hoja.Range['A1:A1'].Select;
        Excel.Selection.Value := 'Contrato';
        FormatoEncabezado;
        Hoja.Range['B1:B1'].Select;
        Excel.Selection.Value := 'Nivel';
        FormatoEncabezado;
        Hoja.Range['C1:C1'].Select;
        Excel.Selection.Value := 'Actividad';
        FormatoEncabezado;
        Hoja.Range['D1:D1'].Select;
        Excel.Selection.Value := 'Especif.';
        FormatoEncabezado;
        Hoja.Range['E1:E1'].Select;
        Excel.Selection.Value := 'Descripcion';
        FormatoEncabezado;
        Hoja.Range['F1:F1'].Select;
        Excel.Selection.Value := 'Medida';
        FormatoEncabezado;
        Hoja.Range['G1:G1'].Select;
        Excel.Selection.Value := 'Cantidad';
        FormatoEncabezado;
        Hoja.Range['H1:H1'].Select;
        Excel.Selection.Value := 'Ponderado';
        FormatoEncabezado;
        Hoja.Range['I1:I1'].Select;
        Excel.Selection.Value := 'Precio MN';
        FormatoEncabezado;
        Hoja.Range['J1:J1'].Select;
        Excel.Selection.Value := 'Precio DLL';
        FormatoEncabezado;
        Hoja.Range['K1:K1'].Select;
        Excel.Selection.Value := 'Fase_Proyecto';
        FormatoEncabezado;
        Hoja.Range['L1:L1'].Select;
        Excel.Selection.Value := 'Fecha_Inicio';
        FormatoEncabezado;
        Hoja.Range['M1:M1'].Select;
        Excel.Selection.Value := 'Fecha_Final';
        FormatoEncabezado;
        Hoja.Range['N1:N1'].Select;
        Excel.Selection.Value := 'Id_Anexo';
        FormatoEncabezado;
        Hoja.Range['O1:O1'].Select;
        Excel.Selection.Value := 'Tipo(PU,ADM)';
        FormatoEncabezado;
        Hoja.Range['P1:P1'].Select;
        Excel.Selection.Value := 'Extraordinaria(Si/No)';
        FormatoEncabezado;
        for i := 2 to 12 do
        begin
          Hoja.Cells[i, 1].Select;
          Excel.Selection.Value := global_contrato;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Cells[i, 9].Select;
          Excel.Selection.Value := 0;
          Hoja.Cells[i, 10].Select;
          Excel.Selection.Value := 0;
        end;
        Hoja.Cells[2, 2].Select;
        Excel.Selection.Value := '0';
        Hoja.Range['A1:P1'].Select;
      end;
      {$endregion}
     {$region 'Plantilla Programa'}
      //Genera plantilla de Programa de Trabajo..
      if rPrograma.Checked then
      begin
        zq_Esp.Active:=False;
        zq_Esp.SQL.Clear;
        zq_Esp.SQL.Add('select sNumeroOrden, sIdFolio from ordenesdetrabajo where sContrato =:contrato and cIdStatus = "P" order by iOrden DESC ');
        zq_Esp.ParamByName('Contrato').AsString:=global_contrato;
        zq_Esp.Open;

         //Columnas Excel LookUpComboExcel Companias
        Excel.sheets.add;
        Excel.ActiveSheet.Name := LeftStr('Folios_Proceso',31);
        Excel.Sheets['Folios_Proceso'].Select;
        zq_Esp.First;
        while Not zq_Esp.Eof do
        begin
          Excel.Cells[zq_Esp.RecNo,1] := zq_Esp.FieldByName('sNumeroOrden').AsString;
          Excel.Cells[zq_Esp.RecNo,2] := zq_Esp.FieldByName('sIdFolio').AsString;
          zq_Esp.Next;
        end;

        Excel.sheets.add;
        Excel.ActiveSheet.Name := LeftStr('PROGRAMA DE TRABAJO',31);
        Excel.Sheets['PROGRAMA DE TRABAJO'].Select;

        for i := 0 to 12 do
        begin
            createComboExcel(Excel, 'PROGRAMA DE TRABAJO', 'C' + IntToStr(i + 1), '=Folios_Proceso!$B:$B');
        end;

        Excel.Sheets['PUNTOS DE PROGRAMA'].delete;
        Excel.Sheets['Folios_Proceso'].visible := False;

        Excel.Columns['A:A'].ColumnWidth := 15;
        Excel.Columns['B:B'].ColumnWidth := 0;
        Excel.Columns['C:C'].ColumnWidth := 43;
        Excel.Columns['D:D'].ColumnWidth := 8;
        Excel.Columns['E:E'].ColumnWidth := 10;
        Excel.Columns['F:F'].ColumnWidth := 40;
        Excel.Columns['G:G'].ColumnWidth := 5.57;
        Excel.Columns['H:K'].ColumnWidth := 12;
        Excel.Columns['L:M'].ColumnWidth := 0;

        // Colocar los encabezados de la plantilla...
        Excel.Cells[1,1].Select;
        Excel.Selection.Value := 'O.T.';
        FormatoEncabezado;
        Excel.Cells[1,2].Select;
        Excel.Selection.Value := 'Reprog.';
        FormatoEncabezado;
        Excel.Cells[1,3].Select;
        Excel.Selection.Value := 'Folio';
        FormatoEncabezado;
        Excel.Cells[1,4].Select;
        Excel.Selection.Value := 'Nivel';
        FormatoEncabezado;
        Excel.Cells[1,5].Select;
        Excel.Selection.NumberFormat := '@';
        Excel.Selection.Value := 'Partida';
        FormatoEncabezado;
        Excel.Cells[1,6].Select;
        Excel.Selection.Value := 'Descripcion';
        FormatoEncabezado;
        Excel.Cells[1,7].Select;
        Excel.Selection.Value := 'Tipo';
        FormatoEncabezado;
        Excel.Cells[1,8].Select;
        Excel.Selection.NumberFormat := '@';
        Excel.Selection.Value := 'Cantidad';
        FormatoEncabezado;
        Excel.Cells[1,9].Select;
        Excel.Selection.NumberFormat := '@';
        Excel.Selection.Value := 'Ponderado';
        FormatoEncabezado;
        Excel.Cells[1,10].Select;
        Excel.Selection.NumberFormat := '@';
        Excel.Selection.Value := 'Fecha_Inicio';
        FormatoEncabezado;
        Excel.Selection.NumberFormat := '@';
        Excel.Cells[1,11].Select;
        Excel.Selection.Value := 'Fecha_Final';
        FormatoEncabezado;
        Excel.Cells[1,12].Select;
        Excel.Selection.Value := 'Tipo(PU,ADM)';
        FormatoEncabezado;
        Excel.Cells[1,13].Select;
        Excel.Selection.Value := 'Anexo';
        FormatoEncabezado;

        for i := 2 to 12 do
        begin
          Excel.Cells[i, 1].Select;
          Excel.Selection.Value := global_contrato;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Excel.Cells[i, 2].Select;
          Excel.Selection.Value := '1';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Excel.Cells[i, 7].Select;
          if i= 2 then
             Excel.Selection.Value := ''
          else
             Excel.Selection.Value := 'Part.';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Excel.Cells[i, 8].Select;
          Excel.Selection.Value := '1';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Excel.Cells[i, 12].Select;
          Excel.Selection.Value := 'ADM';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Excel.Cells[i, 13].Select;
          Excel.Selection.Value := 'C';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';
        end;
        Excel.Cells[2, 4].Select;
        Excel.Selection.Value := '0';
        Excel.Range['C2:C2'].Select;
      end;
      {$endregion}
     {$region 'Plantilla Materiales'}
      //Genera Plantilla de Insumos..
      if rbInsumos.Checked then
      begin
        Excel.Columns['A:A'].ColumnWidth := 15;
        Excel.Columns['B:B'].ColumnWidth := 10;
        Excel.Columns['C:C'].ColumnWidth := 40;
        Excel.Columns['D:L'].ColumnWidth := 12;

      // Colocar los encabezados de la plantilla...
        Hoja.Range['A1:A1'].Select;
        Excel.Selection.Value := 'Id_Insumo';
        FormatoEncabezado;
        Hoja.Range['B1:B1'].Select;
        Excel.Selection.Value := 'Tipo';
        FormatoEncabezado;
        Hoja.Range['C1:C1'].Select;
        Excel.Selection.Value := 'Descripcion';
        FormatoEncabezado;
        Hoja.Range['D1:D1'].Select;
        Excel.Selection.Value := 'Medida';
        FormatoEncabezado;
        Hoja.Range['E1:E1'].Select;
        Excel.Selection.Value := 'Cantidad';
        FormatoEncabezado;
        Hoja.Range['F1:F1'].Select;
        Excel.Selection.Value := 'Cantidad Inst.';
        FormatoEncabezado;
        Hoja.Range['G1:G1'].Select;
        Excel.Selection.Value := 'Fecha';
        FormatoEncabezado;
        Hoja.Range['H1:H1'].Select;
        Excel.Selection.Value := 'Costo MN';
        FormatoEncabezado;
        Hoja.Range['I1:I1'].Select;
        Excel.Selection.Value := 'Costo DLL';
        FormatoEncabezado;
        Hoja.Range['J1:J1'].Select;
        Excel.Selection.Value := 'Venta MN';
        FormatoEncabezado;
        Hoja.Range['K1:K1'].Select;
        Excel.Selection.Value := 'Venta DLL';
        FormatoEncabezado;

        for i := 2 to 12 do
        begin
          Hoja.Cells[i, 2].Select;
          Excel.Selection.Value := 'Material';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Cells[i, 8].Select;
          Excel.Selection.Value := 0;
          Hoja.Cells[i, 9].Select;
          Excel.Selection.Value := 0;
          Hoja.Cells[i, 10].Select;
          Excel.Selection.Value := 0;
          Hoja.Cells[i, 11].Select;
          Excel.Selection.Value := 0;
        end;
        Hoja.Range['A1:K1'].Select;
      end;
      {$endregion}
     {$region 'Plantilla de Personal y Equipo'}
     //Genera plantilla de Personal y de Equipos....
      if (rAnexoPersonal.Checked) or (rAnexoEquipo.Checked) then
      begin
        Excel.Columns['A:A'].ColumnWidth := 15;
        Excel.Columns['B:B'].ColumnWidth := 7.29;
        Excel.Columns['C:C'].ColumnWidth := 15;
        Excel.Columns['D:D'].ColumnWidth := 40;
        Excel.Columns['E:N'].ColumnWidth := 12;

      // Colocar los encabezados de la plantilla...
        Hoja.Range['A1:A1'].Select;
        Excel.Selection.Value := 'Contrato';
        FormatoEncabezado;
        Hoja.Range['B1:B1'].Select;
        if rAnexoPersonal.Checked then
          Excel.Selection.Value := 'Id_Personal'
        else
          Excel.Selection.Value := 'Id_Equipo';
        FormatoEncabezado;
        Hoja.Range['C1:C1'].Select;
        Excel.Selection.Value := 'Ordenamiento';
        FormatoEncabezado;
        Hoja.Range['D1:D1'].Select;
        Excel.Selection.Value := 'Descripcion';
        FormatoEncabezado;
        Hoja.Range['E1:E1'].Select;
        Excel.Selection.Value := 'Medida';
        FormatoEncabezado;
        Hoja.Range['F1:F1'].Select;
        Excel.Selection.Value := 'Cantidad';
        FormatoEncabezado;
        Hoja.Range['G1:G1'].Select;
        Excel.Selection.Value := 'Costo MN';
        FormatoEncabezado;
        Hoja.Range['H1:H1'].Select;
        Excel.Selection.Value := 'Costo DLL';
        FormatoEncabezado;
        Hoja.Range['I1:I1'].Select;
        Excel.Selection.Value := 'Venta MN';
        FormatoEncabezado;
        Hoja.Range['J1:J1'].Select;
        Excel.Selection.Value := 'Venta DLL';
        FormatoEncabezado;
        Hoja.Range['K1:K1'].Select;
        Excel.Selection.Value := 'Fecha_Inicio';
        FormatoEncabezado;
        Hoja.Range['L1:L1'].Select;
        Excel.Selection.Value := 'Fecha_Final';
        FormatoEncabezado;
        Hoja.Range['M1:M1'].Select;
        if rAnexoPersonal.Checked then
          Excel.Selection.Value := 'Id_TipoPersonal'
        else
          Excel.Selection.Value := 'Id_TipoEquipo';
        FormatoEncabezado;
        Hoja.Range['N1:N1'].Select;
        Excel.Selection.Value := 'Jornada';
        FormatoEncabezado;

        for i := 2 to 12 do
        begin
          Hoja.Cells[i, 1].Select;
          Excel.Selection.Value := global_contrato;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Cells[i, 7].Select;
          Excel.Selection.Value := 0;
          Hoja.Cells[i, 8].Select;
          Excel.Selection.Value := 0;
          Hoja.Cells[i, 9].Select;
          Excel.Selection.Value := 0;
          Hoja.Cells[i, 10].Select;
          Excel.Selection.Value := 0;
        end;
        Hoja.Range['A1:N1'].Select;
      end;
      {$endregion}
     {$region 'Plantilla de Avances'}
     //Genera Plantilla para EL AVANCE PROGRAMADO
      if (rAvances.Checked) then
      begin
        Excel.Columns['A:A'].ColumnWidth := 15;
        Excel.Columns['B:I'].ColumnWidth := 12;

      // Colocar los encabezados de la plantilla...
        Hoja.Range['A1:A1'].Select;
        Excel.Selection.Value := 'Contrato';
        FormatoEncabezado;
        Hoja.Range['B1:B1'].Select;
        Excel.Selection.Value := 'Fecha';
        FormatoEncabezado;
        Hoja.Range['C1:C1'].Select;
        Excel.Selection.Value := 'Frente';
        FormatoEncabezado;
        Hoja.Range['D1:D1'].Select;
        Excel.Selection.Value := 'Avance';
        FormatoEncabezado;
        Hoja.Range['E1:E1'].Select;
        Excel.Selection.Value := 'Acumulado';
        FormatoEncabezado;
        Hoja.Range['F1:F1'].Select;
        Excel.Selection.Value := 'AvanceFinanciero';
        FormatoEncabezado;
        Hoja.Range['G1:G1'].Select;
        Excel.Selection.Value := 'Numero Gerencial';
        FormatoEncabezado;
        Hoja.Range['H1:H1'].Select;
        Excel.Selection.Value := 'Duracion';
        FormatoEncabezado;
        Hoja.Range['I1:I1'].Select;
        Excel.Selection.Value := 'Horario';
        FormatoEncabezado;

        Fecha := Now();
        Fecha := IncDay(Fecha);
        Fecha := IncDay(Fecha);
        Acumulado := 0;

        with connection.QryBusca do
        begin
          Active := False;
          SQL.Text := 'select '+
                        '* '+
                      'from avancesglobales '+
                      'where sContrato = :orden and '+
                            'sIdConvenio = :convenio '+
                      'order by sNumeroOrden, dIdFecha, iNumeroGerencial, sHorario';
          ParamByName( 'orden' ).AsString := global_contrato;
          ParamByName( 'convenio' ).AsString := global_convenio;
          Open;
          First;
        end;

        i := 2;
        while not connection.qrybusca.eof do
        begin
          Hoja.Cells[i, 1].Select;
          Excel.Selection.Value := global_contrato;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Fecha := IncDay(Fecha);
          Hoja.Cells[i, 2].Select;
          Excel.Selection.Value := FormatDateTime( 'YYYY-MM-DD', connection.qrybusca.FieldByName( 'dIdFecha' ).AsDateTime );
          FormatoEncabezado;

          Hoja.Cells[i, 3].Select;
          Excel.Selection.Value := connection.QryBusca.FieldByName( 'sNumeroOrden' ).AsString;

          Hoja.Cells[i, 4].Select;
          Excel.Selection.Value := connection.QryBusca.FieldByName( 'dAvancePonderadoDia' ).AsString;

          Acumulado := Acumulado + i;
          Hoja.Cells[i, 5].Select;
          Excel.Selection.Value := connection.QryBusca.FieldByName( 'dAvancePonderadoGlobal' ).AsString;

          Hoja.Cells[i, 6].Select;
          Excel.Selection.Value := connection.QryBusca.FieldByName( 'dAvanceFinanciero' ).AsString;

          Hoja.Cells[i, 7].Select;
          Excel.Selection.Value := connection.QryBusca.FieldByName( 'iNumeroGerencial' ).AsString;

          Hoja.Cells[i, 8].Select;
          Excel.Selection.Value := connection.QryBusca.FieldByName( 'sDuracion' ).AsString;

          Hoja.Cells[i, 9].Select;
          Excel.Selection.Value := connection.QryBusca.FieldByName( 'sHorario' ).AsString;

          connection.QryBusca.Next;

          Inc( i );
        end;
        Hoja.Range['A1:I1'].Select;
      end;
      {$endregion}


    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Excel.ActiveWindow.Zoom := 85;
      Hoja.Select;
      if rAnexoC.Checked then
        Hoja.Name := 'ANEXO C'
      else
        if rAnexoA.Checked then
          Hoja.Name := 'ANEXO A'
        else
          if rAnexoDT.Checked then
            Hoja.Name := 'ANEXO DT'
          else
            if rAnexoDE.Checked then
              Hoja.Name := 'ANEXO DE'
            else
              if rAnexoDEDLL.Checked then
                Hoja.Name := 'ANEXO DE DLL'
              else
                if rbAlcances.Checked then
                  Hoja.Name := 'FASES X PARTIDA'
                else
                  if rbInsumosxP.Checked then
                    Hoja.Name := 'INSUMOS X PARTIDA'
                  else
                    if rbPersonalxP.Checked then
                      Hoja.Name := 'PERSONAL X PARTIDA'
                    else
                      if rbEquipoxP.Checked then
                        Hoja.Name := 'EQUIPO X PARTIDA'
                      else
                        if rHerrxPartida.Checked then
                          Hoja.Name := 'HERRAMIENTA X PARTIDA'
                        else
                          if rBasicosxPart.Checked then
                            Hoja.Name := 'BASICOS X PARTIDA'
                          else
                            if rPrograma.Checked then
                              Hoja.Name := 'PUNTOS DE PROGRAMA'
                            else
                              if rbInsumos.Checked then
                                Hoja.Name := 'MATERIALES'
                              else
                                if rAnexoPersonal.Checked then
                                  Hoja.Name := 'PERSONAL'
                                else
                                  if rAnexoHerr.Checked then
                                    Hoja.Name := 'HERRAMIENTA'
                                  else
                                    if rAnexoBasicos.Checked then
                                      Hoja.Name := 'BASICOS'
                                    else
                                      if rAnexoEquipo.Checked then
                                        Hoja.Name := 'EQUIPO'
                                      else
                                        if rAnexoDTStruct.Checked then
                                          Hoja.Name := 'ANEXO DT Estructurado'
                                        else
                                          if rAvances.Checked then
                                            Hoja.Name := 'AVANCES PROGRAMADOS'
                                          else
                                            if rAnexoDTCia.Checked then
                                              Hoja.Name := 'ANEXO DT (CIA)'
                                            else
                                                if rAnexoDTCia.Checked then
                                                  Hoja.Name := 'DETALLES POR ACTIVIDAD';



      DatosPlantilla;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      Excel.DisplayAlerts := True;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al generar la hoja de movimientos de EXISTENCIAS:' + #10 + #10 + e.Message;
        Excel.DisplayAlerts := True;
      end;
    end;

    Result := Resultado;
  end;

begin
  // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
  if not SaveDialog1.Execute then
    Exit;

  // Generar el ambiente de excel
  try
    Excel := CreateOleObject('Excel.Application');
  except
    on e: exception do begin
      FreeAndNil(Excel);
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al importar datos', 0);
      Exit;
    end;
  end;

  Excel.Visible := True;
  Excel.DisplayAlerts := False;
  Excel.ScreenUpdating := True;

  Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

  // Verificar si cuenta con las hojas necesarias
  while Libro.Sheets.Count < 2 do
    Libro.Sheets.Add;

  // Verificar si se pasa de hojas necesarias
  Libro.Sheets[1].Select;
  while Libro.Sheets.Count > 1 do
    Excel.ActiveWindow.SelectedSheets.Delete;

  // Proceder a generar la hoja REPORTE
  CadError := '';

  if GenerarPlantilla then
    // Grabar el archivo de excel con el nombre dado
    messageDlg('La Plantilla se Genero Correctamente!', mtConfirmation, [mbOk], 0);

  Excel := '';

  if CadError <> '' then
    showmessage(CadError)
  else
    tsarchivo.Text := SaveDialog1.FileName;

end;

procedure TfrmImportaciondeDatos.btnFilesClick(Sender: TObject);
var
  x, y, i: Integer;
begin
  if rbtnPrograma.Checked then
  begin
    if ArchivoMsP.Execute then
      TsArchivo.Text:=ArchivoMsP.FileName;
  end
  else
  begin
    OpenXLS.Title := 'Inserta Archivo de Consulta';
    if OpenXLS.Execute then
    begin
      tsArchivo.Text := OpenXLS.FileName;

       // soad - > Llenado del array..
      for x := 1 to 26 do
        columnas[x] := Chr(64 + x);

      i := 27;
      for x := 1 to 9 do
      begin
        for y := 1 to 26 do
        begin
          columnas[i] := Chr(64 + x) + Chr(64 + y);
          i := i + 1;
        end;
      end;
    end;
  end;
end;

procedure TfrmImportaciondeDatos.acumularDiferencia(suma, sMensaje: string);
begin
  RxMDValida.Append;
  RxMDValida.FieldByName('sNumeroActividad').Value := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
  RxMDValida.FieldByName('sWbs').Value := connection.QryBusca.FieldByName('sWbs').AsString;
  RxMDValida.FieldByName('dCantidad').Value := connection.QryBusca.FieldByName('dCantidad').AsString;
  RxMDValida.FieldByName('suma').Value := suma;
  RxMDValida.FieldByName('aMN').Value := connection.QryBusca.FieldByName('aMN').AsString;
  RxMDValida.FieldByName('aDLL').Value := connection.QryBusca.FieldByName('aDLL').AsString;
  RxMDValida.FieldByName('dCantidadAnexo').Value := connection.QryBusca.FieldByName('dCantidadAnexo').AsString;
  RxMDValida.FieldByName('bMN').Value := connection.QryBusca.FieldByName('bMN').AsString;
  RxMDValida.FieldByName('bDLL').Value := connection.QryBusca.FieldByName('bDLL').AsString;
  RxMDValida.FieldByName('descripcion').Value := connection.QryBusca.FieldByName('descripcion').AsString;
  RxMDValida.FieldByName('mensaje').Value := sMensaje;
  RxMDValida.FieldByName('sNumeroOrden').Value := connection.QryBusca.FieldByName('sNumeroOrden').AsString;
  RxMDValida.FieldByName('sWbs2').Value := connection.QryBusca.FieldByName('wbs2').AsString;
  RxMDValida.Post;
end;

function TfrmImportaciondeDatos.cantidadesDiferentes(sWBSContrato: string): string;
var
  sSQL: string;
begin
  result := '';

  sSQL := 'SELECT ' +
    'sum(a.dCantidad) as suma ' +
    'FROM actividadesxorden a ' +
    'INNER JOIN  actividadesxanexo b ' +
    'ON a.sContrato = b.sContrato ' +
    'AND a.sIdConvenio = b.sIdConvenio ' +
    'AND a.sWbsContrato = b.sWbs ' +
    'AND a.sTipoActividad = "Actividad" ' +
    'WHERE a.sContrato = :contrato ' +
    'AND a.sIdConvenio = :convenio ' +
    'AND a.sWbsContrato = :wbscontrato ' +
    'AND a.sTipoActividad = "Actividad"';

  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add(sSQL);
  connection.QryBusca.ParamByName('wbscontrato').Value := sWBSContrato;
  connection.QryBusca.ParamByName('contrato').Value := global_contrato;
  connection.QryBusca.ParamByName('convenio').Value := global_convenio;
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
    result := connection.QryBusca.FieldByName('suma').AsString
end;

procedure TfrmImportaciondeDatos.ventasDiferentes(sWBSContrato, suma: string);
var
  sSQL: string;
  lError1, lError2: boolean;
begin
  sSQL := 'SELECT ' +
    'b.sNumeroActividad, b.sWbs, a.dCantidad, substr(b.mDescripcion,1,255) as descripcion, ' +
    'a.dVentaMN as aMN, a.dVentaDLL as aDLL, a.sTipoActividad, a.sNumeroOrden, a.sWbs as wbs2, ' +
    'b.dCantidadAnexo,  b.dVentaMN as bMN, b.dVentaDLL as bDLL  ' +
    'FROM actividadesxorden a ' +
    'INNER JOIN  actividadesxanexo b ' +
    'ON a.sContrato = b.sContrato ' +
    'AND a.sIdConvenio = b.sIdConvenio ' +
    'AND a.sWbsContrato = b.sWbs ' +
    'AND a.sTipoActividad = "Actividad" ' +
    'WHERE a.sContrato = :contrato ' +
    'AND a.sIdConvenio = :convenio ' +
    'AND a.sWbsContrato = :wbscontrato ' +
    'AND a.sTipoActividad = "Actividad" ' +
    'ORDER BY b.sWbs';

  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add(sSQL);
  connection.QryBusca.ParamByName('wbscontrato').Value := sWBSContrato;
  connection.QryBusca.ParamByName('contrato').Value := global_contrato;
  connection.QryBusca.ParamByName('convenio').Value := global_convenio;
  connection.QryBusca.Open;

  lError1 := false;
  lError2 := false;
  while not connection.QryBusca.Eof do begin
    if (connection.QryBusca.FieldByName('aMN').Value <>
      connection.QryBusca.FieldByName('bMN').Value)
      or (connection.QryBusca.FieldByName('aDLL').Value <>
      connection.QryBusca.FieldByName('bDLL').Value) then begin
      acumularDiferencia(suma, 'Existe diferencia entre los valores de ventas');
      lError1 := true;
    end
    else begin
      if (not lError1) and (not lError2) then begin
        if (connection.QryBusca.FieldByName('dCantidadAnexo').Value <> suma) then
          lError2 := true;
      end;
    end;
    connection.QryBusca.Next;
  end;
  if (not lError1) and (lError2) then begin
    connection.QryBusca.First;
    while not connection.QryBusca.Eof do begin
      acumularDiferencia(suma, 'Existe diferencia entre la suma total de las cantidades y la cantidad del anexo');
      connection.QryBusca.Next;
    end;
  end;
end;

procedure TfrmImportaciondeDatos.CalcDiferenciasOT(lista: TStringList);
var
  ii: integer;
begin
  RxMDValida.Active := True;
  if RxMDValida.RecordCount > 0 then
    RxMDValida.EmptyTable;
  for ii := 0 to Lista.Count - 1 do begin
    if Lista.Strings[ii] <> '' then
      ventasDiferentes(Lista.Strings[ii], cantidadesDiferentes(Lista.Strings[ii]));
  end;
end;


procedure TfrmImportaciondeDatos.tnResumidoClick(Sender: TObject);
var
  //Ventanas
  flcid, Fila, iFila: Integer;
  ListaActividades: TStringList;

  //Anexos C
  ImpmDescripcion: WideString;
  iNivel: Byte;

  sValue, ImpsContrato, ImpsConvenio, ImpsNumeroActividad, ImpsPartida, ImpsEspecificacion, ImpdCantidadAnexo, ImpdFechaInicio, ImpdFechaFinal,
  ImpdCostoMN, ImpdCostoDLL, ImpdVentaMN, ImpdVentaDLL, ImpdPonderado, ImpsWbsAnterior, ImpsMedida, ImpsAnexo,
  ImpsWbsContrato, ImpiItemOrden, sItemOrdenAnterior, ImpsFase, ImpsSistema, ImpsExtraordinario, sTipo, sWbs,

  //Anexo A
  ImpsIsometrico, ImpiRevision, ImpsGrupo, ImpsWbs, ImpsPlataforma,

  //Programa
  ImpsActAnterior, ImpsNumeroOrden,

  //Insumos
  ImpdInstalado, sIdAlmacen, ValidaMat, MaterialAuto: string;
  iRegistros, iColumna: integer;

  //Anexo DT
  sFecha, ImpsTipo, sTabla, sRecorrido: string;
  impiFase, impiFaseAnterior: integer;
  dVentaMN, dVentaDLL,
  ImpfValor: Currency;
  dmpiAvance: Double;
  myYear, myMonth, myDay: Word;
  dFecha: TDateTime;
  MiValor, MiAnexo,
  MiValor2: string;

  //Analizador.
  DatoE, sIdRecurso, sSQL, cadena, cadena2: string;
  lActualiza, lObtenerDeAnexo, lEncuentra: Boolean;
  SobreTodos: Boolean;
  Existir: TrxMemoryData;
  Resp: Integer;

  arrFechas: array[1..1000] of string;
  paquete: array[1..3000, 1..3] of string;

  I, x, t, BotonSelec: Integer;

  //Ventana Robert importa informacion
  lImportarDescripcion, lImportarPreciosMN, lImportarPreciosDLL, lImportarMedida,
  lImportarFechaIni, lImportarFechaFin, lImportarCantidad: boolean;
  lMsExcel, lOrdenInteligent: boolean;
  CodErr1, CodErr2: string;

  ImpsIdConvenio,sTmpContrato: string;
  ImpdAvancePonderadoDia, ImpdAvancePonderadoGlobal, ImpdAvanceFinanciero: Double;

  ImpNumeroGerencial : Integer;//Sam
  ImpDuracion,
  ImpHorario : string;//Sam

  ImpdCantidad, ImpdMontoPonderado : Double;
  sTmpOrden: string;

  NuevaFechaInicial,
  NuevaFechaFinal: TDateTime;

  dVigencia: TDate;

  sNombre, sApellidoP, sApellidoM, sIdPersonal, sRfc, sIdCompania, sLibretadeMar: string;

begin
   {cODIGOC COLORES}
  CodigoColor[1] := '';
  CodigoColor[2] := '';
  CodigoColor[3] := '';
  CodigoColor[4] := '';

   // Verificar si se ha seleccionado algun archivo
  if (tsArchivo.Text = '') or ((tsArchivo.Text <> '') and not FileExists(tsArchivo.Text)) then
  begin
    MessageDlg('El archivo seleccionado es incorrecto, verifique esto e intente de nuevo.', mtWarning, [mbOk], 0);
    Exit;
  end;

  if not rbtnPrograma.Checked then
  begin
  {$Region 'Importaciones en Excel'}
     //Busqueda de unidades de medida...
    Connection.zCommand.Active := False;
    Connection.zCommand.SQL.Clear;
    Connection.zCommand.SQL.Add(' select txtValidaMaterial, txtMaterialAutomatico from configuracion where sContrato = :Contrato ');
    Connection.zCommand.Params.ParamByName('Contrato').DataType :=     ftString;
    Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.zCommand.Open;
    if Connection.zCommand.RecordCount > 0 then
    begin
      ValidaMat := Connection.zCommand.FieldValues['txtValidaMaterial'];
      MaterialAuto := Connection.zCommand.FieldValues['txtMaterialAutomatico'];
    end;

     //Asignacion de la ruta del archivo de Excel...
    try
      Existir := TrxMemoryData.Create(nil);
      Existir.FieldDefs.Clear;

      try
        Connection.zConnection.StartTransaction;

        CodErr1 := 'Al generar ambiente de EXCEL para levantamiento de información';
        CodErr2 := 'Al intentar modificar atributos de EXCEL';

        flcid := GetUserDefaultLCID;
        ExcelApplication1.Connect;
        ExcelApplication1.Visible[flcid] := true;
        ExcelApplication1.UserControl := true;

        ExcelWorkbook1.ConnectTo(ExcelApplication1.Workbooks.Open(tsArchivo.Text,
          emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam,
          emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, flcid));

        {Antes de iniciar peguntamos al Usuraio si Tomamos Datos del Contrato Actual o el de Excel..}
        lContratoActual := False;
        lContratoActual := True;

        ExcelWorksheet1.ConnectTo(ExcelWorkbook1.Sheets.Item[1] as ExcelWorkSheet);
        Fila := 1;
        SobreTodos := False;

        {$REGION 'Especialidad Personal'}
        if rbEsPer.Checked then
        begin
          try
            ProgressBar1.Max := ExcelApplication1.Rows.CurrentRegion.Rows.Count;
            ProgressBar1.Position:=0;
            Fila := 2;
            //Procedemos a leer el archivo de Excel..
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

            if not chkBorrar.Checked then
            begin
              connection.QryBusca.Active:=false;
              connection.QryBusca.SQL.Text:='select * from tripulacion_listado where sContrato=:Contrato';
              connection.QryBusca.ParamByName('Contrato').AsString:= global_contrato_barco;
              connection.QryBusca.Open;
              if connection.QryBusca.RecordCount>0 then
                if MessageDlg('Desea Eliminar la Lista de Personal Existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                  chkBorrar.Checked := True;

            end;

            if chkBorrar.Checked then
            begin
              connection.QryBusca.Active:=false;
              connection.QryBusca.SQL.Text:='delete from tripulacion_listado where sContrato=:Contrato';
              connection.QryBusca.ParamByName('Contrato').AsString:= global_contrato_barco;
              connection.QryBusca.ExecSQL;
            end;

            while (sValue <> '') do
            begin
              sNombre := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              sApellidoP := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
              sApellidoM   := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              sIdPersonal  := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
              sRfc := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
              sIdCompania := ExcelWorkSheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;
              sLibretadeMar := ExcelWorkSheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
              dVigencia := ExcelWorkSheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;

              connection.zCommand.Active := False;
              Connection.zCommand.SQL.Clear;
              Connection.zCommand.SQL.Add('INSERT INTO tripulacion_listado ( sContrato , sIdTripulacion, sNombre, sApellidoP, sApellidoM, sIdPersonal, sRfc, sIdCompania, sLibretadeMar, dVigencia) ' +
                                          'VALUES (:contrato, :IdTripulacion, :Nombre , :ApellidoP, :ApellidoM, :IdPersonal, :Rfc, :IdCompania, :LibretadeMar, :Vigencia)');
              Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('contrato').value    := global_contrato_barco;
              Connection.zCommand.Params.ParamByName('IdTripulacion').DataType    := ftString;
              Connection.zCommand.Params.ParamByName('IdTripulacion').value       := sValue;
              Connection.zCommand.Params.ParamByName('Nombre').DataType    := ftString;
              Connection.zCommand.Params.ParamByName('Nombre').value       := sNombre;
              Connection.zCommand.Params.ParamByName('ApellidoP').DataType:= ftString;
              Connection.zCommand.Params.ParamByName('ApellidoP').value   := sApellidoP;
              Connection.zCommand.Params.ParamByName('ApellidoM').DataType := ftString;
              Connection.zCommand.Params.ParamByName('ApellidoM').value    := sApellidoM;
              Connection.zCommand.Params.ParamByName('IdPersonal').DataType       := ftString;
              Connection.zCommand.Params.ParamByName('IdPersonal').value          := sIdPersonal;
              Connection.zCommand.Params.ParamByName('Rfc').DataType      := ftString;
              Connection.zCommand.Params.ParamByName('Rfc').value         := sRfc;
              Connection.zCommand.Params.ParamByName('IdCompania').DataType  := ftString;
              Connection.zCommand.Params.ParamByName('IdCompania').value     := sIdCompania;
              Connection.zCommand.Params.ParamByName('LibretadeMar').DataType  := ftString;
              Connection.zCommand.Params.ParamByName('LibretadeMar').value     := sLibretadeMar;
              Connection.zCommand.Params.ParamByName('Vigencia').DataType  := ftDate;
              Connection.zCommand.Params.ParamByName('Vigencia').value     := dVigencia;
              connection.zCommand.ExecSQL;

              ProgressBar1.Position:=ProgressBar1.Position+1;
              Fila := Fila + 1;
              sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            end;
          except
            on E: Exception do
            begin
                MessageDlg(e.Message, mtInformation, [mbOk], 0)
            end;
          end;
        end;
        {$ENDREGION}

        {$REGION 'CONTROL PROGRAMADO'}
          if rbControlProgramado.Checked then
          begin
            ExcelApplication1.Visible[flcid] := False;
            if not chkBorrar.Checked then
            begin
              connection.QryBusca.Active:=false;
              connection.QryBusca.SQL.Text:='select sWbs from distribuciondeanexo where sContrato=:Contrato';
              connection.QryBusca.ParamByName('Contrato').AsString:= global_contrato;
              connection.QryBusca.Open;
              if connection.QryBusca.RecordCount>0 then
                if MessageDlg('Desea Eliminar los datos existentes?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                  chkBorrar.Checked := True;

            end;

            if chkBorrar.Checked then
            begin
              connection.QryBusca.Active:=false;
              connection.QryBusca.SQL.Text:='delete from distribuciondeanexo where sContrato=:Contrato';
              connection.QryBusca.ParamByName('Contrato').AsString:= global_contrato;
              connection.QryBusca.ExecSQL;
            end;


            importarControlProgramado;
          end;
        {$ENDREGION}

        {$REGION 'DETALLE DE ACTIVIDADES'}
          if rbDetalleDeActividades.Checked then begin
            ProgressBar1.Max := ExcelApplication1.Rows.CurrentRegion.Rows.Count;
            ProgressBar1.Position := 0;
            iFila := 2;
            while ExcelWorksheet1.Range['B' + Trim(IntToStr(iFila)), 'B' + Trim(IntToStr(iFila))].Text <> '' do begin
              Connection.zCommand.SQL.Text := 'START TRANSACTION;';
              Connection.zCommand.ExecSQL;

              Connection.QryBusca.SQL.Text := 'SELECT sWbs, mDescripcion FROM actividadesxorden WHERE sNumeroActividad = ' + QuotedStr(ExcelWorksheet1.Range['C' + Trim(IntToStr(iFila)), 'C' + Trim(IntToStr(iFila))].Text) + ' AND sNumeroOrden = ' + QuotedStr(ExcelWorksheet1.Range['B' + Trim(IntToStr(iFila)), 'B' + Trim(IntToStr(iFila))].Text) + ' AND sContrato = ' + QuotedStr(ExcelWorksheet1.Range['A' + Trim(IntToStr(iFila)), 'A' + Trim(IntToStr(iFila))].Text);
              Connection.QryBusca.Open;

              //MODIFICAR USO DE CONVENIOS - PENDIENTE
              if Connection.QryBusca.RecordCount > 0 then begin
                Connection.zCommand.SQL.Text := 'INSERT INTO actividadesxorden_detalle ' +
                                                '(sContrato, sIdConvenio, sNumeroOrden, sWbs, sNumeroActividad, dFechaInicio, sHoraInicio, dFechaFinal, sHoraFinal) VALUES ' +
                                                      //Contrato                                       //Convenio                      //Folio                                  //Wbs                                                       //Actividad                                //Fecha Inicio                     //Hora Inicio                               //FechaTermino                     //Hora Termino
                                                '('+QuotedStr(ExcelWorksheet1.Range['A' + Trim(IntToStr(iFila)), 'A' + Trim(IntToStr(iFila))].Text)+', '+QuotedStr(Global_Convenio)+', '+QuotedStr(ExcelWorksheet1.Range['B' + Trim(IntToStr(iFila)), 'B' + Trim(IntToStr(iFila))].Text)+', '+QuotedStr(Connection.QryBusca.FieldByName('sWbs').AsString)+', '+QuotedStr(ExcelWorksheet1.Range['C' + Trim(IntToStr(iFila)), 'C' + Trim(IntToStr(iFila))].Text)+', '+QuotedStr(ExcelWorksheet1.Range['D' + Trim(IntToStr(iFila)), 'D' + Trim(IntToStr(iFila))].Text)+', '+QuotedStr(ExcelWorksheet1.Range['E' + Trim(IntToStr(iFila)), 'E' + Trim(IntToStr(iFila))].Text)+', '+QuotedStr(ExcelWorksheet1.Range['F' + Trim(IntToStr(iFila)), 'F' + Trim(IntToStr(iFila))].Text)+', '+QuotedStr(ExcelWorksheet1.Range['G' + Trim(IntToStr(iFila)), 'G' + Trim(IntToStr(iFila))].Text)+'); ';
                Connection.zCommand.ExecSQL;
              end else begin
                ShowMessage('No se encontró la actividad ' + ExcelWorksheet1.Range['C' + Trim(IntToStr(iFila)), 'C' + Trim(IntToStr(iFila))].Text + ' en el folio ' + ExcelWorksheet1.Range['B' + Trim(IntToStr(iFila)), 'B' + Trim(IntToStr(iFila))].Text + ', no se guardarán los cambios.');
                Connection.zCommand.SQL.Text := 'ROLLBACK;';
                Connection.zCommand.ExecSQL;
                Exit;
              end;
              Inc(iFila);
            end;
            Connection.zCommand.SQL.Text := 'COMMIT;';
            Connection.zCommand.ExecSQL;
          end;
        {$ENDREGION}

  {$REGION 'ANEXO A'}
        //IMPORTACION DEL ANEXO A..
        if rAnexoA.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosA() then
            raise Exception.Create('Proceso Cancelado por el Sistema.');

          Fila := 2;
          ProgressBar1.Max := 0;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          if ValidaDeleteAnexosP('isometricos', 'sIsometricoReferencia', '', 'estimacionxpartida') then
            raise Exception.Create('Proceso Cancelado por el Sistema.');

            // Generar una lista de registros que deben existir
          Existir.Close;
          Existir.FieldDefs.Add('sContrato', ftString, 15);
          Existir.FieldDefs.Add('sIsometrico', ftString, 35);
          Existir.Open;
          Existir.EmptyTable;

          while (sValue <> '') do
          begin
            try
                // Generar la estructura de inserción
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('INSERT INTO isometricos ( sContrato, sIsometrico, iRevision, sIdGrupo, sIdPlataforma) ' +
                'VALUES (:contrato, :Isometrico, :Revision, :Grupo, :Plataforma )');
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

                // Verificar el contrato del registro obtenido desde excel
              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsIsometrico := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              ImpiRevision := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
              ImpsGrupo := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              ImpsPlataforma := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;

                // Inserto Datos a la Tabla .....
              CodErr1 := 'Importación de Plantilla de Anexo "A"';
              CodErr2 := 'Al Insetar registro';

              connection.zCommand.Active := False;
              Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
              Connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Isometrico').value := ImpsIsometrico;
              Connection.zCommand.Params.ParamByName('Revision').DataType := ftInteger;
              Connection.zCommand.Params.ParamByName('Revision').value := ImpiRevision;
              Connection.zCommand.Params.ParamByName('Grupo').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Grupo').value := ImpsGrupo;
              Connection.zCommand.Params.ParamByName('Plataforma').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Plataforma').value := ImpsPlataforma;
              connection.zCommand.ExecSQL;

            except
              on e: exception do
              begin
                  // Verificar si se encontró una duplicidad de registros
                if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                begin
                      // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro
                  if not SobreTodos then
                    Resp := MessageDlg('El isometrico identificado en EXCEL ya existe en la base de datos:' + #10 +
                      ImpsContrato + ' - ' + ImpsIsometrico + ' - ' + ImpiRevision + #10 + #10 +
                      '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                  if Resp = mrYesToAll then
                    SobreTodos := True;

                  if (Resp = mrYes) or SobreTodos then
                    Resp := mrYes;

                  if Resp = mrCancel then
                    raise Exception.Create('Proceso Cancelado por el Usuario.');

                  if Resp = mrYes then
                  begin
                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('UPDATE isometricos SET iRevision = :Revision, sIdGrupo = :Grupo, sIdPlataforma = :Plataforma ' +
                      'WHERE sContrato = :Contrato and sIsometrico = :Isometrico');

                    Connection.zCommand.ParamByName('contrato').AsString := ImpsContrato;
                    Connection.zCommand.ParamByName('Isometrico').AsString := ImpsIsometrico;
                    Connection.zCommand.ParamByName('Revision').AsString := ImpiRevision;
                    Connection.zCommand.ParamByName('Grupo').AsString := ImpsGrupo;
                    Connection.zCommand.ParamByName('Plataforma').AsString := ImpsPlataforma;
                    connection.zCommand.ExecSQL;
                  end;
                end
                else
                  raise;
              end;
            end;

              // Cargar la lista de registros procesados
            Existir.Append;
            Existir.FieldByName('sContrato').AsString := ImpsContrato;
            Existir.FieldByName('sIsometrico').AsString := ImpsIsometrico;
            Existir.Post;

            Fila := Fila + 1;
            ProgressBar1.Max := ProgressBar1.Max + 1;
            ProgressBar1.Position := ProgressBar1.Position + 1;
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          end;


            // Verificar los registros que debería ser eliminados
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('Select * from isometricos where sContrato = :Contrato');
          Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
          Connection.zCommand.Open;

          if Connection.zCommand.RecordCount > Existir.RecordCount then
          begin
            Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
              '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
            if Resp = mrCancel then
              raise Exception.Create('Proceso Cancelado por el Usuario.');

            if Resp = mrYes then
            begin
              connection.zCommand.First;
              while not connection.zCommand.Eof do
              begin
                if not Existir.Locate('sIsometrico', connection.zCommand.FieldByName('sIsometrico').AsString, []) then
                  Connection.zCommand.Delete;
                connection.zCommand.Next;
              end;
            end;
          end;
        end;

  {$ENDREGION}
  {$REGION 'PROGRAMA DE TRABAJO'}
        //IMPORTACION DEL PROGRAMA  DE TRABAJO ...
        //*******************************************************************************************
        if rPrograma.Checked then
        begin
            CodErr1 := '';
            CodErr2 := '';
            
            zq_Esp.Active:=False;
            zq_Esp.SQL.Clear;
            zq_Esp.SQL.Add('select sNumeroOrden, sIdFolio from ordenesdetrabajo where sContrato =:contrato and cIdStatus = "P" ');
            zq_Esp.ParamByName('Contrato').AsString:=global_contrato;
            zq_Esp.Open;

            if ValidaAnexosC('Programa') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            Fila := 2;
            sValue       := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
            ImpsConvenio := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;;
            ImpsNumeroOrden := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;

            BotonSelec := MessageDlg('Desea remplazar el programa de trabajo existente??', mtConfirmation, [mbYes, mbNo], 0);
            // Se elimina el catalogo de Anexo..
            if BotonSelec = mrYes then
            begin
              {Ahora llamamos la funcion que verifica si se puede eliinar el Anexo..}
              cadena := AntesEliminarFrente(sValue, sValue + '.%', ImpsNumeroOrden, 'Paquete', ImpsConvenio);
              if cadena <> '' then
              begin
                MessageDlg('No se puede Eliminar!. El Frente de Trabajo contine Partidas registradas en: ' + #13 + cadena, mtWarning, [mbOk], 0);
                exit;
              end
              else
                //Sino se encontraron datos se procede a eliminar..
                chkBorrar.Checked := True;
            end;

            if chkBorrar.Checked then
            begin
              iNivel := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              ImpsNumeroActividad := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;

              zq_Esp.Locate('sIdFolio', ImpsNumeroOrden, [loCaseInsensitive]);
              ImpsNumeroOrden := zq_Esp.FieldByName('sNumeroOrden').AsString;

              //Eliminamos las distribuciones,,
              DistribucionesFrente(ImpsNumeroOrden, ImpsNumeroActividad, 'Paquete', iNivel);

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('DELETE FROM actividadesxorden Where sContrato = :contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value    := Global_Contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value    := ImpsConvenio;
              Connection.zCommand.Params.ParamByName('Orden').DataType    := ftString;
              Connection.zCommand.Params.ParamByName('Orden').Value       := ImpsNumeroOrden;
              connection.zCommand.ExecSQL();

              //Eliminamos avances,
              EliminaAvances(ImpsNumeroOrden, ImpsConvenio);
            end;

            //Preguntamos al usuario dese subir las partidas de acuerdo a Excel o que inteligent las ordene,,
            Application.CreateForm(TFrmPopUpImportacionC, FrmPopUpImportacionC);
            FrmPopUpImportacionC.Left :=  trunc((Screen.Width) / 2) - trunc((FrmPopUpImportacionC.Width) / 2);
            FrmPopUpImportacionC.Top := trunc((screen.Height) / 2) - trunc((FrmPopUpImportacionC.Height) / 2);
            FrmPopUpImportacionC.Caption := 'Importando Programa de Trabajo';

            if FrmPopUpImportacionC.ShowModal = mrOk then
            begin
              lMsExcel := FrmPopUpImportacionC.chkExcel.Checked;
              lOrdenInteligent := FrmPopUpImportacionC.chkInteligent.Checked;
            end
            else
              raise Exception.Create('Proceso cancelado por el usuario.');

            ListaActividades := TStringList.Create;
            lObtenerDeAnexo  := False;

            lImportarDescripcion := false;
            lImportarPreciosMN   := false;
            lImportarPreciosDLL  := false;
            lImportarMedida      := false;
            lImportarFechaIni    := false;
            lImportarFechaFin    := false;
            lImportarCantidad    := false;

            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

             // Generar una lista de registros que deben existir
            Existir.Close;
            Existir.FieldDefs.Add('sContrato', ftString, 15);
            Existir.FieldDefs.Add('sIdConvenio', ftString, 5);
            Existir.FieldDefs.Add('sNumeroOrden', ftString, 35);
            Existir.FieldDefs.Add('sWbs', ftString, 100);
            Existir.FieldDefs.Add('sPaquete', ftString, 10);
            Existir.FieldDefs.Add('sNumeroActividad', ftString, 20);
            Existir.FieldDefs.Add('sTipoActividad', ftString, 15);
            Existir.Open;
            Existir.EmptyTable;

            t := 1;
            ProgressBar1.Max := 0;
            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato  := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

              ImpsConvenio    := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              if ImpsConvenio = trim('') then
                 ImpsConvenio := '1';
              ImpsNumeroOrden := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
              zq_Esp.Locate('sIdFolio', ImpsNumeroOrden, [loCaseInsensitive]);
              ImpsNumeroOrden := zq_Esp.FieldByName('sNumeroOrden').AsString;

              iNivel             := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              ImpsNumeroActividad:= ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
              ImpmDescripcion    := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
              ImpsMedida         := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
              ImpdCantidadAnexo  := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
              ImpdPonderado      := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;
              ImpdFechaInicio    := DateToStr(ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2);
              ImpdFechaFinal     := DateToStr(ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2);
              ImpsTipo           := ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;
              ImpsAnexo          := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;

              NuevaFechaInicial :=           ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value;
              NuevaFechaFinal   :=           ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value;

              if Trim(ImpsMedida) = '' then
                sTipo := 'Paquete'
              else
                sTipo := 'Actividad';

              sWbs := '';
              if iNivel <> 0 then
              begin
                for x := 1 to t - 1 do
                begin
                  if iNivel - 1 >= strToint(paquete[x][1]) then
                  begin
                    if (sTipo = 'Actividad') and (ImpsAnexo <> '') then
                      sWbs := paquete[x][2] + '.' + ImpsAnexo + '.'
                    else
                      sWbs := paquete[x][2] + '.';
                    ImpsWbsAnterior := paquete[x][2];
                  end;
                end;

                {Obtenemos la Wbs Anterior si selccionamos la Opcion ordenar x inteligent..}
                if lOrdenInteligent then
                begin
                  connection.QryBusca.Active := False;
                  connection.QryBusca.SQL.Clear;
                  connection.QryBusca.SQL.Add('select iItemOrden from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and iNivel =:Nivel and sWbs =:WbsAnt order by iItemOrden ');
                  connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
                  connection.QryBusca.Params.ParamByName('Contrato').value    := global_contrato;
                  connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
                  connection.QryBusca.Params.ParamByName('Convenio').value    := global_convenio;
                  connection.QryBusca.Params.ParamByName('WbsAnt').DataType   := ftString;
                  if (sTipo = 'Actividad') and (ImpsAnexo <> '') then
                    connection.QryBusca.Params.ParamByName('WbsAnt').value := copy(sWbs, 1, length(sWbs) - (length(ImpsAnexo) + 2))
                  else
                    connection.QryBusca.Params.ParamByName('WbsAnt').value := copy(sWbs, 1, length(sWbs) - 1);
                  connection.QryBusca.Params.ParamByName('Nivel').DataType :=  ftInteger;
                  connection.QryBusca.Params.ParamByName('Nivel').value    := iNivel - 1;
                  connection.QryBusca.Open;

                  if connection.QryBusca.RecordCount > 0 then
                    sItemOrdenAnterior := connection.QryBusca.FieldValues['iItemOrden']
                  else
                    sItemOrdenAnterior := '';
                end;

                sWbs := sWbs + ImpsNumeroActividad;
                if lMsExcel then
                  ImpiItemOrden := sFnInsertaItem(ImpsConvenio, ImpsNumeroActividad, ImpsWbsAnterior, '', sTipo, 'actividadesxorden', ImpsNumeroOrden, '', iNivel);

                if lOrdenInteligent then
                  ImpiItemOrden := sItemOrdenAnterior + sFnBuscaItem(ImpsConvenio,ImpsNumeroActividad, ImpsWbsAnterior, sItemOrdenAnterior, sTipo, '', 'actividadesxorden', iNivel);
              end
              else
              begin
                if lMsExcel then
                  ImpiItemOrden := sFnInsertaItem(ImpsConvenio, ImpsNumeroActividad,
                    ImpsWbsAnterior, '', sTipo, 'actividadesxorden', ImpsNumeroOrden, '', iNivel);
                if lOrdenInteligent then
                  ImpiItemOrden := sFnBuscaItem(ImpsConvenio,ImpsNumeroActividad, ImpsWbsAnterior, sItemOrdenAnterior, sTipo, '', 'actividadesxorden', iNivel);

                ImpsWbsAnterior := '';
                sWbs := ImpsNumeroActividad;
              end;

              if sTipo = 'Paquete' then
              begin
                paquete[t][1] := inttostr(iNivel);
                paquete[t][2] :=             sWbs;
                paquete[t][3] :=    ImpiItemOrden;
                t := t + 1;
              end;

              {Verificamos si es una actividad o un paquete, para traer WbsContrato..}
              if sTipo = 'Actividad' then
              begin
                Connection.qryBusca.Active := False;
                Connection.qryBusca.SQL.Clear;
                Connection.qryBusca.SQL.Add('Select sWbs, sActividadAnterior, mDescripcion, dVentaMN, dVentaDLL, sMedida, dCantidadAnexo, dFechaInicio, dFechaFinal, sAnexo, sTipoAnexo ' +
                  'From actividadesxanexo Where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad ' +
                  'and sTipoActividad = "Actividad" and sAnexo=:anexo ');
                Connection.qryBusca.Params.ParamByName('Contrato').DataType  :=            ftString;
                Connection.qryBusca.Params.ParamByName('Contrato').Value     :=     global_contrato;
                Connection.qryBusca.Params.ParamByName('Convenio').DataType  :=            ftString;
                Connection.qryBusca.Params.ParamByName('Convenio').Value     :=     global_convenio;
                Connection.qryBusca.Params.ParamByName('Actividad').DataType :=            ftString;
                Connection.qryBusca.Params.ParamByName('Actividad').Value    := ImpsNumeroActividad;
                Connection.qryBusca.Params.ParamByName('anexo').DataType     :=            ftString;
                Connection.qryBusca.Params.ParamByName('anexo').Value        :=           ImpsAnexo;
                Connection.qryBusca.Open;

                if Connection.qryBusca.RecordCount > 0 then
                begin
                  ImpsWbsContrato := Connection.qryBusca.FieldValues['sWbs'];

                  if lObtenerDeAnexo then
                  begin
                    ImpsActAnterior := Connection.qryBusca.FieldValues['sActividadAnterior'];

                    if lImportarDescripcion then
                      ImpmDescripcion :=     Connection.qryBusca.FieldValues['mDescripcion'];

                    if lImportarFechaIni then
                      ImpdFechaInicio :=     Connection.qryBusca.FieldValues['dFechaInicio'];

                    if lImportarFechaFin then
                      ImpdFechaFinal :=       Connection.qryBusca.FieldValues['dFechaFinal'];

                    ImpsAnexo :=     Connection.qryBusca.FieldValues['sAnexo'];
                    ImpsTipo  := Connection.qryBusca.FieldValues['sTipoAnexo'];

                    if lImportarPreciosMN then
                      ImpdVentaMN  :=  Connection.qryBusca.FieldValues['dVentaMN'];

                    if lImportarPreciosDLL then
                      ImpdVentaDLL := Connection.qryBusca.FieldValues['dVentaDLL'];

                    if lImportarMedida then
                      ImpsMedida   :=   Connection.qryBusca.FieldValues['sMedida'];

                    if lImportarCantidad then
                      ImpdCantidadAnexo := Connection.qryBusca.FieldValues['dCantidadAnexo'];
                  end;
                end; // Termina if Connection.qryBusca.RecordCount > 0 ....
              end
              else
              begin
                      {Para el caso de la integirdad de datos.. tomamos la sWbs del Paquete principal..}
                Connection.qryBusca.Active := False;
                Connection.qryBusca.SQL.Clear;
                Connection.qryBusca.SQL.Add('Select sWbs From actividadesxanexo Where sContrato = :Contrato And sIdConvenio =:Convenio and sTipoActividad = "Paquete" and iNivel = 0');
                Connection.qryBusca.Params.ParamByName('Contrato').DataType :=        ftString;
                Connection.qryBusca.Params.ParamByName('Contrato').Value    := global_contrato;
                Connection.qryBusca.Params.ParamByName('Convenio').DataType :=        ftString;
                Connection.qryBusca.Params.ParamByName('Convenio').Value    := global_convenio;
                Connection.qryBusca.Open;

                if connection.QryBusca.RecordCount > 0 then
                  ImpsWbsContrato := Connection.qryBusca.FieldValues['sWbs']
                else
                  ImpsWbsContrato := '';

              end; // Termina if ObtenerDeAnexo ....

              if sTipo = 'Paquete' then
              begin
                ImpdCantidadAnexo := '1.00';
                ImpsMedida        :=     '';
                ImpdVentaMN       := '0.00';
                ImpdVentaDLL      := '0.00';
              end;

              try
                      // Inserto Datos a la Tabla .....
                CodErr1 := 'Al importar información del programa de trabajo desde EXCEL';
                CodErr2 :=                  'Al insertar registros de actividadesxorden';

                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                sSQL := 'INSERT INTO actividadesxorden ( sContrato , sNumeroOrden, sIdConvenio, sTipoActividad, sWbsAnterior, ' +
                                    'sWbs, sNumeroActividad, iItemOrden , mDescripcion, dFechaInicio, dDuracion, dFechaFinal, ' +
                              'dVentaMN, dVentaDLL, sMedida, dCantidad, dPonderado, iColor, lGerencial, iNivel, mComentarios, ' +
                                                     'sTipoAnexo, sWbsContrato, sAnexo, sActividadAnterior, lExtraordinario ) ' +
                              'VALUES (:contrato, :orden, :convenio, :tipo, :anterior, :wbs, :actividad, :Item, :Descripcion, ' +
                             ':Inicio, :Duracion, :Final, :MN, :DLL, :Medida, :CantidadAnexo, :Ponderado, :color, :Gerencial, ' +
                                      ':Nivel, :Comentarios, :TipoA, :WbsContrato, :Anexo, :ActividadAnterior, :Extraordinario)';
                connection.zCommand.SQL.Add(sSQL);
                Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('contrato').value    := ImpsContrato;
                Connection.zCommand.Params.ParamByName('orden').DataType    := ftString;
                Connection.zCommand.Params.ParamByName('orden').value       := ImpsNumeroOrden;
                Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                Connection.zCommand.Params.ParamByName('convenio').value    := ImpsConvenio;
                Connection.zCommand.Params.ParamByName('tipo').DataType     := ftString;
                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('tipo').value           := sTipo
                else
                  Connection.zCommand.Params.ParamByName('tipo').value           := 'Actividad';
                Connection.zCommand.Params.ParamByName('anterior').DataType      := ftString;
                Connection.zCommand.Params.ParamByName('anterior').value         := Trim(ImpsWbsAnterior);
                Connection.zCommand.Params.ParamByName('wbs').DataType           := ftString;
                if Trim(ImpsWbsAnterior) <> '' then
                  Connection.zCommand.Params.ParamByName('wbs').value := sWbs
                else
                  Connection.zCommand.Params.ParamByName('wbs').value            := Trim(ImpsNumeroActividad);
                Connection.zCommand.Params.ParamByName('WbsContrato').DataType   := ftString;
                Connection.zCommand.Params.ParamByName('WbsContrato').value      := ImpsWbsContrato;
                Connection.zCommand.Params.ParamByName('actividad').DataType     := ftString;
                Connection.zCommand.Params.ParamByName('actividad').value        := Trim(ImpsNumeroActividad);
                Connection.zCommand.Params.ParamByName('Item').DataType          := ftString;
                Connection.zCommand.Params.ParamByName('Item').value             := ImpiItemOrden;
                Connection.zCommand.Params.ParamByName('Descripcion').DataType   := ftMemo;
                Connection.zCommand.Params.ParamByName('Descripcion').value      := Trim(ImpmDescripcion);
                Connection.zCommand.Params.ParamByName('Inicio').AsDateTime      := NuevaFechaInicial;
                Connection.zCommand.Params.ParamByName('Duracion').value         := (StrToDate(ImpdFechaFinal) - StrToDate(ImpdFechaInicio)) + 1; // DaysBetween(StrToDate(ImpdFechaInicio),StrToDate(ImpdFechaFinal) )+1;
                Connection.zCommand.Params.ParamByName('Final').AsDateTime       := NuevaFechaFinal;
                Connection.zCommand.Params.ParamByName('MN').DataType            := ftFloat;
                Connection.zCommand.Params.ParamByName('MN').value               := ImpdVentaMN;
                Connection.zCommand.Params.ParamByName('DLL').DataType           := ftFloat;
                Connection.zCommand.Params.ParamByName('DLL').value              := ImpdVentaDLL;
                Connection.zCommand.Params.ParamByName('Medida').DataType        := ftString;
                Connection.zCommand.Params.ParamByName('Medida').value           := Trim(ImpsMedida);
                Connection.zCommand.Params.ParamByName('CantidadAnexo').DataType := ftFloat;
                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('CantidadAnexo').value  := 1
                else
                  Connection.zCommand.Params.ParamByName('CantidadAnexo').value  := ImpdCantidadAnexo;
                Connection.zCommand.Params.ParamByName('Ponderado').DataType     := ftFloat;
                Connection.zCommand.Params.ParamByName('Ponderado').value        := ImpdPonderado;
                Connection.zCommand.Params.ParamByName('Color').DataType         := ftInteger;

                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('Color').value          := 12
                else
                  Connection.zCommand.Params.ParamByName('Color').value          := 0;
                Connection.zCommand.Params.ParamByName('Gerencial').DataType     := ftString;

                if sTipo = 'Paquete' then
                  Connection.zCommand.Params.ParamByName('Gerencial').value      := 'Si'
                else
                  Connection.zCommand.Params.ParamByName('Gerencial').value      := 'No';

                Connection.zCommand.Params.ParamByName('Nivel').DataType             :=       ftInteger;
                Connection.zCommand.Params.ParamByName('Nivel').value                :=          iNivel;
                Connection.zCommand.Params.ParamByName('Comentarios').DataType       :=        ftString;
                Connection.zCommand.Params.ParamByName('Comentarios').value          :=             '*';
                Connection.zCommand.Params.ParamByName('TipoA').DataType             :=        ftString;
                Connection.zCommand.Params.ParamByName('TipoA').value                :=        ImpsTipo;
                Connection.zCommand.Params.ParamByName('Anexo').DataType             :=        ftString;
                Connection.zCommand.Params.ParamByName('Anexo').value                :=       ImpsAnexo;
                Connection.zCommand.Params.ParamByName('ActividadAnterior').DataType :=        ftString;
                Connection.zCommand.Params.ParamByName('ActividadAnterior').value    := ImpsActAnterior;
                Connection.zCommand.Params.ParamByName('Extraordinario').DataType    :=        ftString;
                Connection.zCommand.Params.ParamByName('Extraordinario').value       :=            'No';
                connection.zCommand.ExecSQL;
              except
                on e: exception do
                begin
                end;
              end;

              // Cargar la lista de registros procesados
              Existir.Append;
              Existir.FieldByName('sContrato').AsString    := global_contrato;
              Existir.FieldByName('sIdConvenio').AsString  := global_convenio;
              Existir.FieldByName('sNumeroOrden').AsString := ImpsNumeroOrden;
              Existir.FieldByName('sWbs').AsString         :=            sWbs;
              Existir.FieldByName('sPaquete').AsString     :=             '0';
              Existir.FieldByName('sNumeroActividad').AsString := ImpsNumeroActividad;
              Existir.FieldByName('sTipoActividad').AsString :=         sTipo;
              Existir.Post;

              {Se crea lista con actividades...}
              if ListaActividades.IndexOf(Trim(ImpsWbsContrato)) < 0 then
                ListaActividades.Add(Trim(ImpsWbsContrato));

              //Agregar las unidades de medida a la configuracion en automatico...
              CodErr1 :=  'Al registrar datos de programa de trabajo';
              CodErr2 := 'Al actualizar información de configuracion';
              x := pos(ImpsMedida, ValidaMat);
              if (x < 1) and (trim(ImpsMedida) <> '') then
              begin
                ValidaMat := ValidaMat + ImpsMedida + '|';
                Connection.zCommand.Active := False;
                Connection.zCommand.SQL.Clear;
                Connection.zCommand.SQL.Add('Update configuracion set txtValidaMaterial = :Medidas where sContrato = :Contrato');
                Connection.zCommand.Params.ParamByName('Contrato').DataType :=        ftString;
                Connection.zCommand.Params.ParamByName('Contrato').Value    := global_contrato;
                Connection.zCommand.Params.ParamByName('Medidas').DataType  :=        ftString;
                Connection.zCommand.Params.ParamByName('Medidas').Value     :=       ValidaMat;
                Connection.zCommand.ExecSQL;
              end;

              x := pos(ImpsMedida, MaterialAuto);
              if (x < 1) and (trim(ImpsMedida) <> '') then
              begin
                MaterialAuto := MaterialAuto + ImpsMedida + '|';
                Connection.zCommand.Active := False;
                Connection.zCommand.SQL.Clear;
                Connection.zCommand.SQL.Add('Update configuracion set txtMaterialAutomatico = :Medidas where sContrato = :Contrato');
                Connection.zCommand.Params.ParamByName('Contrato').DataType :=        ftString;
                Connection.zCommand.Params.ParamByName('Contrato').Value    := global_contrato;
                Connection.zCommand.Params.ParamByName('Medidas').DataType  :=        ftString;
                Connection.zCommand.Params.ParamByName('Medidas').Value     :=    MaterialAuto;
                Connection.zCommand.ExecSQL;
              end;

              ImpdVentaMN := '0';
              ImpdVentaDLL := '0';

              ProgressBar1.Max      := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              fila := fila + 1;
              sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);
            end; // Termino del While

            //Aquí insertamos el convenio 1 en automático en repogramaciones del contrato
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('insert into convenios (sContrato, sNumeroOrden, sIdConvenio, sDescripcion, dFecha, dFechaInicio, dFechaFinal) '+
                                        'values (:contrato, :orden, :convenio, :descripcion, :Fecha, :fechaI, :fechaF)');
            Connection.zCommand.Params.ParamByName('Contrato').AsString  := Global_Contrato;
            Connection.zCommand.Params.ParamByName('Convenio').AsString  := ImpsConvenio;
            Connection.zCommand.Params.ParamByName('Orden').AsString     := ImpsNumeroOrden;
            Connection.zCommand.Params.ParamByName('descripcion').AsString := ImpsNumeroOrden +' C-'+ ImpsConvenio;
            Connection.zCommand.Params.ParamByName('fecha').AsDate         := Now();
            Connection.zCommand.Params.ParamByName('fechaI').AsDate        := StrToDate(ImpdFechaInicio);
            Connection.zCommand.Params.ParamByName('fechaF').AsDate        := StrToDate(ImpdFechaFinal);
            connection.zCommand.ExecSQL();

            // Verificar los registros que debería ser eliminados
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Select * from actividadesxorden where sContrato = :Contrato and sIdConvenio = :Convenio');
            Connection.zCommand.ParamByName('Contrato').AsString :=    ImpsContrato;
            Connection.zCommand.ParamByName('Convenio').AsString := ImpsConvenio;
            Connection.zCommand.Open;

            if Connection.zCommand.RecordCount > Existir.RecordCount then
            begin
                //Actualizamos las fechas de inicio y fin de ordenesdetrabajo.
                connection.QryBusca.Active := False;
                connection.QryBusca.SQL.Clear;
                connection.QryBusca.SQL.Add('update ordenesdetrabajo set dFiProgramado =:FechaI, dFfProgramado =:FechaF where sContrato =:Contrato and sNumeroOrden =:Orden');
                Connection.QryBusca.ParamByName('Contrato').AsString := ImpsContrato;
                Connection.QryBusca.ParamByName('Orden').AsString    := ImpsNumeroOrden;
                Connection.QryBusca.ParamByName('FechaI').AsDate     := Connection.zCommand.FieldByName('dFechaInicio').AsDateTime;
                Connection.QryBusca.ParamByName('FechaF').AsDate     := Connection.zCommand.FieldByName('dFechaFinal').AsDateTime;
                Connection.QryBusca.ExecSQL;

//                Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
//                  '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
//                if Resp = mrCancel then
//                  raise Exception.Create('Proceso Cancelado por el Usuario.');
//  
//                if Resp = mrYes then
//                begin
//                  connection.zCommand.First;
//                  while not connection.zCommand.Eof do
//                  begin
//                    if not Existir.Locate('sNumeroOrden;sWbs;sPaquete;sNumeroActividad;sTipoActividad', VarArrayOf([connection.zCommand.FieldByName('sNumeroOrden').AsString, connection.zCommand.FieldByName('sWbs').AsString, connection.zCommand.FieldByName('sPaquete').AsString, connection.zCommand.FieldByName('sNumeroActividad').AsString, connection.zCommand.FieldByName('sTipoActividad').AsString]), []) then
//                      Connection.zCommand.Delete;
//                    connection.zCommand.Next;
//                  end;
//                end;
            end;

            if BotonSelec = mrYes then
            begin
              try
                Kardex('Importacion de Datos', 'Termina Proceso', 'Frente de Trabajo', '', '', '', '','Tarifa Diaria','Importacion de Datos');
              except
                on e: exception do
                begin
                  // Aquí si se debe dejar independiente esta excepción debido a que si no se puede registrar el kardex tampoco se quiere que se cancele todo el proceso.
                  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al registrar en kardex Importacion de Frente de Trabajo', 0);
                end;
              end;
            end;

        end;
  {$ENDREGION}
  {$REGION 'ANEXO C'}
        //Aqui se importa el  Anexo "C" ...
        //*******************************************************************************************
        if rAnexoC.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosC('AnexoC') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          BotonSelec := MessageDlg('¿Desea remplazar el catalogo de partidas anexo existente?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
          if BotonSelec = mrCancel then
            raise Exception.Create('Proceso Cancelado por el usuario');

          Fila := 2;
          sValue := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;

          // Se elimina el catalogo de Anexo..
          if BotonSelec = mrYes then
          begin
             {Ahora llamamos la funcion que verifica si se puede eliinar el Anexo..}
            cadena := AntesEliminarAnexo(sValue, sValue + '.%', 'Paquete');
            if cadena <> '' then
            begin
              MessageDlg('No se puede Eliminar!. El Anexo C contine Partidas registradas en: ' + #13 + cadena, mtWarning, [mbOk], 0);
              exit;
            end
            else
                //Sino se encontraron datos se procede a eliminar..
              chkBorrar.Checked := True;
          end;

          if chkBorrar.Checked then
          begin
            CodErr1 := 'Al importar anexo "C"';
            CodErr2 := 'Al tratar de eliminar registros de actividadesxanexo existentes';

            iNivel := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            ImpsNumeroActividad := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;

            connection.zCommand.Active := false;
            connection.zCommand.SQL.Text := 'SET FOREIGN_KEY_CHECKS=0';
            connection.zCommand.ExecSQL;
            //Eliminamos las distribuciones,,
            DistribucionesAnexo(ImpsNumeroActividad, 'Paquete', iNivel);

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('DELETE FROM actividadesxanexo Where sContrato = :contrato And sIdConvenio = :Convenio');
            Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
            Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
            Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
            connection.zCommand.ExecSQL();

            //Eliminamos avances,
            EliminaAvances('', '');
          end;

          CodErr1 := '';
          CodErr2 := '';

          //Preguntamos al usuario dese subir las partidas de acuerdo a Excel o que inteligent las ordene,,
          Application.CreateForm(TFrmPopUpImportacionC, FrmPopUpImportacionC);
          FrmPopUpImportacionC.Left := trunc((Screen.Width) / 2) - trunc((FrmPopUpImportacionC.Width) / 2);
          FrmPopUpImportacionC.Top := trunc((screen.Height) / 2) - trunc((FrmPopUpImportacionC.Height) / 2);

          if FrmPopUpImportacionC.ShowModal = mrOk then
          begin
            lMsExcel := FrmPopUpImportacionC.chkExcel.Checked;
            lOrdenInteligent := FrmPopUpImportacionC.chkInteligent.Checked;
          end
          else
          begin
            FrmPopUpImportacion.Free;
            exit;
          end;

          I := 0;
          t := 1;
          //Procedemos a leer el archivo de Excel..
          sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
          while (sValue <> '') do
          begin
            CodErr1 := '';
            CodErr2 := '';

            ProgressBar1.Position := Fila - 1;
            if lContratoActual then
              ImpsContrato := global_contrato
            else
              ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

            inc(I);
            iNivel := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            ImpsNumeroActividad := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
            ImpsEspecificacion := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
            ImpmDescripcion := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
            ImpsMedida := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
            ImpdCantidadAnexo := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
            ImpdPonderado := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
            ImpdVentaMN := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;
            ImpdVentaDLL := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
            ImpsFase := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;
            ImpdFechaInicio := ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;
            ImpdFechaFinal := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;
            ImpsAnexo := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;
            ImpsTipo := ExcelWorksheet1.Range['O' + Trim(IntToStr(Fila)), 'O' + Trim(IntToStr(Fila))].Value2;
            ImpsExtraordinario := ExcelWorksheet1.Range['P' + Trim(IntToStr(Fila)), 'P' + Trim(IntToStr(Fila))].Value2;
            ImpdFechaInicio := DateToStr(ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2);
            ImpdFechaFinal := DateToStr(ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2);

            if Trim(ImpsMedida) = '' then
              sTipo := 'Paquete'
            else
              sTipo := 'Actividad';

            sWbs := '';
            if iNivel <> 0 then
            begin
              for x := 1 to t - 1 do
              begin
                if iNivel - 1 >= strToint(paquete[x][1]) then
                begin
                  if (sTipo = 'Actividad') and (ImpsAnexo <> '') then
                    sWbs := paquete[x][2] + '.' + ImpsAnexo + '.'
                  else
                    sWbs := paquete[x][2] + '.';
                  ImpsWbsAnterior := paquete[x][2];
                end;
              end;

              {Obtenemos la Wbs Anterior si selccionamos la Opcion ordenar x inteligent..}
              if lOrdenInteligent then
              begin
                connection.QryBusca.Active := False;
                connection.QryBusca.SQL.Clear;
                connection.QryBusca.SQL.Add('select iItemOrden from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and iNivel =:Nivel and sWbs =:WbsAnt order by iItemOrden ');
                connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
                connection.QryBusca.Params.ParamByName('Contrato').value := global_contrato;
                connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
                connection.QryBusca.Params.ParamByName('Convenio').value := global_convenio;
                connection.QryBusca.Params.ParamByName('WbsAnt').DataType := ftString;
                if (sTipo = 'Actividad') and (ImpsAnexo <> '') then
                  connection.QryBusca.Params.ParamByName('WbsAnt').value := copy(sWbs, 1, length(sWbs) - (length(ImpsAnexo) + 2))
                else
                  connection.QryBusca.Params.ParamByName('WbsAnt').value := copy(sWbs, 1, length(sWbs) - 1);
                connection.QryBusca.Params.ParamByName('Nivel').DataType := ftInteger;
                connection.QryBusca.Params.ParamByName('Nivel').value := iNivel - 1;
                connection.QryBusca.Open;

                if connection.QryBusca.RecordCount > 0 then
                  sItemOrdenAnterior := connection.QryBusca.FieldValues['iItemOrden']
                else
                  sItemOrdenAnterior := '';
              end;

              sWbs := sWbs + ImpsNumeroActividad;
              if lMsExcel then
                ImpiItemOrden := sFnInsertaItem(ImpsConvenio,ImpsNumeroActividad, ImpsWbsAnterior, '', sTipo, 'actividadesxanexo', '', ImpsAnexo, iNivel);

              if lOrdenInteligent then
                ImpiItemOrden := sItemOrdenAnterior + sFnBuscaItem(ImpsConvenio,ImpsNumeroActividad, ImpsWbsAnterior, sItemOrdenAnterior, sTipo, '', 'actividadesxanexo', iNivel);

            end
            else
            begin
              if lMsExcel then
                ImpiItemOrden := sFnInsertaItem(ImpsConvenio,ImpsNumeroActividad, ImpsWbsAnterior, '', sTipo, 'actividadesxanexo', '', ImpsAnexo, iNivel);

              if lOrdenInteligent then
                ImpiItemOrden := sFnBuscaItem(ImpsConvenio,ImpsNumeroActividad, ImpsWbsAnterior, sItemOrdenAnterior, sTipo, '', 'actividadesxanexo', iNivel);

              ImpsWbsAnterior := '';
              sWbs := ImpsNumeroActividad;
            end;

            if sTipo = 'Paquete' then
            begin
              paquete[t][1] := inttostr(iNivel);
              paquete[t][2] := sWbs;
              paquete[t][3] := ImpiItemOrden;
              t := t + 1;
            end;

            //Checo si existe la partida o paquete en el contrato ....
            Connection.qryBusca.Active := False;
            Connection.qryBusca.SQL.Clear;
            Connection.qryBusca.SQL.Add('Select sWbsAnterior, dCantidadAnexo From actividadesxanexo Where sContrato = :Contrato And sIdConvenio = :Convenio ' +
              'And sNumeroActividad = :Actividad and sWbs = :Wbs');
            Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
            Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
            Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
            Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
            Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
            Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
            Connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
            Connection.qryBusca.Params.ParamByName('Wbs').Value := sWbs;
            Connection.qryBusca.Open;

            if (Connection.QryBusca.RecordCount > 0) then
            begin
              ImpsWbsAnterior := connection.QryBusca.FieldValues['sWbsAnterior'];
              lActualiza := True
            end
            else
              lActualiza := False;

            //Se actualiza informacion en caso de que y exista la partida en el mismo paquete...
            CodErr1 := 'Al importar anexo "C"';
            CodErr2 := 'Al registrar actividadesxanexo';
            if lActualiza then
            begin
              if not SobreTodos then
                Resp := MessageDlg('La partida de orden de trabajo identificada en EXCEL ya existe en la base de datos:' + #10 +
                  ImpsContrato + ' - ' + ImpsNumeroOrden + ' - ' + Global_Convenio + ' - ' + sWbs + ' Partida -> ' + ImpsNumeroActividad + ' Fila [' + IntToStr(Fila) + '] ' + #10 + #10 +
                  '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

              if Resp = mrYesToAll then
                SobreTodos := True;

              if (Resp = mrYes) or SobreTodos then
                Resp := mrYes;

              if Resp = mrCancel then
                raise Exception.Create('Proceso Cancelado por el Usuario.');

              if Resp = mrYes then
              begin
                CodErr1 := 'Al importar anexo "C"';
                CodErr2 := 'Al tratar de actualizar registros en la tabla actividadesxanexo';

                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('UPDATE actividadesxanexo SET sWbsAnterior = :WbsAnterior, dCantidadAnexo = :Cantidad, dFechaInicio = :Inicio, dFechaFinal = :Final, dVentaMN = :VentaMN, dVentaDLL = :VentaDLL ' +
                  'Where sContrato = :contrato and sIdConvenio = :convenio and sNumeroActividad = :Actividad And sTipoActividad = "Actividad" ');
                Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                Connection.zCommand.Params.ParamByName('WbsAnterior').DataType := ftString;
                Connection.zCommand.Params.ParamByName('WbsAnterior').value := Trim(ImpsWbsAnterior);
                Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                Connection.zCommand.Params.ParamByName('actividad').value := Trim(ImpsNumeroActividad);
                Connection.zCommand.Params.ParamByName('Inicio').DataType := ftDate;
                Connection.zCommand.Params.ParamByName('Inicio').value := StrToDate(ImpdFechaInicio);
                Connection.zCommand.Params.ParamByName('Final').DataType := ftDate;
                Connection.zCommand.Params.ParamByName('Final').value := StrToDate(ImpdFechaFinal);
                Connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('Cantidad').value := ImpdCantidadAnexo + connection.QryBusca.FieldValues['dCantidadAnexo'];
                Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaMN').value := ImpdVentaMN;
                Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaDLL').value := ImpdVentaDLL;
                connection.zCommand.ExecSQL;
              end;
            end
            else
            begin
              CodErr1 := 'Al importar anexo "C"';
              CodErr2 := 'Al tratar de insertar registros en la tabla actividadesxanexo';

              connection.zCommand.Active := False;
              Connection.zCommand.SQL.Clear;
              Connection.zCommand.SQL.Add('INSERT INTO actividadesxanexo ( sContrato , sIdConvenio, iNivel,sTipoActividad, sWbsAnterior, sWbs, sNumeroActividad, sActividadAnterior, sEspecificacion, iItemOrden , mDescripcion, dFechaInicio, dDuracion, dFechaFinal, ' +
                'dVentaMN, dVentaDLL, sMedida, dCantidadAnexo, dPonderado, sAnexo, iColor,sSimbolo, sTipoAnexo, sIdFase, lExtraordinario ) ' +
                'VALUES (:contrato, :convenio, :nivel ,:tipo, :anterior, :wbs, :actividad, :actividad, :Especifica, :Item, :Descripcion, :Inicio, :Duracion, :Final, :MN, :DLL, :Medida, :CantidadAnexo, :Ponderado, :Anexo, :Color ,"+", :TipoA, :Fase, :Extraordinario)');
              Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
              Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
              Connection.zCommand.Params.ParamByName('tipo').DataType := ftString;

              if Trim(ImpsMedida) = '' then
                Connection.zCommand.Params.ParamByName('tipo').value := 'Paquete'
              else
                Connection.zCommand.Params.ParamByName('tipo').value := 'Actividad';
              Connection.zCommand.Params.ParamByName('anterior').DataType := ftString;
              Connection.zCommand.Params.ParamByName('anterior').value := Trim(ImpsWbsAnterior);
              Connection.zCommand.Params.ParamByName('wbs').DataType := ftString;

              if Trim(ImpsWbsAnterior) <> '' then
                Connection.zCommand.Params.ParamByName('wbs').value := sWbs
              else
                Connection.zCommand.Params.ParamByName('wbs').value := Trim(ImpsNumeroActividad);

              Connection.zCommand.Params.ParamByName('nivel').DataType := ftInteger;
              Connection.zCommand.Params.ParamByName('nivel').value := iNivel;

              Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;

              Connection.zCommand.Params.ParamByName('actividad').value := Trim(ImpsNumeroActividad);
              Connection.zCommand.Params.ParamByName('Especifica').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Especifica').value := Trim(ImpsEspecificacion);
              Connection.zCommand.Params.ParamByName('Item').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Item').value := ImpiItemOrden;
              Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
              Connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
              Connection.zCommand.Params.ParamByName('Inicio').DataType := ftDate;
              Connection.zCommand.Params.ParamByName('Inicio').value := StrToDate(ImpdFechaInicio);
              Connection.zCommand.Params.ParamByName('Duracion').DataType := ftInteger;
              Connection.zCommand.Params.ParamByName('Duracion').value := (StrToDate(ImpdFechaFinal) - StrToDate(ImpdFechaInicio)) + 1; //DaysBetween(StrToDate(ImpdFechaInicio),StrToDate(ImpdFechaFinal)) + 1;
              Connection.zCommand.Params.ParamByName('Final').DataType := ftDate;
              Connection.zCommand.Params.ParamByName('Final').value := StrToDate(ImpdFechaFinal);
              Connection.zCommand.Params.ParamByName('MN').DataType := ftFloat;
              Connection.zCommand.Params.ParamByName('MN').value := ImpdVentaMN;
              Connection.zCommand.Params.ParamByName('DLL').DataType := ftFloat;
              Connection.zCommand.Params.ParamByName('DLL').value := ImpdVentaDLL;
              Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Medida').value := Trim(ImpsMedida);
              Connection.zCommand.Params.ParamByName('CantidadAnexo').DataType := ftFloat;
              if Trim(ImpsMedida) = '' then
                Connection.zCommand.Params.ParamByName('CantidadAnexo').value := 1
              else
                Connection.zCommand.Params.ParamByName('CantidadAnexo').value := ImpdCantidadAnexo;
              Connection.zCommand.Params.ParamByName('Ponderado').DataType := ftFloat;
              Connection.zCommand.Params.ParamByName('Ponderado').value := ImpdPonderado;
              Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
              Connection.zCommand.Params.ParamByName('Extraordinario').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Extraordinario').value := ImpsExtraordinario;
              Connection.zCommand.Params.ParamByName('TipoA').DataType := ftString;
              Connection.zCommand.Params.ParamByName('TipoA').value    := ImpsTipo;
              Connection.zCommand.Params.ParamByName('Fase').DataType := ftString;
              if sTipo = 'Paquete' then
                Connection.zCommand.Params.ParamByName('Fase').value  := ''
              else
                Connection.zCommand.Params.ParamByName('Fase').value := ImpsFase;
              Connection.zCommand.Params.ParamByName('Color').DataType := ftInteger;

              if Trim(ImpsMedida) = '' then
                Connection.zCommand.Params.ParamByName('Color').value := 12
              else
                Connection.zCommand.Params.ParamByName('Color').value := 0;
              connection.zCommand.ExecSQL;

              //Agregar las unidades de medida a la configuracion en automatico...
              CodErr1 := 'Al importar anexo "C"';
              CodErr2 := 'Al tratar de actualizar unidades de medida en tabla de configuración';

              x := pos(ImpsMedida, ValidaMat);
              if (x < 1) and (trim(ImpsMedida) <> '') then
              begin
                ValidaMat := ValidaMat + ImpsMedida + '|';
                Connection.zCommand.Active := False;
                Connection.zCommand.SQL.Clear;
                Connection.zCommand.SQL.Add('Update configuracion set txtValidaMaterial = :Medidas where sContrato = :Contrato');
                Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
                Connection.zCommand.Params.ParamByName('Medidas').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medidas').Value := ValidaMat;
                Connection.zCommand.ExecSQL;
              end;

              x := pos(ImpsMedida, MaterialAuto);
              if (x < 1) and (trim(ImpsMedida) <> '') then
              begin
                MaterialAuto := MaterialAuto + ImpsMedida + '|';
                Connection.zCommand.Active := False;
                Connection.zCommand.SQL.Clear;
                Connection.zCommand.SQL.Add('Update configuracion set txtMaterialAutomatico = :Medidas where sContrato = :Contrato');
                Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
                Connection.zCommand.Params.ParamByName('Medidas').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medidas').Value := MaterialAuto;
                Connection.zCommand.ExecSQL;
              end;
            end; //Fin else Actualiza....

            Fila := Fila + 1;
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            ProgressBar1.Max := ProgressBar1.Max + 1;
            ProgressBar1.Position := ProgressBar1.Position + 1;
          end; // Fin del while..

          try
            connection.zCommand.Active := false;
            connection.zCommand.SQL.Text := 'SET FOREIGN_KEY_CHECKS=1';
            connection.zCommand.ExecSQL;
            //Kardex('Importacion de Datos','Termina Proceso', 'Anexo C', '', '', '', '' );
            Kardex('Importacion de Datos', 'Termina Proceso', 'Anexo C', '', '', '', '','Tarifa Diaria','Importacion de Datos');
          except
            on e: exception do
              // Esta parte debe validarse por separado para evitar que si se encuentra un error en esta parte se cancele todo lo demas
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al registrar en kardex Importacion de Anexo C', 0);
          end;
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DT'}
        //Para Subir el anexo DT
        if rAnexoDT.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosDT('AnexoDT') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          iColumna := 4;
          progressBar1.Max := 0;
          sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
          while (DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]) <> '30/12/1899') and (iColumna <= 100) do
          begin
            dFecha := ExcelWorkSheet1.Cells.Item[1, iColumna];
            //sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
            DecodeDate(dFecha, myYear, myMonth, myDay);
            if myDay > 1 then
              dFecha := EncodeDate(myYear, myMonth, 1);
            sFecha := DateToStr(dFecha);

            ArrFechas[iColumna - 3] := sFecha;
            iColumna := iColumna + 1;
          end;

          iColumna := iColumna - 4;
          Fila := Fila + 1;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          if MessageDlg('Desea actualizar el anexo DT?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            {Consultamos si existen la distribucion del Anexo DT..}
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select * from anexosmensuales Where sContrato = :contrato And sIdConvenio = :Convenio');
            Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
            Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
            Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
            Connection.zCommand.Open;

            if connection.zCommand.RecordCount > 0 then
              if MessageDlg('Desea reemplzar la Distribucion Existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                chkBorrar.Checked := True;

            // Se elimina el anexo DT..
            if chkBorrar.Checked then
            begin
              CodErr1 := 'Importacion de Plantillas de Anexo DT';
              CodErr2 := 'Al Eliminar registro de anexosmensuales';

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('delete from anexosmensuales Where sContrato = :contrato And sIdConvenio = :Convenio');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
              connection.zCommand.ExecSQL();
            end
            else
            begin
              {Sino Actualizamos en 0 las cantidades..}
              CodErr1 := 'Importacion de Plantillas de Anexo DT';
              CodErr2 := 'Al inicializar a 0 el campos DT en anexosmensuales';

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('update anexosmensuales SET dt = 0 Where sContrato = :contrato And sIdConvenio = :Convenio');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
              connection.zCommand.ExecSQL();
            end;

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorkSheet1.Cells.Item[Fila, 1];

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsWbs := '';
              ImpsNumeroActividad := ExcelWorkSheet1.Cells.Item[Fila, 2];
              ImpsAnexo := ExcelWorkSheet1.Cells.Item[Fila, 3];

              ImpsWbsAnterior := '';
              dVentaMN := 0;
              dVentaDLL := 0;

              Connection.qryBusca.Active := False;
              Connection.qryBusca.SQL.Clear;
              Connection.qryBusca.SQL.Add('Select sWbs, dVentaMN, dVentaDLL From actividadesxanexo Where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad And sTipoActividad = "Actividad" and sAnexo =:Anexo ');
              Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
              Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
              Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Anexo').value := ImpsAnexo;
              Connection.qryBusca.Open;

              if Connection.QryBusca.RecordCount > 0 then
              begin
                ImpsWbsAnterior := Connection.QryBusca.FieldValues['sWbs'];
                dVentaMN := Connection.QryBusca.FieldValues['dVentaMN'];
                dVentaDLL := Connection.QryBusca.FieldValues['dVentaDLL'];
              end;

              for iRegistros := 1 to iColumna do
              begin
                CodErr1 := '';
                CodErr2 := '';

                sRecorrido := ExcelWorkSheet1.Cells.Item[Fila, iRegistros + 3];
                if (sRecorrido = ' ') or (sRecorrido = '  ') or (sRecorrido = '   ') or (sRecorrido = '    ') then
                  MessageDlg('PDA ' + ImpsNumeroActividad + ' CON ESPACIOS EN LA COLUMNA ' + IntToStr(iRegistros), mtInformation, [mbOk], 0)
                else
                  ImpfValor := ExcelWorkSheet1.Cells.Item[Fila, iRegistros + 3];

                if ImpfValor <> 0 then
                begin
                  Connection.qryBusca.Active := False;
                  Connection.qryBusca.SQL.Clear;
                  Connection.qryBusca.SQL.Add('Select dt From anexosmensuales Where sContrato = :Contrato And sIdConvenio = :Convenio And sWbs = :wbs ' +
                    'And sNumeroActividad = :Actividad And dIdFecha = :Fecha AND sAnexo = :Anexo');
                  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
                  Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Convenio').Value := global_Convenio;
                  Connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                  Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
                  Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
                  Connection.qryBusca.Params.ParamByName('Fecha').Value := StrToDate(ArrFechas[iRegistros]);
                  Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Anexo').Value := ImpsAnexo;
                  Connection.qryBusca.Open;

                  if Connection.qryBusca.RecordCount = 0 then
                  begin
                    CodErr1 := 'Importación de Plantilla de Anexo DT';
                    CodErr2 := 'Al Insetar registro en anexosmensuales';

                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('INSERT INTO anexosmensuales ( sContrato , sIdConvenio, dIdFecha, sWbs, sNumeroActividad, DT, DEmn, DEdll, sAnexo) ' +
                      ' VALUES (:contrato, :convenio, :fecha, :Wbs, :actividad, :dt, :DEmn, :DEdll, :Anexo)');
                    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                    Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                    Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                    Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                    Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                    Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                    Connection.zCommand.Params.ParamByName('DT').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('DT').value := ImpfValor;
                    Connection.zCommand.Params.ParamByName('DEmn').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('DEmn').value := ImpfValor * dVentaMN;
                    Connection.zCommand.Params.ParamByName('DEdll').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('DEdll').value := ImpfValor * dVentaDLL;
                    Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                    connection.zCommand.ExecSQL;
                  end
                  else
                  begin
                    if not SobreTodos then
                      Resp := MessageDlg('La partida de la Orden de Trabajo identificada en EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsNumeroOrden + ' - ' + Global_Convenio + ' - ' + sWbs + ' - ' + ImpsNumeroActividad + #10 + #10 +
                        '¿Desea sobreescribirla?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      CodErr1 := 'Importación de Plantilla de Anexo DT';
                      CodErr2 := 'Al Actualizar registro en anexosmensuales';

                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('UPDATE anexosmensuales SET DT = :dt, DEmn = :DEmn, DEdll = :DEdll, sAnexo = :Anexo Where sContrato = :Contrato And ' +
                        'sIdConvenio = :Convenio And sWbs = :Wbs And sNumeroActividad = :Actividad And dIdFecha = :Fecha');
                      Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                      Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                      Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                      Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                      Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                      Connection.zCommand.Params.ParamByName('DT').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DT').value := ImpfValor;
                      Connection.zCommand.Params.ParamByName('DT').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DT').value := ImpfValor + Connection.qryBusca.FieldValues['dt'];
                      Connection.zCommand.Params.ParamByName('DEmn').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DEmn').value := (ImpfValor + Connection.qryBusca.FieldValues['dt']) * dVentaMN;
                      Connection.zCommand.Params.ParamByName('DEdll').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DEdll').value := (ImpfValor + Connection.qryBusca.FieldValues['dt']) * dVentaDLL;
                      Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                      connection.zCommand.ExecSQL;
                    end;
                  end
                end
              end;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DT ESTRUCTURADO'}
        //Para Subir el anexo DT
        if rAnexoDTStruct.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosDT('AnexoDTStruct') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          iColumna := 6;
          progressBar1.Max := 0;
          sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
          while (DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]) <> '30/12/1899') and (iColumna <= 100) do
          begin
            dFecha := ExcelWorkSheet1.Cells.Item[1, iColumna];
            //sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
            DecodeDate(dFecha, myYear, myMonth, myDay);
            if myDay > 1 then
              dFecha := EncodeDate(myYear, myMonth, 1);
            sFecha := DateToStr(dFecha);

            ArrFechas[iColumna - 5] := sFecha;
            iColumna := iColumna + 1;
          end;

          iColumna := iColumna - 6;
          Fila := Fila + 1;

          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          if MessageDlg('Desea actualizar el anexo DT?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            {Consultamos si existen la distribucion del Anexo DT..}
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select * from anexosmensuales Where sContrato = :contrato And sIdConvenio = :Convenio');
            Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
            Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
            Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
            Connection.zCommand.Open;

            if connection.zCommand.RecordCount > 0 then
              if MessageDlg('Desea reemplzar la Distribucion Existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                chkBorrar.Checked := True;

            // Se elimina el anexo DT..
            if chkBorrar.Checked then
            begin
              CodErr1 := 'Importacion de Plantillas de Anexo DT';
              CodErr2 := 'Al Eliminar registro de anexosmensuales';

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('delete from anexosmensuales Where sContrato = :contrato And sIdConvenio = :Convenio');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
              connection.zCommand.ExecSQL();
            end
            else
            begin
              {Sino Actualizamos en 0 las cantidades..}
              CodErr1 := 'Importacion de Plantillas de Anexo DT';
              CodErr2 := 'Al inicializar a 0 el campos DT en anexosmensuales';

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('update anexosmensuales SET dt = 0 Where sContrato = :contrato And sIdConvenio = :Convenio');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
              connection.zCommand.ExecSQL();
            end;

            t := 1;
            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorkSheet1.Cells.Item[Fila, 1];

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              iNivel := ExcelWorkSheet1.Cells.Item[Fila, 2];
              ImpsNumeroActividad := ExcelWorkSheet1.Cells.Item[Fila, 3];
              ImpsAnexo := ExcelWorkSheet1.Cells.Item[Fila, 4];
              ImpsMedida := ExcelWorkSheet1.Cells.Item[Fila, 5];

              ImpsWbsAnterior := '';
              dVentaMN := 0;
              dVentaDLL := 0;


              if Trim(ImpsMedida) = '' then
                sTipo := 'Paquete'
              else
                sTipo := 'Actividad';

              sWbs := '';
              if iNivel <> 0 then
              begin
                for x := 1 to t - 1 do
                begin
                  if iNivel - 1 >= strToint(paquete[x][1]) then
                  begin
                    if (sTipo = 'Actividad') and (ImpsAnexo <> '') then
                      sWbs := paquete[x][2] + '.' + ImpsAnexo + '.'
                    else
                      sWbs := paquete[x][2] + '.';
                  end;
                end;
                sWbs := sWbs + ImpsNumeroActividad;
              end
              else
                sWbs := ImpsNumeroActividad;


              if sTipo = 'Paquete' then
              begin
                paquete[t][1] := inttostr(iNivel);
                paquete[t][2] := sWbs;
                paquete[t][3] := ImpiItemOrden;
                t := t + 1;
              end;


              Connection.qryBusca.Active := False;
              Connection.qryBusca.SQL.Clear;
              Connection.qryBusca.SQL.Add('Select sWbs, dVentaMN, dVentaDLL From actividadesxanexo Where sContrato = :Contrato And sIdConvenio = :Convenio and sWbs =:Wbs And sNumeroActividad = :Actividad  ' +
                'And sTipoActividad = "Actividad" and sAnexo =:Anexo order by iItemOrden');
              Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
              Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
              Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Anexo').value := ImpsAnexo;
              Connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Wbs').value := sWbs;
              Connection.qryBusca.Open;

              if Connection.QryBusca.RecordCount > 0 then
              begin
                ImpsWbsAnterior := Connection.QryBusca.FieldValues['sWbs'];
                dVentaMN := Connection.QryBusca.FieldValues['dVentaMN'];
                dVentaDLL := Connection.QryBusca.FieldValues['dVentaDLL'];
              end;

              for iRegistros := 1 to iColumna do
              begin
                CodErr1 := '';
                CodErr2 := '';

                sRecorrido := ExcelWorkSheet1.Cells.Item[Fila, iRegistros + 5];
                if (sRecorrido = ' ') or (sRecorrido = '  ') or (sRecorrido = '   ') or (sRecorrido = '    ') then
                  MessageDlg('PDA ' + ImpsNumeroActividad + ' CON ESPACIOS EN LA COLUMNA ' + IntToStr(iRegistros), mtInformation, [mbOk], 0)
                else
                  ImpfValor := ExcelWorkSheet1.Cells.Item[Fila, iRegistros + 5];

                if ImpfValor <> 0 then
                begin
                  if ImpsMedida <> '' then
                  begin
                    Connection.qryBusca.Active := False;
                    Connection.qryBusca.SQL.Clear;
                    Connection.qryBusca.SQL.Add('Select dt From anexosmensuales Where sContrato = :Contrato And sIdConvenio = :Convenio And sWbs = :wbs ' +
                      'And sNumeroActividad = :Actividad And dIdFecha = :Fecha AND sAnexo = :Anexo');
                    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
                    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
                    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
                    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_Convenio;
                    Connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
                    Connection.qryBusca.Params.ParamByName('Wbs').value := sWbs;
                    Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
                    Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
                    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
                    Connection.qryBusca.Params.ParamByName('Fecha').Value := StrToDate(ArrFechas[iRegistros]);
                    Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
                    Connection.qryBusca.Params.ParamByName('Anexo').Value := ImpsAnexo;
                    Connection.qryBusca.Open;

                    if Connection.qryBusca.RecordCount = 0 then
                    begin
                      CodErr1 := 'Importación de Plantilla de Anexo DT';
                      CodErr2 := 'Al Insetar registro en anexosmensuales';

                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('INSERT INTO anexosmensuales ( sContrato , sIdConvenio, dIdFecha, sWbs, sNumeroActividad, DT, DEmn, DEdll, sAnexo) ' +
                        ' VALUES (:contrato, :convenio, :fecha, :Wbs, :actividad, :dt, :DEmn, :DEdll, :Anexo)');
                      Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                      Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                      Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Wbs').value := sWbs;
                      Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                      Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                      Connection.zCommand.Params.ParamByName('DT').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DT').value := ImpfValor;
                      Connection.zCommand.Params.ParamByName('DEmn').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DEmn').value := ImpfValor * dVentaMN;
                      Connection.zCommand.Params.ParamByName('DEdll').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DEdll').value := ImpfValor * dVentaDLL;
                      Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                      connection.zCommand.ExecSQL;
                    end
                    else
                    begin
                      if not SobreTodos then
                        Resp := MessageDlg('La partida de la Orden de Trabajo identificada en EXCEL ya existe en la base de datos:' + #10 +
                          ImpsContrato + ' - ' + ImpsNumeroOrden + ' - ' + Global_Convenio + ' - ' + sWbs + ' - ' + ImpsNumeroActividad + #10 + #10 +
                          '¿Desea sobreescribirla?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                      if Resp = mrYesToAll then
                        SobreTodos := True;

                      if (Resp = mrYes) or SobreTodos then
                        Resp := mrYes;

                      if Resp = mrCancel then
                        raise Exception.Create('Proceso Cancelado por el Usuario.');

                      if Resp = mrYes then
                      begin
                        CodErr1 := 'Importación de Plantilla de Anexo DT';
                        CodErr2 := 'Al Actualizar registro en anexosmensuales';

                        connection.zCommand.Active := False;
                        connection.zCommand.SQL.Clear;
                        connection.zCommand.SQL.Add('UPDATE anexosmensuales SET DT = :dt, DEmn = :DEmn, DEdll = :DEdll, sAnexo = :Anexo Where sContrato = :Contrato And ' +
                          'sIdConvenio = :Convenio And sWbs = :Wbs And sNumeroActividad = :Actividad And dIdFecha = :Fecha');
                        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                        Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                        Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                        Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                        Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                        Connection.zCommand.Params.ParamByName('Wbs').value := sWbs;
                        Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                        Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                        Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                        Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                        Connection.zCommand.Params.ParamByName('DT').DataType := ftFloat;
                        Connection.zCommand.Params.ParamByName('DT').value := ImpfValor;
                        Connection.zCommand.Params.ParamByName('DT').DataType := ftFloat;
                        Connection.zCommand.Params.ParamByName('DT').value := ImpfValor + Connection.qryBusca.FieldValues['dt'];
                        Connection.zCommand.Params.ParamByName('DEmn').DataType := ftFloat;
                        Connection.zCommand.Params.ParamByName('DEmn').value := (ImpfValor + Connection.qryBusca.FieldValues['dt']) * dVentaMN;
                        Connection.zCommand.Params.ParamByName('DEdll').DataType := ftFloat;
                        Connection.zCommand.Params.ParamByName('DEdll').value := (ImpfValor + Connection.qryBusca.FieldValues['dt']) * dVentaDLL;
                        Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                        Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                        connection.zCommand.ExecSQL;
                      end;
                    end;
                  end
                end
              end;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DMA, DMO, DME'}
        //Para Subir el anexo DMA, DMO o DME
        if (rAnexoDMA.Checked) or (rAnexoDMO.Checked) or (rAnexoDME.Checked) then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if rAnexoDMA.Checked then
          begin
            if ValidaAnexosDME('Material', 'insumos', 'sIdInsumo') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'distribuciondematerial';
            ImpsWbs := 'sIdMaterial';
            ImpsAnexo := 'DMA';
          end;

          if rAnexoDMO.Checked then
          begin
            if ValidaAnexosDME('Personal', 'personal', 'sIdPersonal') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'distribuciondepersonal';
            ImpsWbs := 'sIdPersonal';
            ImpsAnexo := 'DMO';
          end;

          if rAnexoDME.Checked then
          begin
            if ValidaAnexosDME('Equipo', 'equipos', 'sIdEquipo') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'distribuciondeequipos';
            ImpsWbs := 'sIdEquipo';
            ImpsAnexo := 'DME';
          end;

          iColumna := 3;
          sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
          while (DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]) <> '30/12/1899') and (iColumna <= 100) do
          begin
            sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
            ArrFechas[iColumna - 2] := sFecha;
            iColumna := iColumna + 1;
          end;

          iColumna := iColumna - 3;
          progressBar1.Max := 0;
          Fila := Fila + 1;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          //chkBorrar.Checked := True;
          if MessageDlg('Desea Actualizar el anexo ' + ImpsAnexo + '?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            // Se elimina el DMA existente
            if chkBorrar.Checked then
            begin
              CodErr1 := 'Importación de plantillas de anexos DMA, DMO, DME';
              CodErr2 := 'Al borrar registros en ' + sTabla;

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('delete from ' + sTabla + ' Where sContrato = :contrato');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              connection.zCommand.ExecSQL();
            end;

  //**********************************************************
            // Generar una lista de registros que deben existir
            Existir.Close;
            Existir.FieldDefs.Add('sContrato', ftString, 15);
            Existir.FieldDefs.Add('sIdConvenio', ftString, 100);
            Existir.FieldDefs.Add(ImpsWbs, ftString, 25);
            Existir.FieldDefs.Add('dIdFecha', ftDate, 0);
            Existir.Open;
            Existir.EmptyTable;
  //**********************************************************

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorkSheet1.Cells.Item[Fila, 1];

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsNumeroActividad := ExcelWorkSheet1.Cells.Item[Fila, 2];

              for iRegistros := 1 to iColumna do
              begin
                CodErr1 := '';
                CodErr2 := '';

                ImpfValor := ExcelWorkSheet1.Cells.Item[Fila, iRegistros + 2];

                Connection.qryBusca.Active := False;
                Connection.qryBusca.SQL.Clear;
                Connection.qryBusca.SQL.Add('Select dCantidad From ' + sTabla + ' Where sContrato = :Contrato ' +
                  'And ' + ImpsWbs + ' = :idDistribucion And dIdFecha = :Fecha');
                Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
                Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
                Connection.qryBusca.Params.ParamByName('idDistribucion').DataType := ftString;
                Connection.qryBusca.Params.ParamByName('idDistribucion').Value := ImpsNumeroActividad;
                Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
                Connection.qryBusca.Params.ParamByName('Fecha').Value := StrToDate(ArrFechas[iRegistros]);
                Connection.qryBusca.Open;
                if Connection.qryBusca.RecordCount = 0 then
                begin
                  CodErr1 := 'Importación de plantillas de anexos DMA, DMO, DME';
                  CodErr2 := 'Al insertar registros en ' + sTabla;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zCommand.SQL.Add('INSERT INTO ' + sTabla + ' ( sContrato, dIdFecha, ' + ImpsWbs + ', dCantidad) ' +
                    ' VALUES (:contrato, :fecha, :idDistribucion, :cantidad)');
                  Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                  Connection.zCommand.Params.ParamByName('idDistribucion').DataType := ftString;
                  Connection.zCommand.Params.ParamByName('idDistribucion').value := ImpsNumeroActividad;
                  Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                  Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                  Connection.zCommand.Params.ParamByName('cantidad').DataType := ftFloat;
                  Connection.zCommand.Params.ParamByName('cantidad').value := ImpfValor;
                  connection.zCommand.ExecSQL;
                end
                else
                begin
                  if not SobreTodos then
                    Resp := MessageDlg('El Recurso ' + ImpsAnexo + ' identificado en EXCEL ya existe en la base de datos:' + #10 +
                      ImpsContrato + ' - ' + ImpsIsometrico + ' - ' + ImpiRevision + #10 + #10 +
                      '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                  if Resp = mrYesToAll then
                    SobreTodos := True;

                  if (Resp = mrYes) or SobreTodos then
                    Resp := mrYes;

                  if Resp = mrCancel then
                    raise Exception.Create('Proceso Cancelado por el Usuario.');

                  if Resp = mrYes then
                  begin
                    CodErr1 := 'Importación de plantillas de anexos DMA, DMO, DME';
                    CodErr2 := 'Al actualizar registros en ' + sTabla;

                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('UPDATE ' + sTabla + ' SET dCantidad = :cantidad Where sContrato = :Contrato And ' +
                      ' ' + ImpsWbs + ' = :idDistribucion And dIdFecha = :Fecha');
                    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                    Connection.zCommand.Params.ParamByName('idDistribucion').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('idDistribucion').value := ImpsNumeroActividad;
                    Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                    Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                    Connection.zCommand.Params.ParamByName('cantidad').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('cantidad').value := ImpfValor;
                    connection.zCommand.ExecSQL;
                  end;
                end;
                  // Cargar la lista de registros procesados
                Existir.Append;
                Existir.FieldByName('sContrato').AsString := ImpsContrato;
                Existir.FieldByName('sIdConvenio').AsString := Global_Convenio;
                Existir.FieldByName(ImpsWbs).AsString := ImpsNumeroActividad;
                Existir.FieldByName('dIdFecha').AsDateTime := StrToDate(ArrFechas[iRegistros]);
                Existir.Post;
              end;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];
            end;
          end;
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DE'}
        //Anexo DE Moneda Nacional....
        if rAnexoDE.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosDT('AnexoDEMN') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          iColumna := 4;
          progressBar1.Max := 0;
          sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
          while (DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]) <> '30/12/1899') and (iColumna <= 100) do
          begin
            sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
            ArrFechas[iColumna - 3] := sFecha;
            iColumna := iColumna + 1;
          end;
          iColumna := iColumna - 4;
          Fila := Fila + 1;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          if MessageDlg('Desea actualizar el anexo DE en Moneda Nacional?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            if chkBorrar.Checked then
            begin
              CodErr1 := 'Importación de Plantilla de Anexo DE';
              CodErr2 := 'Al inicializar a 0 registros en anexosmensuales';

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('update anexosmensuales SET DEmn = 0 Where sContrato = :contrato And sIdConvenio = :Convenio');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
              connection.zCommand.ExecSQL();
            end;

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorkSheet1.Cells.Item[Fila, 1];

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsWbs := '';
              ImpsNumeroActividad := ExcelWorkSheet1.Cells.Item[Fila, 2];
              ImpsAnexo := ExcelWorkSheet1.Cells.Item[Fila, 3];

              ImpsWbsAnterior := '';
              dVentaMN := 0;
              dVentaDLL := 0;

              Connection.qryBusca.Active := False;
              Connection.qryBusca.SQL.Clear;
              Connection.qryBusca.SQL.Add('Select sWbs, dVentaMN, dVentaDLL From actividadesxanexo Where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad And sTipoActividad = "Actividad" and sAnexo =:Anexo ');
              Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
              Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
              Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Anexo').value := ImpsAnexo;
              Connection.qryBusca.Open;
              if Connection.QryBusca.RecordCount > 0 then
              begin
                ImpsWbsAnterior := Connection.QryBusca.FieldValues['sWbs'];
                dVentaMN := Connection.QryBusca.FieldValues['dVentaMN'];
                dVentaDLL := Connection.QryBusca.FieldValues['dVentaDLL'];
              end;

              for iRegistros := 1 to iColumna do
              begin
                CodErr1 := '';
                CodErr2 := '';

                ImpfValor := ExcelWorkSheet1.Cells.Item[Fila, iRegistros + 3];
                if ImpfValor <> 0 then
                begin
                  Connection.qryBusca.Active := False;
                  Connection.qryBusca.SQL.Clear;
                  Connection.qryBusca.SQL.Add('Select DEmn From anexosmensuales Where sContrato = :Contrato And sIdConvenio = :Convenio And sWbs = :wbs ' +
                    'And sNumeroActividad = :Actividad And dIdFecha = :Fecha and sAnexo =:Anexo ');
                  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
                  Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Convenio').Value := global_Convenio;
                  Connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                  Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
                  Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
                  Connection.qryBusca.Params.ParamByName('Fecha').Value := StrToDate(ArrFechas[iRegistros]);
                  Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Anexo').value := ImpsAnexo;
                  Connection.qryBusca.Open;
                  if Connection.qryBusca.RecordCount = 0 then
                  begin
                    CodErr1 := 'Importación de Plantilla de Anexo DE';
                    CodErr2 := 'Al insertar registros en anexosmensuales';

                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('INSERT INTO anexosmensuales ( sContrato , sIdConvenio, dIdFecha, sWbs, sNumeroActividad, DEmn, sAnexo) ' +
                      ' VALUES (:contrato, :convenio, :fecha, :Wbs, :actividad, :DEmn, :Anexo)');
                    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                    Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                    Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                    Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                    Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                    Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                    Connection.zCommand.Params.ParamByName('DEmn').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('DEmn').value := ImpfValor;
                    Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                    connection.zCommand.ExecSQL;
                  end
                  else
                  begin
                    if not SobreTodos then
                      Resp := MessageDlg('La partida de la Orden de Trabajo identificada en EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsIsometrico + ' - ' + ImpiRevision + #10 + #10 +
                        '¿Desea sobreescribirla?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      CodErr1 := 'Importación de Plantilla de Anexo DE';
                      CodErr2 := 'Al actualizar registros en anexosmensuales';

                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('UPDATE anexosmensuales SET DEmn = :DEmn Where sContrato = :Contrato And ' +
                        'sIdConvenio = :Convenio And sWbs = :Wbs And sNumeroActividad = :Actividad And dIdFecha = :Fecha and sAnexo =:Anexo ');
                      Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                      Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                      Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                      Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                      Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                      Connection.zCommand.Params.ParamByName('DEmn').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DEmn').value := ImpfValor + connection.QryBusca.FieldValues['DEmn'];
                      Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                      connection.zCommand.ExecSQL;
                    end;
                  end
                end;
              end;
              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DE DLL'}
        //Anexo DE DLL
        if rAnexoDEDLL.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosDT('AnexoDEDLL') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          iColumna := 4;
          progressBar1.Max := 0;
          sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
          while (DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]) <> '30/12/1899') and (iColumna <= 100) do
          begin
            sFecha := DateToStr(ExcelWorkSheet1.Cells.Item[1, iColumna]);
            ArrFechas[iColumna - 3] := sFecha;
            iColumna := iColumna + 1;
          end;
          iColumna := iColumna - 4;
          Fila := Fila + 1;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          if MessageDlg('Desea actualizar el anexo DE en Moneda Extranjera?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            if chkBorrar.Checked then
            begin
              CodErr1 := 'Importación de Plantilla de Anexo DE (Dolares)';
              CodErr2 := 'Al poner a 0 el campo DEdll en la tabla anexosmensuales';

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('update anexosmensuales SET DEdll = 0 Where sContrato = :contrato And sIdConvenio = :Convenio');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value := Global_Convenio;
              connection.zCommand.ExecSQL();
            end;

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorkSheet1.Cells.Item[Fila, 1];

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsWbs := '';
              ImpsNumeroActividad := ExcelWorkSheet1.Cells.Item[Fila, 2];
              ImpsAnexo := ExcelWorkSheet1.Cells.Item[Fila, 3];

              ImpsWbsAnterior := '';
              dVentaMN := 0;
              dVentaDLL := 0;
              Connection.qryBusca.Active := False;
              Connection.qryBusca.SQL.Clear;
              Connection.qryBusca.SQL.Add('Select sWbs, dVentaMN, dVentaDLL From actividadesxanexo Where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad And sTipoActividad = "Actividad" and sAnexo =:Anexo ');
              Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
              Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
              Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
              Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
              Connection.qryBusca.Params.ParamByName('Anexo').value := ImpsAnexo;
              Connection.qryBusca.Open;
              if Connection.QryBusca.RecordCount > 0 then
              begin
                ImpsWbsAnterior := Connection.QryBusca.FieldValues['sWbs'];
                dVentaMN := Connection.QryBusca.FieldValues['dVentaMN'];
                dVentaDLL := Connection.QryBusca.FieldValues['dVentaDLL'];
              end;

              for iRegistros := 1 to iColumna do
              begin
                CodErr1 := '';
                CodErr2 := '';

                ImpfValor := ExcelWorkSheet1.Cells.Item[Fila, iRegistros + 3];
                if ImpfValor <> 0 then
                begin
                  Connection.qryBusca.Active := False;
                  Connection.qryBusca.SQL.Clear;
                  Connection.qryBusca.SQL.Add('Select DEdll From anexosmensuales Where sContrato = :Contrato And sIdConvenio = :Convenio And sWbs = :wbs ' +
                    'And sNumeroActividad = :Actividad And dIdFecha = :Fecha and sAnexo =:Anexo');
                  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
                  Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Convenio').Value := global_Convenio;
                  Connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                  Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
                  Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
                  Connection.qryBusca.Params.ParamByName('Fecha').Value := StrToDate(ArrFechas[iRegistros]);
                  Connection.qryBusca.Params.ParamByName('Anexo').DataType := ftString;
                  Connection.qryBusca.Params.ParamByName('Anexo').value := ImpsAnexo;
                  Connection.qryBusca.Open;
                  if Connection.qryBusca.RecordCount = 0 then
                  begin
                    CodErr1 := 'Importación de Plantilla de Anexo DE (Dolares)';
                    CodErr2 := 'Al insertar registros en la tabla anexosmensuales';

                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('INSERT INTO anexosmensuales ( sContrato , sIdConvenio, dIdFecha, sWbs, sNumeroActividad, DEdll, sAnexo) ' +
                      ' VALUES (:contrato, :convenio, :fecha, :Wbs, :actividad, :DEdll, :Anexo)');
                    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                    Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                    Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                    Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                    Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                    Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                    Connection.zCommand.Params.ParamByName('DEdll').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('DEdll').value := ImpfValor;
                    Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                    connection.zCommand.ExecSQL;
                  end
                  else
                  begin
                    if not SobreTodos then
                      Resp := MessageDlg('La partida de la Orden de Trabajo identificada en EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsIsometrico + ' - ' + ImpiRevision + #10 + #10 +
                        '¿Desea sobreescribirla?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      CodErr1 := 'Importación de Plantilla de Anexo DE (Dolares)';
                      CodErr2 := 'Al actualizar registros en la tabla anexosmensuales';

                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('UPDATE anexosmensuales SET DEdll = :DEdll Where sContrato = :Contrato And ' +
                        'sIdConvenio = :Convenio And sWbs = :Wbs And sNumeroActividad = :Actividad And dIdFecha = :Fecha and sAnexo =:Anexo ');
                      Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato;
                      Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio;
                      Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Wbs').value := ImpsWbsAnterior;
                      Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
                      Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('fecha').value := StrToDate(ArrFechas[iRegistros]);
                      Connection.zCommand.Params.ParamByName('DEdll').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('DEdll').value := ImpfValor + connection.QryBusca.FieldValues['DEdll'];
                      Connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Anexo').value := ImpsAnexo;
                      connection.zCommand.ExecSQL;
                    end;
                  end
                end;
              end;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorkSheet1.Cells.Item[Fila, 1];
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DE PERSONAL'}
        //ANEXO DE PERSONAL..
        if rAnexoPersonal.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosPE('Personal') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          Fila := 2;
          ProgressBar1.Max := 0;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          //if MessageDlg('Desea remplazar el catalogo de personal existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            if ValidaDeleteAnexosP('personal', 'sIdPersonal', 'bitacoradepersonal', 'recursospersonal') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            {Solucion Integridad referencial grupospersonal _ personal... 11 Junio 2011..}
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select * from grupospersonal where sIdGrupo ="S/C"');
            connection.zCommand.Open;

            if connection.zCommand.RecordCount = 0 then
            begin
              try
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('insert into grupospersonal (sIdGrupo, sDescripcion, iOrden) values ("S/C", "SIN CATEGORIA", 0)');
                connection.zCommand.ExecSQL;
              except
              end;
            end;

            {Solucion Integridad referencial tipodepersonal _ personal... 11 Junio 2011..}
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select * from tiposdepersonal where sIdTipoPersonal ="PE-C"');
            connection.zCommand.Open;

            if connection.zCommand.RecordCount = 0 then
            begin
              try
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('insert into tiposdepersonal (sIdTipoPersonal, sDescripcion, sMascara) values ("PE-C", "PERSONAL DE CONSTRUCCION", "PC-")');
                connection.zCommand.ExecSQL;
              except
              end;
            end;
            {Termina solucion integridad..}

            {Se insertan los datos de personal..}
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('INSERT INTO personal ( sContrato, sIdPersonal, iItemOrden, sDescripcion, sIdTipoPersonal, sMedida, dCantidad, ' +
              'dCostoMN, dCostoDLL, dVentaMN, dVentaDLL, dFechaInicio, dFechaFinal, lProrrateo, lCobro, lImprime, lAplicaTM, iJornada, lDistribuye, sAgrupaPersonal ) ' +
              ' VALUES (:contrato, :Personal, :Orden, :Descripcion, :Tipo, :Medida, :Cantidad, :CostoMN, :CostoDLL, :VentaMN, :VentaDLL, :FechaI, :FechaF, "Si", "Si", "Si", "Si" , :Jornada, "Si", :AgrupaP )');

  //**********************************************************
            // Generar una lista de registros que deben existir
            Existir.Close;
            Existir.FieldDefs.Add('sContrato', ftString, 15);
            Existir.FieldDefs.Add('sIdPersonal', ftString, 25);
            Existir.FieldDefs.Add('sIdTipoPersonal', ftString, 4);
            Existir.Open;
            Existir.EmptyTable;
  //**********************************************************

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              ImpiItemOrden := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
              ImpmDescripcion := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              ImpsMedida := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
              ImpdCantidadAnexo := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoMN := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoDLL := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaMN := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaDLL := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
              ImpdFechaInicio := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;
              ImpdFechaFinal := ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;
              ImpsTipo := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;
              ImpsGrupo := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;

              // Inserto Datos a la Tabla .....
              CodErr1 := 'Importación de Plantilla de Anexo de Personal';
              CodErr2 := 'Al tratar de insertar registros en la tabla personal';

              try
                connection.zCommand.Active := False;
                Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                Connection.zCommand.Params.ParamByName('Personal').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Personal').value := ImpsNumeroActividad;
                Connection.zCommand.Params.ParamByName('Tipo').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Tipo').value := ImpsTipo;
                Connection.zCommand.Params.ParamByName('Orden').DataType := ftInteger;
                Connection.zCommand.Params.ParamByName('Orden').value := ImpiItemOrden;
                Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                Connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
                Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medida').value := ImpsMedida;
                Connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('Cantidad').value := ImpdCantidadAnexo;
                Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoMN').value := ImpdCostoMN;
                Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoDLL').value := ImpdCostoDLL;
                Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaMN').value := ImpdVentaMN;
                Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaDLL').value := ImpdVentaDLL;
                Connection.zCommand.Params.ParamByName('FechaI').DataType := ftDate;
                Connection.zCommand.Params.ParamByName('FechaI').value := ImpdFechaInicio;
                Connection.zCommand.Params.ParamByName('FechaF').DataType := ftDate;
                Connection.zCommand.Params.ParamByName('FechaF').value := ImpdFechaFinal;
                Connection.zCommand.Params.ParamByName('Jornada').DataType := ftInteger;
                Connection.zCommand.Params.ParamByName('Jornada').value := ImpsGrupo;
                Connection.zCommand.Params.ParamByName('AgrupaP').DataType := ftString;
                Connection.zCommand.Params.ParamByName('AgrupaP').value := 'S/C';
                Connection.zCommand.ExecSQL;
  //**********************************************************
              except
                on e: exception do
                begin
                  // Verificar si se encontró una duplicidad de registros
                  if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                  begin
                    // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro
                    if not SobreTodos then
                      Resp := MessageDlg('El registro de personal identificado en EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsNumeroActividad + #10 + #10 +
                        '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      connection.QryBusca2.Active := False;
                      connection.QryBusca2.SQL.Clear;
                      connection.QryBusca2.SQL.Add('UPDATE personal SET iItemOrden = :Orden, sDescripcion = :Descripcion, sMedida = :Medida, ' +
                        'dCantidad = :Cantidad, dCostoMN = :CostoMN, dCostoDLL = :CostoDLL, dVentaMN =:VentaMN, ' +
                        'dVentaDLL = :VentaDLL, iJornada = :Jornada ' +
                        'WHERE sContrato = :Contrato and sIdPersonal = :Personal and sIdTipoPersonal =:Tipo and sAgrupaPersonal =:Agrupa ');
                      Connection.QryBusca2.Params.ParamByName('contrato').DataType := ftString;
                      Connection.QryBusca2.Params.ParamByName('contrato').AsString := ImpsContrato;
                      Connection.QryBusca2.Params.ParamByName('Personal').DataType := ftString;
                      Connection.QryBusca2.Params.ParamByName('Personal').AsString := ImpsNumeroActividad;
                      Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
                      Connection.QryBusca2.Params.ParamByName('Orden').AsString := ImpiItemOrden;
                      Connection.QryBusca2.Params.ParamByName('Tipo').DataType := ftString;
                      Connection.QryBusca2.Params.ParamByName('Tipo').AsString := ImpsTipo;
                      Connection.QryBusca2.Params.ParamByName('Descripcion').DataType := ftMemo;
                      Connection.QryBusca2.Params.ParamByName('Descripcion').AsString := Trim(ImpmDescripcion);
                      Connection.QryBusca2.Params.ParamByName('Medida').DataType := ftString;
                      Connection.QryBusca2.Params.ParamByName('Medida').AsString := ImpsMedida;
                      Connection.QryBusca2.Params.ParamByName('Cantidad').DataType := ftFloat;
                      Connection.QryBusca2.Params.ParamByName('Cantidad').AsString := ImpdCantidadAnexo;
                      Connection.QryBusca2.Params.ParamByName('CostoMN').DataType := ftFloat;
                      Connection.QryBusca2.Params.ParamByName('CostoMN').AsString := ImpdCostoMN;
                      Connection.QryBusca2.Params.ParamByName('CostoDLL').DataType := ftFloat;
                      Connection.QryBusca2.Params.ParamByName('CostoDLL').AsString := ImpdCostoDLL;
                      Connection.QryBusca2.Params.ParamByName('VentaMN').DataType := ftFloat;
                      Connection.QryBusca2.Params.ParamByName('VentaMN').AsString := ImpdVentaMN;
                      Connection.QryBusca2.Params.ParamByName('VentaDLL').DataType := ftFloat;
                      Connection.QryBusca2.Params.ParamByName('VentaDLL').AsString := ImpdVentaDLL;
                      Connection.QryBusca2.Params.ParamByName('Jornada').DataType := ftInteger;
                      Connection.QryBusca2.Params.ParamByName('Jornada').AsString := ImpsGrupo;
                      Connection.QryBusca2.Params.ParamByName('Agrupa').DataType := ftString;
                      Connection.QryBusca2.Params.ParamByName('Agrupa').AsString := 'S/C';
                      Connection.QryBusca2.ExecSQL;
                    end;
                  end
                  else
                    raise;
                end;
              end;

              // Cargar la lista de registros procesados
              Existir.Append;
              Existir.FieldByName('sContrato').AsString := ImpsContrato;
              Existir.FieldByName('sIdPersonal').AsString := ImpsNumeroActividad;
              Existir.FieldByName('sIdTipoPersonal').AsString := ImpsTipo;
              Existir.Post;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            end;
  //**********************************************************
            // Verificar los registros que debería ser eliminados
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Select * from personal where sContrato = :Contrato ');
            Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
            Connection.zCommand.Open;

            if Connection.zCommand.RecordCount > Existir.RecordCount then
            begin
              Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
                '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
              if Resp = mrCancel then
                raise Exception.Create('Proceso Cancelado por el Usuario.');

              if Resp = mrYes then
              begin
                connection.zCommand.First;
                while not connection.zCommand.Eof do
                begin
                  if not Existir.Locate('sIdPersonal', connection.zCommand.FieldByName('sIdPersonal').AsString, []) then
                    Connection.zCommand.Delete;
                  connection.zCommand.Next;
                end;
              end;
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ANEXO EQUIPO'}
        //ANEXO DE EQUIPO
        if rAnexoEquipo.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosPE('Equipo') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          Fila := 2;
          ProgressBar1.Max := 0;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          //if MessageDlg('Desea remplazar el catalogo de Equipos Existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            if ValidaDeleteAnexosP('equipos', 'sIdEquipo', 'bitacoradeequipos', 'recursosequipo') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            {Solucion Integridad referencial tipodeequipo _ equipos... 11 Junio 2011..}
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select * from tiposdeequipo where sIdTipoEquipo ="EQ-C"');
            connection.zCommand.Open;

            if connection.zCommand.RecordCount = 0 then
            begin
              try
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('insert into tiposdeequipo (sIdTipoEquipo, sDescripcion, sMascara) values ("EQ-C", "EQUIPO DE CONSTRUCCION", "EC-")');
                connection.zCommand.ExecSQL;
              except
              end;
            end;
            {Termina solucion integridad..}

            // Generar una lista de registros que deben existir
            Existir.Close;
            Existir.FieldDefs.Add('sContrato', ftString, 15);
            Existir.FieldDefs.Add('sIdEquipo', ftString, 25);
            Existir.Open;
            Existir.EmptyTable;
  //**********************************************************

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              ImpiItemOrden := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
              ImpmDescripcion := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              ImpsMedida := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
              ImpdCantidadAnexo := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoMN := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoDLL := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaMN := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaDLL := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
              ImpdFechaInicio := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;
              ImpdFechaFinal := ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;
              ImpsTipo := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;
              ImpsGrupo := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;

              // Inserto Datos a la Tabla .....
              CodErr1 := 'Importación de Plantilla de Anexo de Equipos';
              CodErr2 := 'Al tratar de insertar registros en la tabla equipos';

              try
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('INSERT INTO equipos ( sContrato, sIdEquipo, iItemOrden, sDescripcion, sIdTipoEquipo, sMedida, dCantidad, ' +
                  'dCostoMN, dCostoDLL, dVentaMN, dVentaDLL, dFechaInicio, dFechaFinal, lProrrateo, lCobro, lImprime, iJornada, lDistribuye ) ' +
                  ' VALUES (:contrato, :Equipo, :Orden, :Descripcion, :Tipo, :Medida, :Cantidad, :CostoMN, :CostoDLL, :VentaMN, :VentaDLL, :FechaI, :FechaF, "Si", "Si", "Si", :Jornada, "Si" )');
                Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                Connection.zCommand.Params.ParamByName('Equipo').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Equipo').value := ImpsNumeroActividad;
                Connection.zCommand.Params.ParamByName('Tipo').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Tipo').value := ImpsTipo;
                Connection.zCommand.Params.ParamByName('Orden').DataType := ftInteger;
                Connection.zCommand.Params.ParamByName('Orden').value := ImpiItemOrden;
                Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                Connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
                Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medida').value := ImpsMedida;
                Connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('Cantidad').value := ImpdCantidadAnexo;
                Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoMN').value := ImpdCostoMN;
                Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoDLL').value := ImpdCostoDLL;
                Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaMN').value := ImpdVentaMN;
                Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaDLL').value := ImpdVentaDLL;
                Connection.zCommand.Params.ParamByName('FechaI').DataType := ftDate;
                Connection.zCommand.Params.ParamByName('FechaI').value := ImpdFechaInicio;
                Connection.zCommand.Params.ParamByName('FechaF').DataType := ftDate;
                Connection.zCommand.Params.ParamByName('FechaF').value := ImpdFechaFinal;
                Connection.zCommand.Params.ParamByName('Jornada').DataType := ftInteger;
                Connection.zCommand.Params.ParamByName('Jornada').value := ImpsGrupo;
                Connection.zCommand.ExecSQL;
  //**********************************************************
              except
                on e: exception do
                begin
                  // Verificar si se encontró una duplicidad de registros
                  if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                  begin
                    // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro
                    if not SobreTodos then
                      Resp := MessageDlg('El registro de equipo identificado en EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsNumeroActividad + #10 + #10 +
                        '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('UPDATE equipos SET iItemOrden = :Orden, sDescripcion = :Descripcion, sIdTipoEquipo =:Tipo, ' +
                        'sMedida = :Medida, dCantidad = :Cantidad, dCostoMN = :CostoMN, dCostoDLL = :CostoDLL, ' +
                        'dVentaMN = :VentaMN, dVentaDLL = :VentaDLL, dFechaInicio = :FechaI, dFechaFinal = :FechaF, ' +
                        'lProrrateo = "Si", lCobro = "Si", lImprime = "Si", iJornada = :Jornada, lDistribuye = "Si" ' +
                        'WHERE sContrato = :Contrato and sIdEquipo = :Equipo');
                      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Contrato').Value := ImpsContrato;
                      Connection.zCommand.Params.ParamByName('Equipo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Equipo').Value := ImpsNumeroActividad;
                      Connection.zCommand.Params.ParamByName('Tipo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Tipo').Value := ImpsTipo;
                      Connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Orden').Value := ImpiItemOrden;
                      Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                      Connection.zCommand.Params.ParamByName('Descripcion').Value := Trim(ImpmDescripcion);
                      Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Medida').Value := ImpsMedida;
                      Connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('Cantidad').Value := ImpdCantidadAnexo;
                      Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('CostoMN').Value := ImpdCostoMN;
                      Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('CostoDLL').Value := ImpdCostoDLL;
                      Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('VentaMN').Value := ImpdVentaMN;
                      Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('VentaDLL').Value := ImpdVentaDLL;
                      Connection.zCommand.Params.ParamByName('FechaI').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('FechaI').Value := ImpdFechaInicio;
                      Connection.zCommand.Params.ParamByName('FechaF').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('FechaF').Value := ImpdFechaFinal;
                      Connection.zCommand.Params.ParamByName('Jornada').DataType := ftInteger;
                      Connection.zCommand.Params.ParamByName('Jornada').Value := ImpsGrupo;
                      Connection.zCommand.ExecSQL;
                    end;
                  end
                  else
                    raise;
                end;
              end;

              // Cargar la lista de registros procesados
              Existir.Append;
              Existir.FieldByName('sContrato').AsString := ImpsContrato;
              Existir.FieldByName('sIdEquipo').AsString := ImpsNumeroActividad;
              Existir.Post;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            end;
  //**********************************************************
            // Verificar los registros que debería ser eliminados
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Select * from equipos where sContrato = :Contrato');
            Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
            Connection.zCommand.Open;

            if Connection.zCommand.RecordCount > Existir.RecordCount then
            begin
              Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
                '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
              if Resp = mrCancel then
                raise Exception.Create('Proceso Cancelado por el Usuario.');

              if Resp = mrYes then
              begin
                connection.zCommand.First;
                while not connection.zCommand.Eof do
                begin
                  if not Existir.Locate('sIdEquipo', connection.zCommand.FieldByName('sIdEquipo').AsString, []) then
                    Connection.zCommand.Delete;
                  connection.zCommand.Next;
                end;
              end;
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DE BASICOS'}
        //ANEXO DE BASICOS..
        if rAnexoBasicos.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosBasicos('Basico') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          Fila := 2;
          ProgressBar1.Max := 0;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          //if MessageDlg('Desea remplazar el catalogo de Basicos existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            if ValidaDeleteAnexosP('basicos', 'sIdBasico', '', 'recursosbasicos') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

  //**********************************************************
            // Generar una lista de registros que deben existir
            Existir.Close;
            Existir.FieldDefs.Add('sContrato', ftString, 15);
            Existir.FieldDefs.Add('sIdBasico', ftString, 25);
            Existir.Open;
            Existir.EmptyTable;
  //**********************************************************

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              ImpmDescripcion := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
              ImpsMedida := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoMN := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoDLL := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaMN := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaDLL := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;

              try
                {Se insertan los datos de personal..}
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('INSERT INTO basicos ( sContrato, sIdBasico, sDescripcion, sMedida, dCostoMN, dCostoDLL, dVentaMN, dVentaDLL, sSimbolo) ' +
                  ' VALUES (:contrato, :Basico, :Descripcion, :Medida, :CostoMN, :CostoDLL, :VentaMN, :VentaDLL, "")');

                // Inserto Datos a la Tabla .....
                CodErr1 := 'Importación de Plantilla de Anexo de Básicos';
                CodErr2 := 'Al tratar de insertar registros en la tabla basicos';

                connection.zCommand.Active := False;
                Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                Connection.zCommand.Params.ParamByName('Basico').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Basico').value := ImpsNumeroActividad;
                Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                Connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
                Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medida').value := ImpsMedida;
                Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoMN').value := ImpdCostoMN;
                Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoDLL').value := ImpdCostoDLL;
                Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaMN').value := ImpdVentaMN;
                Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaDLL').value := ImpdVentaDLL;
                Connection.zCommand.ExecSQL;
              except
                on e: exception do
                begin
                  // Verificar si se encontró una duplicidad de registros
                  if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                  begin
                    // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro
                    if not SobreTodos then
                      Resp := MessageDlg('El registro de basicos identificado en EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsNumeroActividad + #10 + #10 +
                        '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('UPDATE basicos SET sDescripcion = :Descripcion, sMedida = :Medida, dCostoMN = :CostoMN, ' +
                        'dCostoDLL = :CostoDLL, dVentaMN = :VentaMN, dVentaDLL = :VentaDLL, sSimbolo = :Simbolo ' +
                        'WHERE sContrato = :Contrato and sIdBasico = :Basico');
                      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Contrato').Value := ImpsContrato;
                      Connection.zCommand.Params.ParamByName('Basico').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Basico').Value := ImpsNumeroActividad;
                      Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                      Connection.zCommand.Params.ParamByName('Descripcion').Value := Trim(ImpmDescripcion);
                      Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Medida').Value := ImpsMedida;
                      Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('CostoMN').Value := ImpdCostoMN;
                      Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('CostoDLL').Value := ImpdCostoDLL;
                      Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('VentaMN').Value := ImpdVentaMN;
                      Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('VentaDLL').Value := ImpdVentaDLL;
                      Connection.zCommand.Params.ParamByName('Simbolo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Simbolo').Value := '';
                      Connection.zCommand.ExecSQL;
                    end;
                  end
                  else
                    raise;
                end;
              end;

              // Cargar la lista de registros procesados
              Existir.Append;
              Existir.FieldByName('sContrato').AsString := ImpsContrato;
              Existir.FieldByName('sIdBasico').AsString := ImpsNumeroActividad;
              Existir.Post;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            end;

  //**********************************************************
            // Verificar los registros que debería ser eliminados
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Select * from basicos where sContrato = :Contrato');
            Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
            Connection.zCommand.Open;

            if Connection.zCommand.RecordCount > Existir.RecordCount then
            begin
              Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
                '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
              if Resp = mrCancel then
                raise Exception.Create('Proceso Cancelado por el Usuario.');

              if Resp = mrYes then
              begin
                connection.zCommand.First;
                while not connection.zCommand.Eof do
                begin
                  if not Existir.Locate('sIdBasico', connection.zCommand.FieldByName('sIdBasico').AsString, []) then
                    Connection.zCommand.Delete;
                  connection.zCommand.Next;
                end;
              end;
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ANEXO DE HERRAMIENTAS'}
        //ANEXO DE HERRAMIENTAS..
        if rAnexoHerr.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosBasicos('Herramienta') then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          Fila := 2;
          ProgressBar1.Max := 0;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          //if MessageDlg('Desea remplazar el catalogo de Herramientas existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            if ValidaDeleteAnexosP('herramientas', 'sIdHerramientas', '', 'recursosherramientas') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

  //**********************************************************
            // Generar una lista de registros que deben existir
            Existir.Close;
            Existir.FieldDefs.Add('sContrato', ftString, 15);
            Existir.FieldDefs.Add('sIdHerramientas', ftString, 25);
            Existir.Open;
            Existir.EmptyTable;
  //**********************************************************

            while (sValue <> '') do
            begin
              CodErr1 := '';
              CodErr2 := '';

              if lContratoActual then
                ImpsContrato := global_contrato
              else
                ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

              if ImpsContrato <> global_contrato then
                raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

              ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              ImpmDescripcion := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
              ImpsMedida := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoMN := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
              ImpdCostoDLL := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaMN := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
              ImpdVentaDLL := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;

              // Inserto Datos a la Tabla .....
              CodErr1 := 'Importación de Plantilla de Anexo de Herramientas';
              CodErr2 := 'Al tratar de insertar registros en la tabla herramientas';

              {Se insertan los datos de personal..}
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('INSERT INTO herramientas ( sContrato, sIdHerramientas, sDescripcion, sMedida, dCostoMN, dCostoDLL, dVentaMN, dVentaDLL, sSimbolo, fFecha) ' +
                ' VALUES (:contrato, :Herramienta, :Descripcion, :Medida, :CostoMN, :CostoDLL, :VentaMN, :VentaDLL, "", :Fecha)');

              try
                connection.zCommand.Active := False;
                Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                Connection.zCommand.Params.ParamByName('Herramienta').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Herramienta').value := ImpsNumeroActividad;
                Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                Connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
                Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                Connection.zCommand.Params.ParamByName('Medida').value := ImpsMedida;
                Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoMN').value := ImpdCostoMN;
                Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('CostoDLL').value := ImpdCostoDLL;
                Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaMN').value := ImpdVentaMN;
                Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                Connection.zCommand.Params.ParamByName('VentaDLL').value := ImpdVentaDLL;
                Connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
                Connection.zCommand.Params.ParamByName('Fecha').value := Date;
                Connection.zCommand.ExecSQL;
              except
                on e: exception do
                begin
                  // Verificar si se encontró una duplicidad de registros
                  if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                  begin
                    // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro
                    if not SobreTodos then
                      Resp := MessageDlg('El registro de herramientas identificado en EXCEL ya existe en la base de datos:' + #10 +
                        ImpsContrato + ' - ' + ImpsNumeroActividad + ' - ' + Trim(ImpmDescripcion) + #10 + #10 +
                        '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                    if Resp = mrYesToAll then
                      SobreTodos := True;

                    if (Resp = mrYes) or SobreTodos then
                      Resp := mrYes;

                    if Resp = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el Usuario.');

                    if Resp = mrYes then
                    begin
                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('UPDATE herramientas SET sDescripcion = :Descripcion, sMedida = :Medida, dCostoMN = :CostoMN, ' +
                        'dCostoDLL = :CostoDLL, dVentaMN = :VentaMN, dVentaDLL = :VentaDLL, sSimbolo = :Simbolo, fFecha = :Fecha ' +
                        'WHERE sContrato = :Contrato and sIdHerramientas = :Herramienta');
                      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Contrato').Value := ImpsContrato;
                      Connection.zCommand.Params.ParamByName('Herramienta').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Herramienta').Value := ImpsNumeroActividad;
                      Connection.zCommand.Params.ParamByName('Descripcion').DataType := ftMemo;
                      Connection.zCommand.Params.ParamByName('Descripcion').Value := Trim(ImpmDescripcion);
                      Connection.zCommand.Params.ParamByName('Medida').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Medida').Value := ImpsMedida;
                      Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('CostoMN').Value := ImpdCostoMN;
                      Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('CostoDLL').Value := ImpdCostoDLL;
                      Connection.zCommand.Params.ParamByName('VentaMN').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('VentaMN').Value := ImpdVentaMN;
                      Connection.zCommand.Params.ParamByName('VentaDLL').DataType := ftFloat;
                      Connection.zCommand.Params.ParamByName('VentaDLL').Value := ImpdVentaDLL;
                      Connection.zCommand.Params.ParamByName('Simbolo').DataType := ftString;
                      Connection.zCommand.Params.ParamByName('Simbolo').Value := '';
                      Connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
                      Connection.zCommand.Params.ParamByName('fecha').Value := Date;
                      Connection.zCommand.ExecSQL;
                    end;
                  end
                  else
                    raise;
                end;
              end;

              // Cargar la lista de registros procesados
              Existir.Append;
              Existir.FieldByName('sContrato').AsString := ImpsContrato;
              Existir.FieldByName('sIdHerramientas').AsString := ImpsNumeroActividad;
              Existir.Post;

              ProgressBar1.Max := ProgressBar1.Max + 1;
              ProgressBar1.Position := ProgressBar1.Position + 1;
              Fila := Fila + 1;
              sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            end;

  //**********************************************************
            // Verificar los registros que debería ser eliminados
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Select * from herramientas where sContrato = :Contrato');
            Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
            Connection.zCommand.Open;

            if Connection.zCommand.RecordCount > Existir.RecordCount then
            begin
              Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
                '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
              if Resp = mrCancel then
                raise Exception.Create('Proceso Cancelado por el Usuario.');

              if Resp = mrYes then
              begin
                connection.zCommand.First;
                while not connection.zCommand.Eof do
                begin
                  if not Existir.Locate('sIdHerramientas', connection.zCommand.FieldByName('sIdHerramientas').AsString, []) then
                    Connection.zCommand.Delete;
                  connection.zCommand.Next;
                end;
              end;
            end;
          end
        end;
  {$ENDREGION}
  {$REGION 'ALCANCES X PARTIDA'}
        //IMPORTACION DE LOS ALCANCES X PARTIDA...
        if rbAlcances.Checked then
        begin
          CodErr1 := '';
          CodErr1 := '';

          Fila := 1;
          iColumna := 0;
          sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
          while sValue <> '' do
          begin
            sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
            if (sValue <> '') then
              Inc(iColumna);

            Fila := Fila + 1;
          end;

          if iColumna <> 12 then
            raise Exception.Create('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar los Alcances x Partida.');

          fila := 2;
          ProgressBar1.Max := 0;
          sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          {Validacione de Contrato..}
          Fila    := 2;
          MiValor := global_contrato;
          sValue  := global_contrato;
          lEncuentra := False;

          EliminaCuadro('N', 0);

          {Si se elige contrato de Excel o contrato actual..}
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          while MiValor <> '' do
          begin
            MiValor := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

            {Si el contrro es Diferente a Caracter en Blanco..}
            if MiValor <> '' then
            begin
                if (sValue <> MiValor) then
                  ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

                {Validaciones de Campos..}
                DatoE := 'Texto';
                cadena := 'B';
                ValidaCampo(DatoE, cadena, Fila, 'Actividad', false, '');
                ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;

                DatoE := 'Numero';
                cadena := 'C';
                ValidaCampo(DatoE, cadena, Fila, 'iNivel', false, '');

                DatoE := 'Texto';
                cadena := 'D';
                ValidaCampo(DatoE, cadena, Fila, 'SubActividad', false, '');

                cadena := 'E';
                ValidaCampo(DatoE, cadena, Fila, 'Sistema', false, '');

                cadena := 'F';
                ValidaCampo(DatoE, cadena, Fila, 'Descripcion', false, '');

                cadena := 'G';
                ValidaCampo(DatoE, cadena, Fila, 'Medida', false, '');

                DatoE := 'Decimal';
                cadena := 'H';
                ValidaCampo(DatoE, cadena, Fila, 'Porcentaje', false, '');

                DatoE := 'Fecha';
                cadena := 'I';
                ValidaCampo(DatoE, cadena, Fila, 'Fecha Inicio', false, '');

                cadena := 'J';
                ValidaCampo(DatoE, cadena, Fila, 'Fecha Final', true, 'I');

                DatoE := 'Decimal';
                cadena := 'K';
                ValidaCampo(DatoE, cadena, Fila, 'Precio MN', false, '');

                cadena := 'L';
                ValidaCampo(DatoE, cadena, Fila, 'Precio DLL', false, '');

                {Verificamos si existe las partidas en el anexo C}
                connection.QryBusca.Active := False;
                Connection.QryBusca.SQL.Clear;
                connection.QryBusca.SQL.Add('select * from actividadesxanexo Where sContrato = :contrato and sIdConvenio =:Convenio and sNumeroActividad =:Actividad ');
                Connection.QryBusca.Params.ParamByName('Contrato').DataType  := ftString;
                Connection.QryBusca.Params.ParamByName('Contrato').Value     := Global_Contrato;
                Connection.QryBusca.Params.ParamByName('Convenio').DataType  := ftString;
                Connection.QryBusca.Params.ParamByName('Convenio').Value     := Global_Convenio;
                Connection.QryBusca.Params.ParamByName('Actividad').DataType := ftString;
                Connection.QryBusca.Params.ParamByName('Actividad').Value    := ImpsNumeroActividad;
                connection.QryBusca.Open;

                {Registramos partidas no permitidas..}
                if connection.QryBusca.RecordCount = 0 then
                  ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 2, 'Actividad');
            end;
            Fila := Fila + 1;
          end;


          {Validaciones  de Alcances Permitirdos..}
          Fila := 2;
          sValue := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
          dmpiAvance := 0;
          MiValor := 'Iniciando..';
          cadena := '';
          DatoE := '';
          while MiValor <> '' do
          begin
            MiValor := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            try
              if (sValue = MiValor) then
              begin
                ImpfValor := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
                dmpiAvance := dmpiAvance + ImpfValor;
              end
              else
              begin
                ImpfValor := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
                if dmpiAvance > 0 then
                begin
                    if dmpiAvance <> 100 then
                    begin
                      sValue := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila - 1)), 'B' + Trim(IntToStr(Fila - 1))].Value2;
                      ColoresErrorExcel('H' + Trim(IntToStr(Fila - 1)), 'H' + Trim(IntToStr(Fila - 1)), 3, 'Alcances');
                      dmpiAvance := ImpfValor;
                    end
                end
                else
                  dmpiAvance := ImpfValor;
                sValue := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
              end;
            except
            end;
            Fila := Fila + 1;
          end;

          CuadroColores('N', 'O', 'P', 'Q');

          {Generamos cuadro de codigo de colores}
          if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
          begin
            ExcelApplication1.UserControl := False;
            ExcelApplication1.Interactive[flcid] := True;
            ExcelApplication1.Disconnect;
            messageDLG('Proceso Cancelado por el Sistema', mtInformation, [mbOk], 0);
            exit;
          end;

          {Temrina Validacion contrato..}
          if lContratoActual then
            sValue := global_contrato;

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          {Validaciones  de Alcances Permitirdos..}
          Fila := 2;
          sValue := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
          dmpiAvance := 0;
          MiValor := 'Iniciando..';
          cadena := '';
          DatoE := '';
          while MiValor <> '' do
          begin
            MiValor := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            if (sValue = MiValor) then
            begin
              ImpfValor := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
              dmpiAvance := dmpiAvance + ImpfValor;
            end
            else
            begin
              ImpfValor := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
              if dmpiAvance > 0 then
                if dmpiAvance <> 100 then
                begin
                    sValue := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila - 1)), 'B' + Trim(IntToStr(Fila - 1))].Value2;
                    cadena := cadena + ' Part. ' + sValue + ',';
                    dmpiAvance := ImpfValor;
                end
                else
                  dmpiAvance := ImpfValor;
              sValue := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            end;
            Fila := Fila + 1;
          end;

          if cadena <> '' then
            raise Exception.Create('Existen Partidas con Alcances Fuera de los Rangos permitidos (100%), Favor de Verificar!' + #13 + cadena);

          fila := 2;
          ProgressBar1.Max := 0;
          sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
          if MessageDlg('Desea remplazar Los Alcances existente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
              //Primeor revisamos si existe infrmacion en el anexo..
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('select iFase from alcancesxactividad Where sContrato = :contrato and sIdConvenio =:Convenio limit 1');
              Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Contrato').Value    := Global_Contrato;
              Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              Connection.zCommand.Params.ParamByName('Convenio').Value    := Global_Convenio;
              connection.zCommand.open;

              if connection.zCommand.RecordCount = 0 then
              begin
                  // Se elimina el catalogo de ALCANCES
                  try
                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('DELETE FROM alcancesxactividad Where sContrato = :contrato and sIdConvenio =:Convenio ');
                    Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Contrato').Value    := Global_Contrato;
                    Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Convenio').Value    := Global_Convenio;
                    connection.zCommand.ExecSQL();
                  except
                    on e: exception do
                    begin
                      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importacion de Plantillas', 'Al Eliminar registro', 0);
                      exit;
                    end;
                  end;
              end
              else
              begin
                  messageDLG('Existen SubActividades cargadas al Anexo, No se pueden reemplazar!', mtInformation, [mbOk], 0);
                  raise Exception.Create('Proceso Cancelado por el Sistema');
              end;
          end;

          lEncuentra := True;
          if lContratoActual then
            ImpsContrato := global_contrato
          else
            ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          try
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('insert into alcancesxactividad (sContrato, sIdConvenio, sWbs, sNumeroActividad, iFase, iFaseAnterior, sNumeroActividadSub, sSistema, iNivel, sSimbolo, sTipoActividad, sDescripcion, sMedida, dCantidad, dAvance, ' +
                                          'dPonderado, dDuracion, dFechaInicio, dFechaFinal, dVentaMN, dVentaDLL) values (:contrato, :convenio, :wbs, :actividad, :fase, :faseantes, :actividad_sub, :sistema, :nivel, :simbolo, :tipoactividad, :descripcion, '+
                                          ':medida, :cantidad, :cantidad, :ponderado, :duracion, :fechaI, :fechaF, :ventaMN, :ventaDLL )');
              CodErr1 := 'Importacion de Plantillas de Alcances por partida';
              CodErr2 := 'Al insertar registros en la tabla alcancesxactividad';

              ImpsNumeroActividad := '';
              impiFaseAnterior    := 0;
              while (sValue <> '') do
              begin
                  ImpsPartida := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
                  if ImpsNumeroActividad <> ImpsPartida then
                  begin
                     connection.QryBusca.Active := False;
                     Connection.QryBusca.SQL.Clear;
                     connection.QryBusca.SQL.Add('select sWbs from actividadesxanexo Where sContrato = :contrato and sIdConvenio =:Convenio and sNumeroActividad =:Actividad ');
                     Connection.QryBusca.Params.ParamByName('Contrato').DataType  := ftString;
                     Connection.QryBusca.Params.ParamByName('Contrato').Value     := Global_Contrato;
                     Connection.QryBusca.Params.ParamByName('Convenio').DataType  := ftString;
                     Connection.QryBusca.Params.ParamByName('Convenio').Value     := Global_Convenio;
                     Connection.QryBusca.Params.ParamByName('Actividad').DataType := ftString;
                     Connection.QryBusca.Params.ParamByName('Actividad').Value    := ImpsPartida;
                     connection.QryBusca.Open;

                     {Consultamos las Wbs.}
                     if connection.QryBusca.RecordCount > 0 then
                        ImpsWbs := connection.QryBusca.FieldValues['sWbs'];

                     //Aqui recorremos los datos para calcular los ponderados de las subactividades,
                     i := Fila;
                     ImpdMontoPonderado := 0;
                     ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
                     while (sValue <> '') do
                     begin
                         ImpsEspecificacion := ExcelWorksheet1.Range['B' + Trim(IntToStr(i)), 'B' + Trim(IntToStr(i))].Value2;
                         if ImpsNumeroActividad = ImpsEspecificacion then
                         begin
                             ImpsMedida   := ExcelWorksheet1.Range['G' + Trim(IntToStr(i)), 'G' + Trim(IntToStr(i))].Value2;
                             ImpdVentaMN  := ExcelWorksheet1.Range['K' + Trim(IntToStr(i)), 'K' + Trim(IntToStr(i))].Value2;

                             if trim(ImpsMedida) <> '' then
                                ImpdMontoPonderado := ImpdMontoPonderado + StrToFloat(ImpdVentaMN);
                             inc(i);
                             sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(i)), 'A' + Trim(IntToStr(i))].Value2;
                         end
                         else
                            sValue := '';
                     end;
                     sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(fila)), 'A' + Trim(IntToStr(fila))].Value2;
                     impiFaseAnterior := 1;
                     impiFase := 0;
                  end;

                  ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
                  ImpsEspecificacion  := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
                  iNivel              := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
                  ImpmDescripcion     := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
                  ImpsMedida          := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
                  ImpdPonderado       := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
                  ImpdVentaMN         := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;
                  ImpdVentaDLL        := ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;
                  ImpdFechaInicio     := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;
                  ImpdFechaFinal      := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
                  ImpsSistema         := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
                  inc(impiFase);

                  // Inserto Datos a la Tabla .....
                  connection.zCommand.Active := False;
                  connection.zcommand.Params.ParamByName('Contrato').DataType      := ftString;
                  connection.zcommand.Params.ParamByName('Contrato').Value         := global_contrato;
                  connection.zcommand.Params.ParamByName('Convenio').DataType      := ftString;
                  connection.zcommand.Params.ParamByName('Convenio').Value         := global_convenio;
                  connection.zcommand.Params.ParamByName('Actividad').DataType     := ftString;
                  connection.zcommand.Params.ParamByName('Actividad').Value        := ImpsNumeroActividad;
                  connection.zcommand.Params.ParamByName('wbs').DataType           := ftString;
                  connection.zcommand.Params.ParamByName('wbs').Value              := ImpsWbs;
                  connection.zcommand.Params.ParamByName('fase').DataType          := ftInteger;
                  connection.zcommand.Params.ParamByName('fase').Value             := impiFase;
                  connection.zcommand.Params.ParamByName('faseantes').DataType     := ftInteger;
                  if trim(ImpsMedida) = '' then
                     connection.zcommand.Params.ParamByName('faseantes').Value     := 0
                  else
                     connection.zcommand.Params.ParamByName('faseantes').Value     := impiFaseAnterior;
                  connection.zcommand.Params.ParamByName('actividad_sub').DataType := ftString;
                  connection.zcommand.Params.ParamByName('actividad_sub').Value    := ImpsEspecificacion;
                  connection.zcommand.Params.ParamByName('sistema').DataType       := ftString;
                  connection.zcommand.Params.ParamByName('sistema').Value          := ImpsSistema;
                  connection.zcommand.Params.ParamByName('nivel').DataType         := ftInteger;
                  connection.zcommand.Params.ParamByName('nivel').Value            := iNivel;
                  connection.zcommand.Params.ParamByName('simbolo').DataType       := ftString;
                  if trim(ImpsMedida) = '' then
                     connection.zcommand.Params.ParamByName('simbolo').Value       := '-'
                  else
                     connection.zcommand.Params.ParamByName('simbolo').Value       := '';
                  connection.zcommand.Params.ParamByName('tipoactividad').DataType := ftString;
                  if trim(ImpsMedida) = '' then
                     connection.zcommand.Params.ParamByName('tipoactividad').Value := 'Paquete'
                  else
                     connection.zcommand.Params.ParamByName('tipoactividad').Value := 'Actividad';
                  connection.zcommand.Params.ParamByName('descripcion').DataType   := ftString;
                  connection.zcommand.Params.ParamByName('descripcion').Value      := ImpmDescripcion;
                  connection.zcommand.Params.ParamByName('medida').DataType        := ftString;
                  connection.zcommand.Params.ParamByName('medida').Value           := ImpsMedida;
                  connection.zcommand.Params.ParamByName('cantidad').DataType      := ftFloat;
                  connection.zcommand.Params.ParamByName('cantidad').Value         := (StrToFloat(ImpdVentaMN)/ImpdMontoPonderado)*100;
                  connection.zcommand.Params.ParamByName('duracion').DataType      := ftInteger;
                  connection.zcommand.Params.ParamByName('duracion').Value         := 0;
                  connection.zcommand.Params.ParamByName('ponderado').DataType     := ftFloat;
                  connection.zcommand.Params.ParamByName('ponderado').Value        := (StrToFloat(ImpdVentaMN)/ImpdMontoPonderado)*100;
                  connection.zcommand.Params.ParamByName('ventaMN').DataType       := ftFloat;
                  connection.zcommand.Params.ParamByName('ventaMN').Value          := ImpdVentaMN;
                  connection.zcommand.Params.ParamByName('ventaDLL').DataType      := ftFloat;
                  connection.zcommand.Params.ParamByName('ventaDLL').Value         := ImpdVentaDLL;
                  connection.zcommand.Params.ParamByName('fechaI').DataType        := ftDate;
                  connection.zcommand.Params.ParamByName('fechaI').Value           := ImpdFechaInicio;
                  connection.zcommand.Params.ParamByName('fechaF').DataType        := ftDate;
                  connection.zcommand.Params.ParamByName('fechaF').Value           := ImpdFechaFinal;
                  connection.zCommand.ExecSQL;
                  ProgressBar1.Max := ProgressBar1.Max + 1;
                  ProgressBar1.Position := ProgressBar1.Position + 1;

                  Fila := Fila + 1;
                  sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
              end;
          Except
          end;
        end;
  {$ENDREGION}
  {$REGION 'Carga de Insumos'}
        //carga de Insumos
        if rbInsumos.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if ValidaAnexosInsumo() then
            raise Exception.Create('Proceso Cancelado por el Sistema');

          //************************RBRITO 09/06/11*******************************
          //Definir el almacen al cual se van asociar los insumos
          connection.QryBusca.Active := False;
          connection.QryBusca.Filtered := False;
          connection.QryBusca.SQL.Clear;
          //buscar el primer almacen
          connection.QryBusca.SQL.Add('SELECT sIdAlmacen FROM almacenes LIMIT 1');
          connection.QryBusca.Open;
          if connection.QryBusca.RecordCount > 0 then
          begin
            sIdAlmacen := connection.QryBusca.FieldByName('sIdAlmacen').AsString;
          end
          else
          begin
            //dar de alta un almacen
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('INSERT INTO almacenes (sIdAlmacen, sDescripcion) VALUES ("ALM-01", "ALMACEN PRINCIPAL")');
            connection.zCommand.ExecSQL;
            sIdAlmacen := 'ALM-01';
          end;
          //************************RBRITO 09/06/11*******************************

          fila := 2;
          ProgressBar1.Max := 0;
          sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          if MessageDlg('¿Desea Eliminar los Materiales Existentes?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
             if ValidaDeleteAnexosP('insumos', 'sIdInsumo', 'bitacorademateriales', 'recursosanexo') then
                raise Exception.Create('Proceso Cancelado por el Sistema');

          ImpsContrato := global_contrato;

          CodErr1 := 'Importación de Plantilla de Insumos';
          CodErr2 := 'Al tratar de insertar registros en tabla insumos';

  //**********************************************************
            // Generar una lista de registros que deben existir
          Existir.Close;
          Existir.FieldDefs.Add('sContrato', ftString, 15);
          Existir.FieldDefs.Add('sIdInsumo', ftString, 25);
          Existir.FieldDefs.Add('sIdAlmacen', ftString, 20);
          Existir.Open;
          Existir.EmptyTable;
  //**********************************************************

          while (sValue <> '') do
          begin
            ImpsNumeroActividad := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            ImpsTipo := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            ImpmDescripcion := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
            ImpsMedida := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
            ImpdCantidadAnexo := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
            ImpdInstalado := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
            ImpdFechaInicio := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
            ImpdCostoMN := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
            ImpdCostoDLL := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;
            ImpdVentaMN := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
            ImpdVentaDLL := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;

              // Inserto Datos a la Tabla .....
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('INSERT INTO insumos ( sContrato, sIdInsumo, sIdProveedor, sIdAlmacen, sTipoActividad, mDescripcion, dFechaInicio, dCostoMN, dCostoDLL, dVentaMN, dVentaDLL, sMedida, dCantidad, dInstalado, sIdGrupo, dNuevoPrecio, sIdFase) ' +
              ' VALUES (:contrato, :insumo, null, :almacen, :tipoactividad, :Descripcion, :fechai, :costoMN, :costoDLL, :ventaMN, :ventaDLL, :medida, :cantidad, :instalado, null, 0, null)');
            try
              connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
              connection.zCommand.Params.ParamByName('insumo').DataType := ftString;
              connection.zCommand.Params.ParamByName('insumo').value := ImpsNumeroActividad;
              connection.zCommand.Params.ParamByName('almacen').DataType := ftString;
              connection.zCommand.Params.ParamByName('almacen').value := sIdAlmacen;
              connection.zCommand.Params.ParamByName('tipoactividad').DataType := ftString;
              connection.zCommand.Params.ParamByName('tipoactividad').value := ImpsTipo;
              connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
              connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
              connection.zCommand.Params.ParamByName('fechai').DataType := ftDate;
              connection.zCommand.Params.ParamByName('fechai').value := ImpdFechaInicio;
              connection.zCommand.Params.ParamByName('costoMN').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('costoMN').value := ImpdCostoMN;
              connection.zCommand.Params.ParamByName('costoDLL').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('costoDLL').value := ImpdCostoDLL;
              connection.zCommand.Params.ParamByName('ventaMN').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('ventaMN').value := ImpdVentaMN;
              connection.zCommand.Params.ParamByName('ventaDLL').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('ventaDLL').value := ImpdVentaDLL;
              connection.zCommand.Params.ParamByName('medida').DataType := ftString;
              connection.zCommand.Params.ParamByName('medida').value := ImpsMedida;
              connection.zCommand.Params.ParamByName('cantidad').DataType := ftInteger;
              connection.zCommand.Params.ParamByName('cantidad').value := ImpdCantidadAnexo;
              connection.zCommand.Params.ParamByName('instalado').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('instalado').value := ImpdInstalado;
              connection.zCommand.ExecSQL;
            except
              on e: exception do
              begin
                  // Verificar si se encontró una duplicidad de registros
                if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                begin
                    // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro
                  if not SobreTodos then
                    Resp := MessageDlg('El registro de insumos identificado en EXCEL ya existe en la base de datos:' + #10 +
                      ImpsContrato + ' - ' + ImpsNumeroActividad + #10 + #10 +
                      '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                  if Resp = mrYesToAll then
                    SobreTodos := True;

                  if (Resp = mrYes) or SobreTodos then
                    Resp := mrYes;

                  if Resp = mrCancel then
                    raise Exception.Create('Proceso Cancelado por el Usuario.');

                  if Resp = mrYes then
                  begin
                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('UPDATE insumos SET sIdProveedor = null, sTipoActividad = :tipoactividad, mDescripcion = :Descripcion, ' +
                      'dFecha = :Fechai, dCostoMN = :CostoMN, dCostoDLL = :CostoDLL, dVentaMN = :VentaMN, dVentaDLL = :VentaDLL, ' +
                      'sMedida = :medida, dCantidad = :Cantidad, dInstalado = :Instalado ' +
                      'WHERE sContrato = :Contrato and sIdInsumo = :Insumo and sIdAlmacen = :Almacen');
                    connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                    connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
                    connection.zCommand.Params.ParamByName('insumo').DataType := ftString;
                    connection.zCommand.Params.ParamByName('insumo').value := ImpsNumeroActividad;
                    connection.zCommand.Params.ParamByName('almacen').DataType := ftString;
                    connection.zCommand.Params.ParamByName('almacen').value := sIdAlmacen;
                    connection.zCommand.Params.ParamByName('tipoactividad').DataType := ftString;
                    connection.zCommand.Params.ParamByName('tipoactividad').value := ImpsTipo;
                    connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
                    connection.zCommand.Params.ParamByName('Descripcion').value := Trim(ImpmDescripcion);
                    connection.zCommand.Params.ParamByName('fechai').DataType := ftDate;
                    connection.zCommand.Params.ParamByName('fechai').value := ImpdFechaInicio;
                    connection.zCommand.Params.ParamByName('costoMN').DataType := ftFloat;
                    connection.zCommand.Params.ParamByName('costoMN').value := ImpdCostoMN;
                    connection.zCommand.Params.ParamByName('costoDLL').DataType := ftFloat;
                    connection.zCommand.Params.ParamByName('costoDLL').value := ImpdCostoDLL;
                    connection.zCommand.Params.ParamByName('ventaMN').DataType := ftFloat;
                    connection.zCommand.Params.ParamByName('ventaMN').value := ImpdVentaMN;
                    connection.zCommand.Params.ParamByName('ventaDLL').DataType := ftFloat;
                    connection.zCommand.Params.ParamByName('ventaDLL').value := ImpdVentaDLL;
                    connection.zCommand.Params.ParamByName('medida').DataType := ftString;
                    connection.zCommand.Params.ParamByName('medida').value := ImpsMedida;
                    connection.zCommand.Params.ParamByName('cantidad').DataType := ftInteger;
                    connection.zCommand.Params.ParamByName('cantidad').value := ImpdCantidadAnexo;
                    connection.zCommand.Params.ParamByName('instalado').DataType := ftFloat;
                    connection.zCommand.Params.ParamByName('instalado').value := ImpdInstalado;
                    Connection.zCommand.ExecSQL;
                  end;
                end
                else
                  raise;
              end;
            end;

              // Cargar la lista de registros procesados
            Existir.Append;
            Existir.FieldByName('sContrato').AsString := ImpsContrato;
            Existir.FieldByName('sIdInsumo').AsString := ImpsNumeroActividad;
            Existir.FieldByName('sIdAlmacen').AsString := '';
            Existir.Post;

            ProgressBar1.Max := ProgressBar1.Max + 1;
            ProgressBar1.Position := ProgressBar1.Position + 1;
            Fila := Fila + 1;
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
          end;

          try
            Kardex('Importacion de Datos', 'Materiales', 'Frente de Trabajo', '', '', '', '','Tarifa Diaria','Importcion de Datos');
          except
            on e: exception do
            begin
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al registrar en kardex Carga de Insumos', 0);
            end;
          end;

  //**********************************************************
            // Verificar los registros que debería ser eliminados
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('Select * from insumos where sContrato = :Contrato and sIdAlmacen = :almacen');
          Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
          Connection.zCommand.ParamByName('almacen').AsString := sIdAlmacen;
          Connection.zCommand.Open;

          if Connection.zCommand.RecordCount > Existir.RecordCount then
          begin
            Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
              '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
            if Resp = mrCancel then
              raise Exception.Create('Proceso Cancelado por el Usuario.');

            if Resp = mrYes then
            begin
              connection.zCommand.First;
              while not connection.zCommand.Eof do
              begin
                if not Existir.Locate('sIdInsumo', connection.zCommand.FieldByName('sIdInsumo').AsString, []) then
                  Connection.zCommand.Delete;
                connection.zCommand.Next;
              end;
            end;
          end;
        end;
  {$ENDREGION}
  {$REGION 'Insumos x Partida'}
        //Carga de Personal x Partida, Equipo x Partida, Basicos x Partida, Herramientas x Partida, Material x Partida..
        if (rbPersonalxP.Checked) or (rbEquipoxP.Checked) or (rbInsumosxP.Checked) or (rHerrxPartida.Checked) or (rBasicosxPart.Checked) then
        begin
          CodErr1 := '';
          CodErr2 := '';

          if rbPersonalxP.Checked then
          begin
            if ValidaAnexosPERxP('Personal', 'personal', 'sIdpersonal') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'recursospersonalnuevos';
            ImpsWbs := 'sIdPersonal';
            ImpsAnexo := 'Personal x Partida';
            DatoE := 'personal';
            cadena := 'p.dCantidad*0';
          end;

          if rbEquipoxP.Checked then
          begin
            if ValidaAnexosPERxP('Equipo', 'equipos', 'sIdEquipo') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'recursosequiposnuevos';
            ImpsWbs := 'sIdEquipo';
            ImpsAnexo := 'Equipo x Partida';
            DatoE := 'equipos';
            cadena := 'p.dCantidad*0';
          end;

          if rbInsumosxP.Checked then
          begin
            if ValidaAnexosPERxP('Insumo', 'insumos', 'sIdInsumo') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'recursosanexosnuevos';
            ImpsWbs := 'sIdInsumo';
            ImpsAnexo := 'Insumos x Partida';
            DatoE := 'insumos';
            cadena := 'p.dCantidad*0';
          end;

          if rHerrxPartida.Checked then
          begin
            if ValidaAnexosPERxP('Herramienta', 'herramientas', 'sIdHerramientas') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'recursosherramientasnuevos';
            ImpsWbs := 'sIdHerramientas';
            ImpsAnexo := 'Herramienta x Partida';
            DatoE := 'herramientas';
            cadena := '';
          end;

          if rBasicosxPart.Checked then
          begin
            if ValidaAnexosPERxP('Basico', 'basicos', 'sIdBasico') then
              raise Exception.Create('Proceso Cancelado por el Sistema');

            sTabla := 'recursosbasicosnuevos';
            ImpsWbs := 'sIdBasico';
            ImpsAnexo := 'Basico x Partida';
            DatoE := 'basicos';
            cadena := '';
          end;

          fila := 2;
          cadena := '';
          ProgressBar1.Max := 0;
          if lContratoActual then
            sValue := global_contrato
          else
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

          if sValue <> global_contrato then
            raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

          {if MessageDlg('¿Desea remplazar el '+ImpsAnexo+' existente?', mtConfirmation, [mbYes, mbNo], 0) = mrnO then
            Raise Exception.Create('Proceso Cancelado por el Usuario');

          // Se elimina el catalogo de Insumos
          CodErr1 := 'Importación de Plantilla de Insumos por Partida';
          CodErr2 := 'Al Eliminar registros en la tabla ' + sTabla;

          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('DELETE FROM '+sTabla+' Where sContrato = :contrato');
          Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
          connection.zCommand.ExecSQL();

          CodErr1 := '';
          CodErr2 := '';}

  //**********************************************************
          // Generar una lista de registros que deben existir
          Existir.Close;
          Existir.FieldDefs.Add('sContrato', ftString, 15);
          Existir.FieldDefs.Add('sWbs', ftString, 30);
          Existir.FieldDefs.Add('sNumeroActividad', ftString, 20);
          Existir.FieldDefs.Add(ImpsWbs, ftString, 25);
          Existir.Open;
          Existir.EmptyTable;
  //**********************************************************
          Cadena2 := '';
          while (sValue <> '') do
          begin
            CodErr1 := '';
            CodErr2 := '';

            if lContratoActual then
              ImpsContrato := global_contrato
            else
              ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

            if ImpsContrato <> global_contrato then
              raise Exception.Create('El archivo que desea importar pertenece a otro contrato');

            ImpsAnexo := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            ImpsNumeroActividad := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
            sIdRecurso := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
            ImpdCantidadAnexo := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;

            // Consultamos datos del anexos C
            connection.QryBusca2.Active := False;
            connection.QryBusca2.SQL.Clear;
            connection.QryBusca2.SQL.Add('select sWbs, dCostoMN, dCostoDLL from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroActividad =:Actividad ' +
              'and sTipoActividad = "Actividad" and sAnexo =:Anexo ');
            Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
            Connection.QryBusca2.Params.ParamByName('Contrato').Value := Global_Contrato;
            Connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
            Connection.QryBusca2.Params.ParamByName('Convenio').Value := Global_Convenio;
            Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString;
            Connection.QryBusca2.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
            Connection.QryBusca2.Params.ParamByName('Anexo').DataType := ftString;
            Connection.QryBusca2.Params.ParamByName('Anexo').Value := ImpsAnexo;
            connection.QryBusca2.Open;

            // Inserto Datos a la Tabla .....
            CodErr1 := 'Importación de Plantilla de Insumos por Partida';
            CodErr2 := 'Al tratar de insertar registros en la tabla ' + sTabla;

            try
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('insert into ' + sTabla + ' (sContrato,sWbs,sNumeroActividad,' + ImpsWbs + ',dCantidad, dCostoMN, dCostoDLL) ' +
                'values (:Contrato, :wbs, :Actividad, :Id, :cantidad, :CostoMN, :CostoDLL)');

              connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
              connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
              connection.zCommand.Params.ParamByName('Wbs').value := connection.QryBusca2.FieldValues['sWbs'];
              connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
              connection.zCommand.Params.ParamByName('actividad').value := ImpsNumeroActividad;
              connection.zCommand.Params.ParamByName('id').DataType := ftString;
              connection.zCommand.Params.ParamByName('id').value := sIdRecurso;
              connection.zCommand.Params.ParamByName('cantidad').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('cantidad').value := StrToFloat(ImpdCantidadAnexo);
              connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('CostoMN').value := connection.QryBusca2.FieldValues['dCostoMN'];
              connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('CostoDLL').value := connection.QryBusca2.FieldValues['dCostoDLL'];
              connection.zCommand.ExecSQL;
            except
              on e: exception do
              begin
                // Verificar si se encontró una duplicidad de registros
                if (CompareText(e.ClassName, 'EZSQLException') = 0) and (Pos('Duplicate entry', e.Message) > 0) then
                begin
                  // Si se trata de un registro duplicado entonces solo tratar de actualizar el registro
                  if not SobreTodos then
                    Resp := MessageDlg('El registro de ' + ImpsAnexo + ' identificado en EXCEL ya existe en la base de datos:' + #10 +
                      ImpsContrato + ' - ' + connection.QryBusca2.FieldValues['sWbs'] + ' - ' + ImpsNumeroActividad + ' - ' + sIdRecurso + #10 + #10 +
                      '¿Desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo, mbYesToAll, mbCancel], 0);

                  if Resp = mrYesToAll then
                    SobreTodos := True;

                  if (Resp = mrYes) or SobreTodos then
                    Resp := mrYes;

                  if Resp = mrCancel then
                    raise Exception.Create('Proceso Cancelado por el Usuario.');

                  if Resp = mrYes then
                  begin
                    connection.zCommand.Active := False;
                    connection.zCommand.SQL.Clear;
                    connection.zCommand.SQL.Add('UPDATE ' + sTabla + ' SET dCantidad = :Cantidad, dCostoMN = :CostoMN, dCostoDLL = :CostoDLL ' +
                      'WHERE sContrato = :Contrato and sWbs = :Wbs and sNumeroActividad = :Actividad and ' + ImpsWbs + ' = :Id');
                    Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Contrato').Value := ImpsContrato;
                    Connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Wbs').Value := connection.QryBusca2.FieldValues['sWbs'];
                    Connection.zCommand.Params.ParamByName('Actividad').DataType := ftMemo;
                    Connection.zCommand.Params.ParamByName('Actividad').Value := ImpsNumeroActividad;
                    Connection.zCommand.Params.ParamByName('Id').DataType := ftString;
                    Connection.zCommand.Params.ParamByName('Id').Value := sIdRecurso;
                    Connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('Cantidad').Value := StrToFloat(ImpdCantidadAnexo);
                    Connection.zCommand.Params.ParamByName('CostoMN').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('CostoMN').Value := connection.QryBusca2.FieldValues['dCostoMN'];
                    Connection.zCommand.Params.ParamByName('CostoDLL').DataType := ftFloat;
                    Connection.zCommand.Params.ParamByName('CostoDLL').Value := connection.QryBusca2.FieldValues['dCostoDLL'];
                    Connection.zCommand.ExecSQL;
                  end;
                end
                else
                  raise;
              end;
            end;

            CodErr1 := '';
            CodErr2 := '';

            {Ahora informamos al usuario que ids no se encontraron en cada catalogo,, y no se guardaron..}
            connection.QryBusca.Active := False;
            Connection.QryBusca.SQL.Clear;
            connection.QryBusca.SQL.Add('select ' + ImpsWbs + ' from ' + sTabla + ' Where sContrato = :contrato and ' + ImpsWbs + ' =:Id ');
            Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
            Connection.QryBusca.Params.ParamByName('Contrato').Value := Global_Contrato;
            Connection.QryBusca.Params.ParamByName('Id').DataType := ftString;
            Connection.QryBusca.Params.ParamByName('Id').Value := sIdRecurso;
            connection.QryBusca.Open;

            {Registramos los Ids no encontrados..}
            if connection.QryBusca.RecordCount = 0 then
              cadena2 := Cadena2 + sIdRecurso + ' , ';

            // Cargar la lista de registros procesados
            Existir.Append;
            Existir.FieldByName('sContrato').AsString := ImpsContrato;
            Existir.FieldByName('sWbs').AsString := connection.QryBusca2.FieldValues['sWbs'];
            Existir.FieldByName('sNumeroActividad').AsString := ImpsNumeroActividad;
            Existir.FieldByName(ImpsWbs).AsString := sIdRecurso;
            Existir.Post;

            ProgressBar1.Max := ProgressBar1.Max + 1;
            ProgressBar1.Position := ProgressBar1.Position + 1;
            Fila := Fila + 1;
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
          end;

          if Cadena2 <> '' then
            raise Exception.Create('No se encontraron los siguientes Ids en el Catalogo de ' + DatoE + '. Para que se puedan Guardar, favor de darlos de alta.');


  //**********************************************************
          // Verificar los registros que debería ser eliminados
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('Select * from ' + sTabla + ' where sContrato = :Contrato');
          Connection.zCommand.ParamByName('Contrato').AsString := ImpsContrato;
          Connection.zCommand.Open;

          if Connection.zCommand.RecordCount > Existir.RecordCount then
          begin
            Resp := MessageDlg('Existen ' + IntToStr(Connection.zCommand.RecordCount - Existir.RecordCount) + ' registros en la base de datos que no fueron obtenidos de la tabla de EXCEL.' + #10 + #10 +
              '¿Desea eliminar estos registros ahora?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
            if Resp = mrCancel then
              raise Exception.Create('Proceso Cancelado por el Usuario.');

            if Resp = mrYes then
            begin
              connection.zCommand.First;
              while not connection.zCommand.Eof do
              begin
                if not Existir.Locate('sWbs;sNumeroActividad;' + ImpsWbs, vararrayof([connection.zCommand.FieldByName('sWbs').AsString, connection.zCommand.FieldByName('sNumeroActividad').AsString, connection.zCommand.FieldByName(ImpsWbs).AsString]), []) then
                  Connection.zCommand.Delete;
                connection.zCommand.Next;
              end;
            end;
          end;
        end;
  {$ENDREGION}
  {$REGION 'AVANCE PROGRAMADO'}
        //IMPORTACION DE LOS AVANCES PROGRAMADOS ...
        //*******************************************************************************************
        if rAvances.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';
          sTmpOrden := ExcelWorksheet1.Name;

          if ValidaAvancesProgramados() then
            raise Exception.Create('Al parecer la validacion del archivo de excel ha fallado, verifique que el formato sea correcto.');

          Fila := 2;
          if lContratoActual then
            ImpsContrato := global_contrato
          else
            ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
          sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);
          ProgressBar1.Max := 0;
          with connection do
          begin 
            QryBusca.Active := false;
            QryBusca.SQL.Clear;
            QryBusca.SQl.Add('select sContrato,sNumeroOrden from avancesglobales where sContrato=:contrato and sNumeroOrden=:orden and sIdConvenio=:convenio');
            QryBusca.ParamByName('contrato').AsString := global_contrato;
            QryBusca.ParamByName('orden').AsString := sTmpOrden;
            QryBusca.ParamByName('convenio').AsString := global_convenio;
            QryBusca.Open;
              //if sTmpOrden = '' then CodErr2 := ' el contrato?';
              //if sTmpOrden <> '' then CodErr2 := ' el la orden ' + sTmpOrden + ' ?';
              if Messagedlg( Format( '¿Desea importar los avances: %s ?', [ sTmpOrden ] ), mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrCancel then
                raise Exception.Create('Proceso Cancelado por el usuario')
              else
              begin
                zCommand.Active := false;
                zCommand.SQL.Clear;
                zCommand.SQL.Add('delete from avancesglobales where sContrato=:contrato and sNumeroOrden=:orden and sIdConvenio=:convenio');
                zCommand.ParamByName('contrato').AsString := global_contrato;
                zCommand.ParamByName('orden').AsString := sTmpOrden;
                zCommand.ParamByName('convenio').AsString := global_convenio;
                zCommand.ExecSQL;
              end;
          end;

          while (sValue <> '') do
          begin
            //ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            ImpsIdConvenio := global_convenio; //ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            dFecha := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            ImpsNumeroOrden := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
            ImpdAvancePonderadoDia := StrToFloat(ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2);
            ImpdAvancePonderadoGlobal := StrToFloat(ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2);
            ImpdAvanceFinanciero := StrToFloat(ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2);
            ImpNumeroGerencial := StrToInt( ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2 );
            ImpDuracion := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Text;
            ImpHorario := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Text;

            if ImpsNumeroOrden = NULL then ImpsNumeroOrden := '';
            if ImpsNumeroOrden = 'Null' then ImpsNumeroOrden := '';

            if not ImpNumeroGerencial in [ 1, 2, 3 ] then
              raise Exception.Create( 'Numero Gerencial Invalido.' );

            try
              with connection do
              begin
                if sTmpOrden <> ImpsNumeroOrden then
                begin
                  sTmpOrden := ImpsNumeroOrden;
                end;
                zCommand.Active := False;
                zCommand.SQL.Clear;
                sSQL := ' INSERT INTO avancesglobales ( sContrato , sIdConvenio, dIdFecha, sNumeroOrden, ' +
                  ' iNumeroGerencial, dAvancePonderadoDia, dAvancePonderadoGlobal, dAvanceFinanciero, sDuracion, sHorario  ) ' +
                  ' VALUES (:contrato, :convenio, :fecha, :orden, :gerencial, :AvanceDia, :AvanceGlobal, :AvanceFinanciero, :duracion, :horario ) ' +
                  ' on duplicate key update dAvancePonderadoDia=:AvanceDia, dAvancePonderadoGlobal=:AvanceGlobal,dAvanceFinanciero=:AvanceFinanciero ';
                zCommand.SQL.Add(sSQL);
                zCommand.ParamByName('contrato').AsString := ImpsContrato;
                zCommand.ParamByName('convenio').AsString := ImpsIdConvenio;
                zCommand.ParamByName('fecha').AsString := FormatDateTime( 'YYYY-MM-DD', dFecha );
                zCommand.ParamByName('orden').AsString := ImpsNumeroOrden;
                zCommand.ParamByName('AvanceDia').AsFloat := ImpdAvancePonderadoDia;
                zCommand.ParamByName('AvanceGlobal').AsFloat := ImpdAvancePonderadoGlobal;
                zCommand.ParamByName('AvanceFinanciero').AsFloat := ImpdAvanceFinanciero;
                zCommand.ParamByName( 'gerencial' ).AsInteger := ImpNumeroGerencial;
                zCommand.ParamByName( 'duracion' ).AsString := ImpDuracion;
                zCommand.ParamByName( 'horario' ).AsString := ImpHorario;
                zCommand.ExecSQL;
              end;
            except
              on e: exception do
              begin
                MessageDlg('Error: ' + e.Message, mtError, [mbok], 0);
                ImpsContrato := '';
              end;
            end;
            ProgressBar1.Max := ProgressBar1.Max + 1;
            ProgressBar1.Position := ProgressBar1.Position + 1;
            fila := fila + 1;
            sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);
          end;


        end;
  {$ENDREGION}
  {$REGION 'AVANCE PROGRAMADO CEDULA'}
       
  {$ENDREGION}


  {$REGION 'ANEXO DT CIA'}
        //IMPORTACION DE LOS AVANCES PROGRAMADOS ...
        //*******************************************************************************************
        if rAnexoDTCia.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';
          sTmpOrden := '<//|\\>';

          if ValidaAnexoDTCia() then
            raise Exception.Create('Al parecer la validacion del archivo de excel ha fallado, verifique que el formato sea correcto.');

          Fila := 2;

          sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);

          ProgressBar1.Max := 0;
          while (sValue <> '') do
          begin
            if lContratoActual then
              ImpsContrato := global_contrato
            else
              ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
            ImpsIdConvenio := {global_convenio;} ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            dFecha := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
            Impswbs := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
            ImpsNumeroActividad := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
            ImpdCantidad := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;

            if (ImpsIdConvenio = Null) or (ImpsIdConvenio = NULL) or (ImpsIdConvenio = 'NULL') or (ImpsIdConvenio = 'Null') then
            begin
              ImpsIdConvenio := '';
            end;
            try
              with connection do
              begin
                if sTmpOrden <> ImpsContrato then
                begin
                  sTmpOrden := ImpsContrato;
                  QryBusca.Active := false;
                  QryBusca.SQL.Clear;
                  QryBusca.SQl.Add('select sContrato from distribuciondeanexocia where sContrato=:contrato and sIdConvenio=:convenio');
                  QryBusca.ParamByName('contrato').AsString := ImpsContrato;
                  QryBusca.ParamByName('convenio').AsString := global_convenio;
                  QryBusca.Open;
                  if QryBusca.RecordCount > 0 then
                  begin

                    BotonSelec := MessageDlg('¿Desea reemplazar el anexo DT de la Cia para contrato?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
                    if BotonSelec = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el usuario')
                    else
                    begin
                      zCommand.Active := false;
                      zCommand.SQL.Clear;
                      zCommand.SQL.Add('delete from distribuciondeanexocia where sContrato=:contrato and sIdConvenio=:convenio');
                      zCommand.ParamByName('contrato').AsString := ImpsContrato;
                      zCommand.ParamByName('convenio').AsString := global_convenio;
                      zCommand.ExecSQL;
                    end;
                  end;
                end;
                zCommand.Active := False;
                zCommand.SQL.Clear;
                sSQL := ' INSERT INTO distribuciondeanexocia ( sContrato , sIdConvenio, dIdFecha, sWbs, ' +
                  ' sNumeroActividad, dCantidad) ' +
                  ' VALUES (:contrato, :convenio, :fecha, :wbs, :actividad, :cantidad) ' +
                  ' on duplicate key update dCantidad=:cantidad ';
                zCommand.SQL.Add(sSQL);
                zCommand.ParamByName('contrato').AsString := ImpsContrato;
                zCommand.ParamByName('convenio').AsString := ImpsIdConvenio;
                zCommand.ParamByName('fecha').AsDate := dFecha;
                zCommand.ParamByName('wbs').AsString := ImpsWbs;
                zCommand.ParamByName('actividad').AsString := ImpsNumeroActividad;
                zCommand.ParamByName('cantidad').AsFloat := ImpdCantidad;
                zCommand.ExecSQL;
              end;
            except
              on e: exception do
              begin
                MessageDlg('Error: ' + e.Message, mtError, [mbok], 0);
                ImpsContrato := '';
              end;
            end;
            ProgressBar1.Max := ProgressBar1.Max + 1;
            ProgressBar1.Position := ProgressBar1.Position + 1;
            fila := fila + 1;
            sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);
          end;


        end;
  {$ENDREGION}
  {$REGION 'ANEXO DT ORDEN CIA'}
        //IMPORTACION DE LOS AVANCES PROGRAMADOS ...
        //*******************************************************************************************
        if rAnexoDTOrdenCia.Checked then
        begin
          CodErr1 := '';
          CodErr2 := '';
          sTmpOrden := '<//|\\>';

          if ValidaAnexoDTOrdenCia() then
            raise Exception.Create('Al parecer la validacion del archivo de excel ha fallado, verifique que el formato sea correcto.');

          Fila := 2;

          sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);

          ProgressBar1.Max := 0;
          while (sValue <> '') do
          begin
            if lContratoActual then
              ImpsContrato  := global_contrato
            else
              ImpsContrato  := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

            ImpsIdConvenio  :=  ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            ImpsNumeroOrden := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;          
            Impswbs := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
            ImpsNumeroActividad := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
            dFecha := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
            ImpdCantidad := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;

            if (ImpsIdConvenio = Null) or (ImpsIdConvenio = NULL) or (ImpsIdConvenio = 'NULL') or (ImpsIdConvenio = 'Null') then
            begin
              ImpsIdConvenio := '';
            end;
            try
              with connection do
              begin
                if (sTmpOrden <> ImpsNumeroOrden) or (sTmpContrato <> ImpsContrato) then
                begin
                  sTmpOrden    := ImpsNumeroOrden;
                  sTmpContrato := ImpsContrato;
                  QryBusca.Active := false;
                  QryBusca.SQL.Clear;
                  QryBusca.SQl.Add('select sContrato from distribuciondeactividadescia where sContrato=:contrato and sNumeroOrden = :orden and sIdConvenio=:convenio');
                  QryBusca.ParamByName('contrato').AsString := ImpsContrato;
                  QryBusca.ParamByName('convenio').AsString := global_convenio;
                  QryBusca.ParamByName('orden').AsString    := ImpsNumeroOrden;
                  QryBusca.Open;
                  if QryBusca.RecordCount > 0 then
                  begin

                    BotonSelec := MessageDlg('¿Desea reemplazar la programacion para la orden de la Cia?', mtConfirmation, [mbYes, mbNo, mbCancel], 0);
                    if BotonSelec = mrCancel then
                      raise Exception.Create('Proceso Cancelado por el usuario')
                    else
                    begin
                      zCommand.Active := false;
                      zCommand.SQL.Clear;
                      zCommand.SQL.Add('delete from distribuciondeactividadescia where sContrato=:contrato and sNumeroOrden=:orden and sIdConvenio=:convenio');
                      zCommand.ParamByName('contrato').AsString := ImpsContrato;
                      zCommand.ParamByName('convenio').AsString := global_convenio;
                      zCommand.ParamByName('orden').AsString := ImpsNumeroOrden;                    
                      zCommand.ExecSQL;
                    end;
                  end;
                end;
                zCommand.Active := False;
                zCommand.SQL.Clear;
                sSQL := ' INSERT INTO distribuciondeactividadescia ( sContrato , sIdConvenio, sNumeroOrden, sWbs, ' +
                  ' sNumeroActividad,dIdFecha, dCantidad) ' +
                  ' VALUES (:contrato, :convenio, :orden, :wbs, :actividad, :fecha, :cantidad) ' +
                  ' on duplicate key update dCantidad=:cantidad ';
                zCommand.SQL.Add(sSQL);
                zCommand.ParamByName('contrato').AsString := ImpsContrato;
                zCommand.ParamByName('convenio').AsString := ImpsIdConvenio;
                zCommand.ParamByName('orden').AsString := ImpsNumeroOrden;              
                zCommand.ParamByName('fecha').AsDate := dFecha;
                zCommand.ParamByName('wbs').AsString := ImpsWbs;
                zCommand.ParamByName('actividad').AsString := ImpsNumeroActividad;
                zCommand.ParamByName('cantidad').AsFloat := ImpdCantidad;
                zCommand.ExecSQL;
              end;
            except
              on e: exception do
              begin
                MessageDlg('Error: ' + e.Message, mtError, [mbok], 0);
                ImpsContrato := '';
              end;
            end;
            ProgressBar1.Max := ProgressBar1.Max + 1;
            ProgressBar1.Position := ProgressBar1.Position + 1;
            fila := fila + 1;
            sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);
          end;


        end;
  {$ENDREGION}

        Connection.zConnection.Commit; // Si todo ha sido correcto se deberá generar físicamente la información en la base de datos
        MessageDlg('Proceso Terminado con exito...', mtInformation, [mbOk], 0);
      except
        on E: Exception do
        begin
          Connection.zConnection.RollBack; // Ante un error, cancelar todos los cambios realizados

              //MessageDlg(e.ClassName + ' - ' + e.Message, mtInformation, [mbOk], 0)
          if (CompareText(e.ClassName, 'Exception') = 0) then
            MessageDlg(e.ClassName + '   ' + e.Message, mtInformation, [mbOk], 0)
          else
            if CodErr1 = '' then
              MessageDlg(e.Message, mtInformation, [mbOk], 0)
            else
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, CodErr1, CodErr2, 0);
        end;
      end;
    finally
        //Termina Exception para subir anexos...
      Existir.Destroy;
      ProgressBar1.Max := 0;
      try
        ExcelApplication1.Quit;
      except
          // No Hacer nada, solamente es para evitar los errores cuando excel se encuentre ocupado
      end;
      ExcelApplication1.Disconnect;
    end;
  {$EndRegion}
  end
  else
  begin
    {$Region 'Importaciones en Ms Project'}
    if rbtnPrograma.Checked then
    begin
      ImportOTProject(TsArchivo.Text);

    end;
    {$EndRegion}

  end;
end;

procedure TfrmImportaciondeDatos.tsArchivoEnter(Sender: TObject);
begin
  tsarchivo.Color := global_color_entrada;
end;

procedure TfrmImportaciondeDatos.tsArchivoExit(Sender: TObject);
begin
  tsarchivo.Color := global_color_salida;
end;

procedure TfrmImportaciondeDatos.btnSalirClick(Sender: TObject);
begin
  Close
end;

procedure TfrmImportaciondeDatos.Button1Click(Sender: TObject);
var
  lista: TStringList;
  sSQL, ss: string;
  ii: integer;
begin
  try
    lista := TStringList.Create;
    sSQL := 'SELECT * ' +
      'FROM actividadesxorden a ' +
      'WHERE a.sContrato = :contrato ' +
      'AND a.sIdConvenio = :convenio ' +
      'AND a.sTipoActividad = "Actividad" ' +
      'GROUP BY a.sNumeroActividad ' +
      'ORDER BY a.sNumeroActividad';

    connection.QryBusca.Active := false;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add(sSQL);
    connection.QryBusca.ParamByName('contrato').Value := global_contrato;
    connection.QryBusca.ParamByName('convenio').Value := global_convenio;
    connection.QryBusca.Open;

    while not connection.QryBusca.Eof do begin
      lista.Add(connection.QryBusca.FieldByName('sNumeroActividad').AsString);
      connection.QryBusca.Next;
    end;

    self.CalcDiferenciasOT(lista);
    RxMDValida.First;
    ss := '';
    while not RxMDValida.Eof do begin
      ss := ss + RxMDValida.FieldByName('sNumeroActividad').AsString + #13;
      RxMDValida.Next;
    end;
    showmessage(ss);

    MessageDlg('Existen diferencias. Oprima aceptar para ver el reporte.', mtInformation, [mbOk], 0);
    frxReporte.LoadFromFile(global_files + 'validaActOrden.fr3');
    frxReporte.PreviewOptions.MDIChild := True;
    frxReporte.PreviewOptions.Modal := False;
    frxReporte.PreviewOptions.Maximized := lCheckMaximized;
    frxReporte.PreviewOptions.ShowCaptions := False;
    frxReporte.Previewoptions.ZoomMode := zmPageWidth;
    frxReporte.ShowReport;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al hacer clic en el boton 1', 0);
    end;
  end;

end;


procedure TfrmImportaciondeDatos.Button2Click(Sender: TObject);
begin
  try
        //MessageDlg('Existen diferencias. Oprima aceptar para ver el reporte.' , mtInformation, [mbOk], 0) ;
    frxReporte.LoadFromFile(global_files + 'validaActOrden.fr3');
    frxReporte.PreviewOptions.MDIChild := True;
    frxReporte.PreviewOptions.Modal := False;
    frxReporte.PreviewOptions.Maximized := lCheckMaximized;
    frxReporte.PreviewOptions.ShowCaptions := False;
    frxReporte.Previewoptions.ZoomMode := zmPageWidth;
    frxReporte.ShowReport;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importacion de Plantillas', 'Al hacer clic en el Button2', 0);
    end;
  end;
end;

procedure TfrmImportaciondeDatos.formatoEncabezado;
begin
  Excel.Selection.MergeCells := False;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size := 11;
  Excel.Selection.Font.Bold := False;
  Excel.Selection.Font.Name := 'Calibri';
  Excel.Selection.Interior.Color := $00FFFAE1;
  Excel.Selection.Font.Color := clBlack;
end;

{$REGION 'ANEXO DT,DE MN, DE DLL (VALIDA)'}

function TfrmImportaciondedatos.ValidaAnexosDT(dParamTipo: string): boolean;
var
  Fila, iColumna: integer;
  Contrato, ContratoExcel,
    AnexoExcel, campo: string;
  lContinua, lExiste,
    lEncuentra, lEncuentraAnexo: boolean;
  cadena, cadena2,
    Cadena3, cadena4: string;
  Actividad, Anexo,
    AnexoAux,
    sFecha, TipoDato,
    medida, sTipo: string;
  paquete: array[1..3000, 1..3] of string;
  iNivel, x, t: integer;
  CantidadDT: currency;
  dTotalDT: currency;
  dIdFecha: tDate;
  dTotalAnexo: currency;
  CreaWbs: string;
   {Decalracion de Querys,,}
  zAnexos, zAnexosC: TZReadonlyQuery;
begin
    {Primero las validaciones de las Columnas de Fehcas..}
  ValidaAnexosDT := False;

  if dParamTipo = 'AnexoDTStruct' then
    iColumna := 6
  else
    iColumna := 4;

  zAnexos := TZReadOnlyQuery.Create(self);
  zAnexos.Connection := connection.zConnection;

  zAnexosC := TZReadOnlyQuery.Create(self);
  zAnexosC.Connection := connection.zConnection;

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  zAnexosC.Active := False;
  zAnexosC.SQL.Clear;
  zAnexosC.SQL.Add('select min(dFechaInicio) as dFechaInicio, max(dFechaFinal) as dFechaFinal from actividadesxanexo Where sContrato = :contrato and sIdConvenio =:Convenio group by sContrato');
  zAnexosC.Params.ParamByName('Contrato').DataType := ftString;
  zAnexosC.Params.ParamByName('Contrato').Value := Contrato;
  zAnexosC.Params.ParamByName('Convenio').DataType := ftString;
  zAnexosC.Params.ParamByName('Convenio').Value := Global_Convenio;
  zAnexosC.Open;

  Fila := 1;
  try
    sFecha := DateToStr(ExcelWorksheet1.Range[columnas[iColumna] + '1', columnas[iColumna] + '1'].Value2);
  except
    ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'fIncorrecto');
  end;

  while (sFecha <> '') and (sFecha <> '30/12/1899') do
  begin
    cadena := columnas[iColumna];

    TipoDato := 'Fecha';
    try
      if (sFecha <> '') and (sFecha <> '30/12/1899') then
        dIdFecha := StrToDate(sFecha);
    except
      ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'fIncorrecto');
    end;

    if zAnexosC.RecordCount > 0 then
    begin
             {Primero los años..}
      if ((StrToInt(copy(sFecha, 7, 4))) < (StrToInt(copy(dateToStr(zAnexosC.FieldValues['dFechaInicio']), 7, 4)))) then
        ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor')
      else
      begin {Ahora por meses..}
        if ((StrToInt(copy(sFecha, 7, 4))) = (StrToInt(copy(dateToStr(zAnexosC.FieldValues['dFechaInicio']), 7, 4)))) then
          if ((StrToInt(copy(sFecha, 4, 2))) < (StrToInt(copy(dateToStr(zAnexosC.FieldValues['dFechaInicio']), 4, 2)))) then
            ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor');
      end;
             {Cotinuamos con los años.}
      if ((StrToInt(copy(sFecha, 7, 4))) > (StrToInt(copy(dateToStr(zAnexosC.FieldValues['dFechaFinal']), 7, 4)))) then
        ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor')
      else
      begin {Continuamos con los meses..}
        if ((StrToInt(copy(sFecha, 7, 4))) = (StrToInt(copy(dateToStr(zAnexosC.FieldValues['dFechaFinal']), 7, 4)))) then
          if ((StrToInt(copy(sFecha, 4, 2))) > (StrToInt(copy(dateToStr(zAnexosC.FieldValues['dFechaFinal']), 4, 2)))) then
            ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor');
      end;
    end;
    Inc(iColumna);
    try
      sFecha := DateToStr(ExcelWorksheet1.Range[columnas[iColumna] + '1', columnas[iColumna] + '1'].Value2);
    except
      ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'fIncorrecto');
    end;
  end;
    {Verificamos si existen los grupos en fasesxproyecto}
  zAnexos.Active := False;
  zAnexos.SQL.Clear;
  zAnexos.SQL.Add('select sAnexo from anexos ');
  zAnexos.Open;

  zAnexosC.Active := False;
  zAnexosC.SQL.Clear;
  if dParamTipo = 'AnexoDTStruct' then
    zAnexosC.SQL.Add('select sNumeroActividad, sWbs, sAnexo, dCantidadAnexo, dVentaMN, dVentaDLL from actividadesxanexo Where sContrato = :contrato and sIdConvenio =:Convenio order by iItemOrden ')
  else
    zAnexosC.SQL.Add('select sNumeroActividad, sWbs, sAnexo, dCantidadAnexo, dVentaMN, dVentaDLL from actividadesxanexo Where sContrato = :contrato and sIdConvenio =:Convenio and sTipoActividad = "Actividad" order by iItemOrden ');
  zAnexosC.Params.ParamByName('Contrato').DataType := ftString;
  zAnexosC.Params.ParamByName('Contrato').Value := Contrato;
  zAnexosC.Params.ParamByName('Convenio').DataType := ftString;
  zAnexosC.Params.ParamByName('Convenio').Value := Global_Convenio;
  zAnexosC.Open;

    {Validacione de Contrato..}
  Fila := 2;
  Contrato := global_contrato;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;
  AnexoExcel := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;

    {Borramos Cuadro de Colores..}
  EliminaCuadro(columnas[iColumna], iColumna);

  t := 1;
  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
            {Existen datos..}
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

      if dParamTipo = 'AnexoDTStruct' then
      begin
        TipoDato := 'Numero';
        Campo := 'Nivel';
        cadena := 'B';
        iNivel := StrToInt(ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2);
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        TipoDato := 'Texto';
        campo := 'Actividad';
        cadena := 'C';
        Actividad := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        campo := 'Anexo';
        cadena := 'D';
        Anexo := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        Campo := 'Medida';
        cadena := 'E';
        medida := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        iColumna := 6;
      end;

      if dParamTipo <> 'AnexoDTStruct' then
      begin
        TipoDato := 'Texto';
        campo := 'Actividad';
        cadena := 'B';
        Actividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        campo := 'Anexo';
        cadena := 'C';
        Anexo := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');
        iColumna := 4;
      end;

      dTotalAnexo := 0;
      sFecha := ExcelWorksheet1.Range[columnas[iColumna] + '1', columnas[iColumna] + '1'].Value2;
      while (trim(sFecha) <> '') and (sFecha <> '30/12/1899') and (sFecha <> '34425') do
      begin
        cadena := columnas[iColumna];
        campo := 'Cantidad DT';
        TipoDato := 'Decimal';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');
        try
          CantidadDT := ExcelWorksheet1.Range[columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila))].Value2;
          dTotalAnexo := dTotalAnexo + CantidadDT;
        except
        end;
        Inc(iColumna);
        try
          sFecha := ExcelWorksheet1.Range[columnas[iColumna] + '1', columnas[iColumna] + '1'].Value2;
        except
        end;
      end;

      if dParamTipo = 'AnexoDTStruct' then
      begin
                {Calculamos Wbs}
        if Trim(medida) = '' then
          sTipo := 'Paquete'
        else
          sTipo := 'Actividad';

        CreaWbs := '';
        if iNivel <> 0 then
        begin
          for x := 1 to t - 1 do
          begin
            if iNivel - 1 >= strToint(paquete[x][1]) then
            begin
              if (sTipo = 'Actividad') and (Anexo <> '') then
                CreaWbs := paquete[x][2] + '.' + Anexo + '.'
              else
                CreaWbs := paquete[x][2] + '.';
            end;
          end;
          CreaWbs := CreaWbs + Actividad;
        end
        else
          CreaWbs := Actividad;

        if sTipo = 'Paquete' then
        begin
          paquete[t][1] := inttostr(iNivel);
          paquete[t][2] := CreaWbs;
          t := t + 1;
        end;
      end;

            {Verificamos si existe las partidas en el anexo C}
      zAnexosC.First;
      lEncuentra := False;
            {Buscamos los partidas en Catalogo..}
      while not zAnexosC.Eof do
      begin
        if dParamTipo = 'AnexoDTStruct' then
        begin
          if (Anexo = zAnexosC.FieldValues['sAnexo']) and (Actividad = zAnexosC.FieldValues['sNumeroActividad'])
            and (CreaWbs = zAnexosC.FieldValues['sWbs']) then
          begin
            lEncuentra := True;
            if medida <> '' then //ABBY
            begin
                              {Registramos las cantidades de Anexo Diferentes..}
                               //if RoundTo(dTotalAnexo, -5) <> RoundTo(zAnexosC.FieldValues['dCantidadAnexo'],-5) then
              if comparevalue(RoundTo(dTotalAnexo, -5), RoundTo(zAnexosC.FieldByName('dCantidadAnexo').asfloat, -5), 0.02) <> 0 then
                ColoresErrorExcel('C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila)), 3, 'DT');
            end;
          end;
        end
        else
        begin
          if (Anexo = zAnexosC.FieldValues['sAnexo']) and (Actividad = zAnexosC.FieldValues['sNumeroActividad']) then
          begin
            lEncuentra := True;

                          {Registramos las cantidades de Anexo Diferentes..}
            if dParamTipo = 'AnexoDT' then
                             //if RoundTo(dTotalAnexo, -5) <> RoundTo(zAnexosC.FieldValues['dCantidadAnexo'],-5) then
              if comparevalue(RoundTo(dTotalAnexo, -5), RoundTo(zAnexosC.FieldByName('dCantidadAnexo').asfloat, -5), 0.02) <> 0 then
                ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 3, 'DT');

            if dParamTipo = 'AnexoDEMN' then
              if comparevalue((dTotalAnexo), (zAnexosC.FieldByName('dCantidadAnexo').asfloat * zAnexosC.FieldByName('dVentaMN').asfloat), 0.02) <> 0 then
                            //if (dTotalAnexo) <> (zAnexosC.FieldValues['dCantidadAnexo'] * zAnexosC.FieldValues['dVentaMN']) then
                ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 3, 'DT');

            if dParamTipo = 'AnexoDEDLL' then
              if CompareValue((dTotalAnexo), (zAnexosC.FieldByName('dCantidadAnexo').AsFloat * zAnexosC.FieldByName('dVentaDLL').AsFloat), 0.02) <> 0 then
                ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 3, 'DT');
          end;
        end;
        zAnexosC.Next;
        if lEncuentra then
          zAnexosC.Last;
      end;

            {Validamos los anexos}
      zAnexos.First;
      lEncuentraAnexo := False;
            {Buscamos los anexos..}
      while not zAnexos.Eof do
      begin
        if Anexo = zAnexos.FieldValues['sAnexo'] then
          lEncuentraAnexo := True;
        zAnexos.Next;
      end;

      if dParamTipo = 'AnexoDTStruct' then
      begin
        if lEncuentraAnexo = False then
          ColoresErrorExcel('D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila)), 2, 'Anexo');

        if lEncuentraAnexo = True then
          if lEncuentra = False then
            ColoresErrorExcel('C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila)), 2, 'Actividad');
      end;

      if dParamTipo <> 'AnexoDTStruct' then
      begin
        if lEncuentraAnexo = False then
          ColoresErrorExcel('C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila)), 2, 'Anexo');

        if lEncuentraAnexo = True then
          if lEncuentra = False then
            ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 2, 'Actividad');
      end;

    end;
    Fila := Fila + 1;
  end;

  CuadroColores(columnas[iColumna + 1], columnas[iColumna + 2], columnas[iColumna + 3], columnas[iColumna + 9]);

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosDT := True;
    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;
    {Temrina Validacion contrato..}
end;
{$ENDREGION}
{$REGION 'ANEXOS DME, DMO, MDA (VALIDACION)'}

function TfrmImportaciondedatos.ValidaAnexosDME(dParamTipo: string; dParamTabla: string; dParamId: string): boolean;
var
  Fila, iColumna: integer;
  Contrato, ContratoExcel,
    AnexoExcel, campo: string;
  lContinua, lExiste, lEncuentra: boolean;
  cadena: string;
  Actividad, Anexo,
    AnexoAux,
    sFecha, TipoDato: string;
  CantidadDT: double;
  dIdFecha: tDate;
  dTotalAnexo: double;

   {Decalracion de Querys,,}
  zRecurso: TZReadonlyQuery;
begin
    {Primero las validaciones de las Columnas de Fehcas..}
  zRecurso := TZReadOnlyQuery.Create(self);
  zRecurso.Connection := connection.zConnection;

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  zRecurso.Active := False;
  zRecurso.SQL.Clear;
  zRecurso.SQL.Add('select min(dFechaInicio) as dFechaInicio, max(dFechaFinal) as dFechaFinal from actividadesxanexo Where sContrato = :contrato and sIdConvenio =:Convenio group by sContrato');
  zRecurso.Params.ParamByName('Contrato').DataType := ftString;
  zRecurso.Params.ParamByName('Contrato').Value := Contrato;
  zRecurso.Params.ParamByName('Convenio').DataType := ftString;
  zRecurso.Params.ParamByName('Convenio').Value := Global_Convenio;
  zRecurso.Open;

  ValidaAnexosDME := False;
  Fila := 1;
  iColumna := 3;
  try
    sFecha := DateToStr(ExcelWorksheet1.Range[columnas[iColumna] + '1', columnas[iColumna] + '1'].Value2);
  except
    ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'fIncorrecto');
  end;

  while (sFecha <> '') and (sFecha <> '30/12/1899') do
  begin
    cadena := columnas[iColumna];

    TipoDato := 'Fecha';
    try
      if (sFecha <> '') and (sFecha <> '30/12/1899') then
        dIdFecha := StrToDate(sFecha);
    except
      ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'fIncorrecto');
    end;
    if zRecurso.RecordCount > 0 then
    begin
             {Primero los años..}
      if ((StrToInt(copy(sFecha, 7, 4))) < (StrToInt(copy(dateToStr(zRecurso.FieldValues['dFechaInicio']), 7, 4)))) then
        ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor')
      else
      begin {Ahora por meses..}
        if ((StrToInt(copy(sFecha, 7, 4))) = (StrToInt(copy(dateToStr(zRecurso.FieldValues['dFechaInicio']), 7, 4)))) then
          if ((StrToInt(copy(sFecha, 4, 2))) < (StrToInt(copy(dateToStr(zRecurso.FieldValues['dFechaInicio']), 4, 2)))) then
            ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor');
      end;
             {Cotinuamos con los años.}
      if ((StrToInt(copy(sFecha, 7, 4))) > (StrToInt(copy(dateToStr(zRecurso.FieldValues['dFechaFinal']), 7, 4)))) then
        ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor')
      else
      begin {Continuamos con los meses..}
        if ((StrToInt(copy(sFecha, 7, 4))) = (StrToInt(copy(dateToStr(zRecurso.FieldValues['dFechaFinal']), 7, 4)))) then
          if ((StrToInt(copy(sFecha, 4, 2))) > (StrToInt(copy(dateToStr(zRecurso.FieldValues['dFechaFinal']), 4, 2)))) then
            ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'dtFechaMayor');
      end;
    end;
    Inc(iColumna);
    try
      sFecha := DateToStr(ExcelWorksheet1.Range[columnas[iColumna] + '1', columnas[iColumna] + '1'].Value2);
    except
      ColoresErrorExcel(columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila)), 4, 'fIncorrecto');
    end;
  end;

    {Verificamos si existen los grupos en fasesxproyecto}
  zRecurso.Active := False;
  zRecurso.SQL.Clear;
  zRecurso.SQL.Add('select ' + dParamId + ', dCantidad from ' + dParamTabla + ' Where sContrato = :contrato ');
  zRecurso.ParamByName('Contrato').AsString := contrato;
  zRecurso.Open;

    {Validacione de Contrato..}
  Fila := 2;
  ContratoExcel := global_contrato;
  lContinua := False;
  AnexoExcel := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;

    {Borramos Cuadro de Colores..}
  EliminaCuadro(columnas[iColumna], iColumna);

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}
      TipoDato := 'Texto';
      cadena := 'B';
      Actividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      iColumna := 3;
      dTotalAnexo := 0;
      sFecha := ExcelWorksheet1.Range[columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila))].Value2;
      if sFecha = '' then
        sFecha := '0';
      while (sFecha <> '') and (sFecha <> '30/12/1899') do
      begin
        cadena := columnas[iColumna];
        TipoDato := 'Decimal';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');
        try
          CantidadDT := StrToFloat(ExcelWorksheet1.Range[columnas[iColumna] + Trim(IntToStr(Fila)), columnas[iColumna] + Trim(IntToStr(Fila))].Value2);
          dTotalAnexo := dTotalAnexo + CantidadDT;
        except
        end;
        Inc(iColumna);
        try
          sFecha := DateToStr(ExcelWorksheet1.Range[columnas[iColumna] + '1', columnas[iColumna] + '1'].Value2);
        except
        end;
      end;

            {Validamos los recursos}
      zRecurso.First;
      lEncuentra := False;
            {Buscamos los recursos..}
      while not zRecurso.Eof do
      begin
        if Actividad = zRecurso.FieldValues[dParamId] then
        begin
          lEncuentra := True;
          if comparevalue(dTotalAnexo, zRecurso.Fieldbyname('dCantidad').AsFloat, 0.02) <> 0 then
                   //if (dTotalAnexo) <> (zRecurso.FieldValues['dCantidad']) then    //lobo
            ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 3, 'DMOEA');
        end;
        zRecurso.Next;
      end;

      if lEncuentra = False then
        ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 2, dParamTipo);
    end;

    Fila := Fila + 1;
  end;

  CuadroColores(columnas[iColumna + 1], columnas[iColumna + 2], columnas[iColumna + 3], columnas[iColumna + 6]);

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosDME := True;

    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;
    {Temrina Validacion contrato..}
end;
{$ENDREGION}
{$REGION 'MATERIALES (VALIDACION)'}

function TfrmImportaciondedatos.ValidaAnexosInsumo(): boolean;
var
  Fila, iColumna: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra: boolean;
  cadena, campo: string;
  Actividad, Anexo,
    AnexoAux, sValue,
    sFecha, TipoDato: string;
  CantidadDT: double;
  dIdFecha: tDate;
   //Datos
  DatosCadena: string;
  Cantidad, Costos: double;
begin

  ValidaAnexosInsumo := False;

  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 11 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el Anexo de Materiales.', mtInformation, [mbOk], 0);
    ValidaAnexosInsumo := True;
    exit;
  end;

    {Validaciones de Datos..}
  Fila := 2;
  Contrato := global_contrato;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;
  lContinua := True;

    {Borramos Cuadro de Colores..}
  EliminaCuadro('M', 0);

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
      lExiste := True;
            {Validaciones de Campos..}

      TipoDato := 'Texto';
      Campo := 'Insumo';
      cadena := 'A';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'B';
      TipoDato := 'Cadena';
      Campo := 'Tipo';
      DatosCadena := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
      if (DatosCadena <> 'Material') and (DatosCadena <> 'Consumible') and (DatosCadena <> 'Auxiliares') then
        ColoresErrorExcel(cadena + Trim(IntToStr(Fila)), cadena + Trim(IntToStr(Fila)), 2, 'Tipo');

      cadena := 'E';
      TipoDato := 'Decimal';
      Campo := 'Cantidad';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'F';
      Campo := 'Cantidad Inst.';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      cadena := 'C';
      Campo := 'Descripcion';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'D';
      Campo := 'Medida';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'G';
      TipoDato := 'Fecha';
      Campo := 'Fecha';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'H';
      TipoDato := 'Decimal';
      Campo := 'Costo MN';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'I';
      Campo := 'Costo DLL';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'J';
      Campo := 'Venta MN';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'K';
      Campo := 'Venta DLL';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

    end;
    Fila := Fila + 1;
  end;
  CuadroColores('O', 'P', 'Q', 'T');

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosInsumo := True;
    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;


    //freeandnil(QrBusca);
    {Temrina Validacion contrato..}
end;
{$ENDREGION}
{$REGION 'PERSONAL Y EQUIPO (VALIDACION)'}

function TfrmImportaciondedatos.ValidaAnexosPE(dParamTipo: string): boolean;
var
  Fila, iColumna, Orden: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra: boolean;
  cadena, campo, sTipoPerEq: string;
  Actividad, Anexo,
    AnexoAux, sValue,
    sFecha, TipoDato: string;
  CantidadDT: double;
  dIdFecha: tDate;
   //Datos
  DatosCadena: string;
  Cantidad, Costos: double;

  zTiposEQPER: TZReadonlyQuery;
begin

  ValidaAnexosPE := False;

  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 14 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el Anexo de ' + dParamTipo, mtInformation, [mbOk], 0);
    ValidaAnexosPE := True;
    exit;
  end;

  zTiposEQPER := TZReadOnlyQuery.Create(self);
  zTiposEQPER.Connection := connection.zConnection;

    {Verificamos si existen los grupos en fasesxproyecto}
  zTiposEQPER.Active := False;
  zTiposEQPER.SQL.Clear;
  if dParamTipo = 'Personal' then
    zTiposEQPER.SQL.Add('select sIdTipoPersonal from tiposdepersonal ')
  else
    zTiposEQPER.SQL.Add('select sIdTipoEquipo from tiposdeequipo ');
  zTiposEQPER.Open;

    {Validaciones de Datos..}
  Fila := 2;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;
  lContinua := True;

    {Borramos Cuadro de Colores..}
  EliminaCuadro('O', 0);

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
    sTipoPerEq := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
      lExiste := True;
            {Validaciones de Campos..}

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

      TipoDato := 'Texto';
      Campo := 'Id_' + dParamTipo;
      cadena := 'B';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'C';
      TipoDato := 'Numero';
      Campo := 'Ordenamiento';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Descripcion';
      cadena := 'D';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      Campo := 'Medida';
      cadena := 'E';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Decimal';
      Campo := 'Cantiad';
      cadena := 'F';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'G';
      Campo := 'Costo MN';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'H';
      Campo := 'Costo DLL';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'I';
      Campo := 'Venta MN';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'J';
      Campo := 'Venta DLL';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Fecha';
      Campo := 'Fecha Inicio';
      cadena := 'K';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'L';
      Campo := 'Fecha Final';
      ValidaCampo(TipoDato, cadena, Fila, campo, True, 'K');

      TipoDato := 'Texto';
      Campo := 'Id Tipo Personal';
      cadena := 'M';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Numero';
      Campo := 'Jornada';
      cadena := 'N';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      if ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2 > 24 then
        ColoresErrorExcel('N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila)), 3, 'Jornada');

      zTiposEQPER.First;
      lEncuentra := False;
            {Buscamos los frentes de trabajo..}
      while not zTiposEQPER.Eof do
      begin
        if dParamTipo = 'Personal' then
          if sTipoPerEq = zTiposEQPER.FieldValues['sIdTipoPersonal'] then
            lEncuentra := True;

        if dParamTipo = 'Equipo' then
          if sTipoPerEq = zTiposEQPER.FieldValues['sIdTipoEquipo'] then
            lEncuentra := True;
        zTiposEQPER.Next;
      end;

      if lEncuentra = False then
      begin
        if dParamTipo = 'Personal' then
          ColoresErrorExcel('M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila)), 2, 'TipoPersonal');

        if dParamTipo = 'Equipo' then
          ColoresErrorExcel('M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila)), 2, 'TipoEquipo')
      end;
    end;

    Fila := Fila + 1;
  end;
  CuadroColores('P', 'Q', 'R', 'U');

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosPE := True;
    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;
    {Temrina Validacion contrato..}
end;
{$ENDREGION}
{$REGION 'BASICOS Y HERRAMIENTAS (VALIDACION)'}

function TfrmImportaciondedatos.ValidaAnexosBasicos(dParamTipo: string): boolean;
var
  Fila, iColumna, Orden: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra: boolean;
  cadena, campo: string;
  Actividad, Anexo,
    AnexoAux, sValue,
    sFecha, TipoDato: string;
  CantidadDT: double;
  dIdFecha: tDate;
   //Datos
  DatosCadena: string;
  Cantidad, Costos: double;
begin

  ValidaAnexosBasicos := False;

  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 8 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el Anexo de ' + dParamTipo, mtInformation, [mbOk], 0);
    ValidaAnexosBasicos := True;
    exit;
  end;

    {Validaciones de Datos..}
  Fila := 2;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;
  lContinua := True;

   {Borramos Cuadro de Colores..}
  EliminaCuadro('J', 0);

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}

      TipoDato := 'Texto';
      Campo := 'Id_' + dParamTipo;
      cadena := 'B';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      Campo := 'Descripcion';
      cadena := 'C';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      Campo := 'Medida';
      cadena := 'D';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Decimal';
      Campo := 'Costo MN';
      cadena := 'E';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'F';
      Campo := 'Costo DLL';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'G';
      Campo := 'Venta MN';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      cadena := 'H';
      Campo := 'Venta DLL';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');
    end;

    Fila := Fila + 1;
  end;
  CuadroColores('J', 'K', 'L', 'O');

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosBasicos := True;
    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;
    {Temrina Validacion contrato..}
end;
{$ENDREGION}
{$REGION 'PERSONAL X PARTIDA, EQUIPO, HERRMIENTA, BASICOS (VALIDACION)'}

function TfrmImportaciondedatos.ValidaAnexosPERxP(dParamTipo: string; sParamTabla: string; dParamCampo: string): boolean;
var
  Fila, iColumna, Orden: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra: boolean;
  cadena, campo: string;
  Actividad, Anexo,
    AnexoAux, sValue,
    sFecha, TipoDato: string;
  CantidadDT: double;
  dIdFecha: tDate;
   //Datos
  DatosCadena: string;
  Cantidad, Costos: double;

   {Decalracion de Querys,,}
  zAnexo, zAnexoC, zPersonal: TZReadonlyQuery;
begin

  ValidaAnexosPERxP := False;

  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 5 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el Anexo de ' + dParamTipo, mtInformation, [mbOk], 0);
    ValidaAnexosPERxP := True;
    exit;
  end;

  zAnexo := TZReadOnlyQuery.Create(self);
  zAnexo.Connection := connection.zConnection;

  zAnexoC := TZReadOnlyQuery.Create(self);
  zAnexoC.Connection := connection.zConnection;

  zPersonal := TZReadOnlyQuery.Create(self);
  zPersonal.Connection := connection.zConnection;

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

    {Verificamos si existen anexos}
  zAnexo.Active := False;
  zAnexo.SQL.Clear;
  zAnexo.SQL.Add('select sAnexo from anexos');
  zAnexo.Open;

    {Verificamos si existen partidas en anexo C}
  zAnexoC.Active := False;
  zAnexoC.SQL.Clear;
  zAnexoC.SQL.Add('select sAnexo, sNumeroActividad from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and sTipoActividad = "Actividad"');
  zAnexoC.ParamByName('Contrato').AsString := Contrato;
  zAnexoC.ParamByName('Convenio').AsString := global_convenio;
  zAnexoC.Open;

    {Verificamos si existen el personal, equipo, material, basico, herramienta.}
  zPersonal.Active := False;
  zPersonal.SQL.Clear;
  zPersonal.SQL.Add('select ' + dParamCampo + ' from ' + sParamTabla + ' where sContrato =:Contrato ');
  zPersonal.ParamByName('Contrato').AsString := Contrato;
  zPersonal.Open;

    {Validaciones de Datos..}
  Fila := 2;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;
  lContinua := False;
  AnexoExcel := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;

    {Borramos Cuadro de Colores..}
  EliminaCuadro('G', 0);

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}
      Anexo := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
      DatosCadena := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
      TipoDato := 'Texto';
      Campo := 'Id_' + dParamTipo;
      cadena := 'D';
      Actividad := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
      TipoDato := 'Decimal';
      Campo := 'Cantidad';
      cadena := 'E';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      zAnexo.First;
      lEncuentra := False;
            {Buscamos las anexos..}
      while not zAnexo.Eof do
      begin
        if anexo = zAnexo.FieldValues['sAnexo'] then
          lEncuentra := True;
        zAnexo.Next;
      end;

      if lEncuentra = False then
        ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 2, 'Anexo');

      zPersonal.First;
      lEncuentra := False;
            {Buscamos las personal..}
      while not zPersonal.Eof do
      begin
        if Actividad = zPersonal.FieldValues[dParamCampo] then
          lEncuentra := True;
        zPersonal.Next;
      end;

      if lEncuentra = False then
        ColoresErrorExcel('D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila)), 2, dParamTipo);

            {Verificamos si existe las partidas en el anexo C}
      zAnexoC.First;
      lEncuentra := False;
            {Buscamos los partidas en Catalogo..}
      while not zAnexoC.Eof do
      begin
        if (anexo = zAnexoC.FieldValues['sAnexo']) and (DatosCadena = zAnexoC.FieldValues['sNumeroActividad']) then
          lEncuentra := True;

        zAnexoC.Next;
        if lEncuentra then
          zAnexoC.Last;
      end;

      if lEncuentra = False then
        ColoresErrorExcel('C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila)), 2, 'Actividad');

    end;
    Fila := Fila + 1;
  end;
  CuadroColores('G', 'H', 'I', 'L');

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosPERxP := True;
    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;
    {Temrina Validacion contrato..}
end;
{$ENDREGION}
{$REGION 'VALIDACION DE AVANCES'}

function TfrmImportaciondedatos.ValidaAvancesProgramados(): boolean;
var
  Fila, iColumna, Nivel, iDato: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra, lValidaAnexo: boolean;
  campo, cadena, sValue, medida: string;
  Actividad, Anexo,
    TipoDato,
    grupo, tipo: string;



begin
  Application.ProcessMessages;

    //Validamos antes de reemplazar Anexo C..
  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 9 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el Avance Programado.', mtInformation, [mbOk], 0);
    ValidaAvancesProgramados := True;
    exit;
  end;

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  {Validaciones de Datos..}
  Fila := 2;
  Contrato := global_contrato;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;

  EliminaCuadro('R', 0);

  EliminaCuadro('O', 0);

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el Contrato es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
            {Existen datos..}
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}

      TipoDato := 'Fecha';
      Campo := 'Fecha';
      cadena := 'B';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Numero de Orden';
      cadena := 'C';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Decimal';
      Campo := 'Avance Ponderado del Dia';
      cadena := 'D';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Decimal';
      Campo := 'Avance Ponderado Global';
      cadena := 'E';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Decimal';
      Campo := 'Avance Ponderado Financiero';
      cadena := 'E';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Numero';
      Campo := 'Numero Gerencial';
      cadena := 'G';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Duracion';
      cadena := 'H';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Horario';
      cadena := 'I';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

    end;
    Fila := Fila + 1;
  end;



    {Temrina Validacion inicial..}
end;

{$ENDREGION}
{$REGION 'VALIDACION ANEXO DT CIA'}

function TfrmImportaciondedatos.ValidaAnexoDTCia(): boolean;
var
  Fila, iColumna, Nivel, iDato: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra, lValidaAnexo: boolean;
  campo, cadena, sValue, medida: string;
  Actividad, Anexo,
    TipoDato,
    grupo, tipo: string;
begin
  Application.ProcessMessages;

    //Validamos antes de reemplazar Anexo C..
  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 6 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el ANEXO DT Cia.', mtInformation, [mbOk], 0);
    ValidaAnexoDTCia := True;
    exit;
  end;

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  {Validaciones de Datos..}
  Fila := 2;
  Contrato := global_contrato;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;

  EliminaCuadro('R', 0);

  EliminaCuadro('O', 0);

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el Contrato es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
            {Existen datos..}
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}

      TipoDato := 'Texto';
      Campo := 'Convenio';
      cadena := 'B';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Fecha';
      Campo := 'Fecha';
      cadena := 'C';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Wbs';
      cadena := 'D';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Actividad';
      cadena := 'E';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Decimal';
      Campo := 'Cantidad';
      cadena := 'F';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');
    end;
    Fila := Fila + 1;
  end;



    {Temrina Validacion inicial..}
end;

{$ENDREGION}
{$REGION 'VALIDACION ANEXO ORDEN DT CIA'}

function TfrmImportaciondedatos.ValidaAnexoDTOrdenCia(): boolean;
var
  Fila, iColumna, Nivel, iDato: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra, lValidaAnexo: boolean;
  campo, cadena, sValue, medida: string;
  Actividad, Anexo,
    TipoDato,
    grupo, tipo: string;
begin
  Application.ProcessMessages;

    //Validamos antes de reemplazar Anexo C..
  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 7 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el ANEXO DT Cia.', mtInformation, [mbOk], 0);
    ValidaAnexoDTOrdenCia := True;
    exit;
  end;

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  {Validaciones de Datos..}
  Fila := 2;
  Contrato := global_contrato;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;

  EliminaCuadro('R', 0);

  EliminaCuadro('O', 0);

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el Contrato es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
            {Existen datos..}
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}

      TipoDato := 'Texto';
      Campo := 'Convenio';
      cadena := 'B';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Frente';
      cadena := 'C';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Wbs';
      cadena := 'D';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Texto';
      Campo := 'Actividad';
      cadena := 'E';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');
                  
      TipoDato := 'Fecha';
      Campo := 'Fecha';
      cadena := 'F';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Decimal';
      Campo := 'Cantidad';
      cadena := 'G';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

    end;
    Fila := Fila + 1;
  end;



    {Temrina Validacion inicial..}
end;

{$ENDREGION}
{$REGION 'ANEXO A (VALIDACION)'}
function TfrmImportaciondedatos.ValidaAnexosA(): boolean;
var
  Fila, iColumna: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra: boolean;
  cadena, Campo: string;
  Actividad, Anexo, sValue,
    AnexoAux, TipoDato,
    grupo, plataforma: string;
  CantidadDT: double;
  dIdFecha: tDate;
   //Datos
  DatosCadena: string;
  Cantidad, Costos: double;
   {Decalracion de Querys,,}
  zGrupo, zPlataforma: TZReadonlyQuery;

begin
  ValidaAnexosA := False;

  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if iColumna <> 5 then
  begin
    MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el Anexo A.', mtInformation, [mbOk], 0);
    ValidaAnexosA := True;
    exit;
  end;

  zGrupo := TZReadOnlyQuery.Create(self);
  zGrupo.Connection := connection.zConnection;

  zPlataforma := TZReadOnlyQuery.Create(self);
  zPlataforma.Connection := connection.zConnection;

    {Verificamos si existen los grupos de isometricos}
  zGrupo.Active := False;
  zGrupo.SQL.Clear;
  zGrupo.SQL.Add('select sIdGrupo from gruposisometrico');
  zGrupo.Open;

    {Verificamos si existen plataformas}
  zPlataforma.Active := False;
  zPlataforma.SQL.Clear;
  zPlataforma.SQL.Add('select sIdPlataforma from plataformas ');
  zPlataforma.Open;

    {Validaciones de Datos..}
  Fila := 2;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;
  lContinua := False;
  AnexoExcel := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;

    {Borramos Cuadro de Colores..}
  EliminaCuadro('G', 0);

    {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
    Contrato := global_contrato
  else
    Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}
      TipoDato := 'Texto';
      cadena := 'B';
      Campo := 'No. Plano';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      TipoDato := 'Numero';
      cadena := 'C';
      Campo := 'Revision';
      ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

      grupo := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
      plataforma := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;

      zGrupo.First;
      lEncuentra := False;
            {Buscamos los grupos..}
      while not zGrupo.Eof do
      begin
        if grupo = zGrupo.FieldValues['sIdGrupo'] then
          lEncuentra := True;
        zGrupo.Next;
      end;

      if lEncuentra = False then
        ColoresErrorExcel('D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila)), 2, 'Grupo');

      zPlataforma.First;
      lEncuentra := False;
            {Buscamos las plataformas..}
      while not zPlataforma.Eof do
      begin
        if plataforma = zPlataforma.FieldValues['sIdPlataforma'] then
          lEncuentra := True;
        zPlataforma.Next;
      end;

      if lEncuentra = False then
        ColoresErrorExcel('E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila)), 2, 'Plataforma');
    end;

    Fila := Fila + 1;
  end;

  CuadroColores('G', 'H', 'I', 'L');

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosA := True;
    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;
    {Temrina Validacion contrato..}
end;
{$ENDREGION}
{$REGION 'ANEXO C (VALIDACION)'}

function TfrmImportaciondedatos.ValidaAnexosC(dParamAnexo: string): boolean;
var
  Fila, iColumna, Nivel, iDato: integer;
  Contrato, ContratoExcel,
    AnexoExcel: string;
  lContinua, lExiste, lEncuentra, lValidaAnexo: boolean;
  campo, cadena, sValue, medida: string;
  Actividad, Anexo,
    TipoDato,
    grupo, tipo: string;
  dIdFecha: tDate;
   //Datos
  DatosCadena, Frente,  sConvenio : string;
  Cantidad, Costos, Ponderado: double;
   {Decalracion de Querys,,}
  zFases, zAnexos, zReprog : TZReadonlyQuery;

begin
  Application.ProcessMessages;
  ValidaAnexosC := False;
  lValidaAnexo := False;
    //Validamos antes de reemplazar Anexo C..
  Fila := 1;
  iColumna := 0;
  sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
  while sValue <> '' do
  begin
    sValue := ExcelWorksheet1.Range[columnas[Fila] + '1', columnas[Fila] + '1'].Value2;
    if (sValue <> '') then
      Inc(iColumna);

    Fila := Fila + 1;
  end;

  if dParamAnexo = 'AnexoC' then
  begin
    if iColumna <> 16 then
    begin
      MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar el Anexo C.', mtInformation, [mbOk], 0);
      ValidaAnexosC := True;
      lValidaAnexo := True;
      exit;
    end;
  end
  else
  begin
    if iColumna <> 13 then
    begin
      MessageDlG('El Archivo de Excel Seleccionado no Corresponde al Formato (Plantilla) para Importar Los puntos de programa.', mtInformation, [mbOk], 0);
      ValidaAnexosC := True;
      lValidaAnexo := True;
      exit;
    end;
  end;

  zFases := TZReadOnlyQuery.Create(self);
  zFases.Connection := connection.zConnection;

  zAnexos := TZReadOnlyQuery.Create(self);
  zAnexos.Connection := connection.zConnection;

  zReprog := TZReadOnlyQuery.Create(self);
  zReprog.Connection := connection.zConnection;

  {Si se elige contrato de Excel o contrato actual..}
  if lContratoActual then
     Contrato := global_contrato
  else
     Contrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

  {Cosnultamos Fases, Anexos o Frentes}
  if dParamAnexo = 'AnexoC' then
  begin
     {Verificamos si existen los grupos en fasesxproyecto}
     zFases.Active := False;
     zFases.SQL.Clear;
     zFases.SQL.Add('select sIdFase from fasesxproyecto ');
     zFases.Open;
  end
  else
  begin
     {Verificamos si existen los grupos en fasesxproyecto}
     zFases.Active := False;
     zFases.SQL.Clear;
     zFases.SQL.Add('select sNumeroOrden from ordenesdetrabajo where sContrato =:Contrato ');
     zFases.Params.ParamByName('Contrato').DataType := ftString;
     zFases.Params.ParamByName('Contrato').Value := contrato;
     zFases.Open;

     zReprog.Active := False;
     zReprog.SQL.Clear;
     zReprog.SQL.Add('select sIdConvenio from convenios where sContrato =:Contrato and sNumeroOrden =:Folio ');
     zReprog.Params.ParamByName('Contrato').DataType := ftString;
     zReprog.Params.ParamByName('Contrato').Value    := contrato;
     zReprog.Params.ParamByName('Folio').DataType    := ftString;
     zReprog.Params.ParamByName('Folio').Value       := ExcelWorksheet1.Range['C' + Trim(IntToStr(2)), 'C' + Trim(IntToStr(2))].Value2;
     zReprog.Open;
  end;

  {Verificamos si existen las anexos}
  zAnexos.Active := False;
  zAnexos.SQL.Clear;
  zAnexos.SQL.Add('select sAnexo from anexos ');
  zAnexos.Open;

    {Validaciones de Datos..}
  Fila := 2;
  Ponderado := 0;
  Contrato := global_contrato;
  ContratoExcel := global_contrato;
  lContinua := False;
  lExiste := False;
  AnexoExcel := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;

    {Borramos Cuadro de Colores..}
  if dParamAnexo = 'AnexoC' then
    EliminaCuadro('R', 0)
  else
    EliminaCuadro('O', 0);

  while ContratoExcel <> '' do
  begin
    ContratoExcel := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

        {Si el contrro es Diferente a Caracter en Blanco..}
    if ContratoExcel <> '' then
    begin
            {Existen datos..}
      lExiste := True;

      if contrato <> ContratoExcel then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila)), 2, 'Contrato');

            {Validaciones de Campos..}
      if dParamAnexo = 'AnexoC' then
      begin
        TipoDato := 'Numero';
        Campo := 'Nivel';
        cadena := 'B';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        TipoDato := 'Texto';
        Campo := 'Actividad';
        cadena := 'C';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        Campo := 'Descripcion';
        cadena := 'E';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        Campo := 'Medida';
        cadena := 'F';
        medida := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        TipoDato := 'Decimal';
        Campo := 'Cantidad';
        cadena := 'G';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        cadena := 'H';
        Campo := 'Ponderado';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        cadena := 'I';
        Campo := 'Precio MN';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        cadena := 'J';
        Campo := 'Precio DLL';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        TipoDato := 'Fecha';
        Campo := 'Fecha Inicio';
        cadena := 'L';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        Campo := 'Fecha Final';
        cadena := 'M';
        ValidaCampo(TipoDato, cadena, Fila, campo, true, 'L');

        grupo := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;
        Anexo := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;

        TipoDato := 'Texto';
        Campo := 'Tipo(PU/ADM)';
        cadena := 'O';
        tipo := ExcelWorksheet1.Range['O' + Trim(IntToStr(Fila)), 'O' + Trim(IntToStr(Fila))].Value2;

        if (trim(medida) <> '') and (tipo <> 'PU') and (tipo <> 'ADM') then
          ColoresErrorExcel(cadena + Trim(IntToStr(Fila)), cadena + Trim(IntToStr(Fila)), 2, 'Tipo');

        TipoDato := 'Texto';
        Campo := 'Extraordinaria(Si/No)';
        cadena := 'P';
        tipo := ExcelWorksheet1.Range['P' + Trim(IntToStr(Fila)), 'P' + Trim(IntToStr(Fila))].Value2;
        if (tipo <> 'Si') and (tipo <> 'No') then
          ColoresErrorExcel(cadena + Trim(IntToStr(Fila)), cadena + Trim(IntToStr(Fila)), 2, 'Extraordinaria');
      end
      else
      begin {Frentes de trabajo o actividadesxorden..}
        sConvenio := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;

        Frente    := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
        zq_Esp.Locate('sIdFolio', Frente, [loCaseInsensitive]);
        Frente := zq_Esp.FieldByName('sNumeroOrden').AsString;

        TipoDato := 'Numero';
        Campo := 'Nivel';
        cadena := 'D';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        TipoDato := 'Texto';
        Campo := 'Actividad';
        cadena := 'E';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        Campo := 'Descripcion';
        cadena := 'F';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        Campo := 'Medida';
        cadena := 'G';
        medida := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        TipoDato := 'Decimal';
        Campo := 'Cantidad';
        cadena := 'H';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        cadena := 'I';
        Campo := 'Ponderado';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');
        if trim(medida) <> '' then
           Ponderado :=  XRound(Ponderado, 2) + XRound(ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2, 2);

        TipoDato := 'Fecha';
        Campo := 'Fecha Inicio';
        cadena := 'J';
        ValidaCampo(TipoDato, cadena, Fila, campo, false, '');

        cadena := 'K';
        Campo := 'Fecha Final';
        ValidaCampo(TipoDato, cadena, Fila, campo, True, 'K');

        Anexo := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;

        TipoDato := 'Texto';
        Campo := 'Tipo(PU/ADM)';
        cadena := 'L';
        tipo := ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;

        if (trim(medida) <> '') and (tipo <> 'PU') and (tipo <> 'ADM') then
          ColoresErrorExcel(cadena + Trim(IntToStr(Fila)), cadena + Trim(IntToStr(Fila)), 2, 'Tipo');
      end;

      if dParamAnexo = 'AnexoC' then
      begin
                {Buscamos las fases..}
        zFases.First;
        lEncuentra := False;
        while not zFases.Eof do
        begin
          if grupo = zFases.FieldValues['sIdFase'] then
            lEncuentra := True;
          zFases.Next;
        end;

        if (trim(medida) = '') and (trim(grupo) <> '') then
          ColoresErrorExcel('K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila)), 2, 'FasePaq');

        if trim(medida) <> '' then
          if lEncuentra = False then
            ColoresErrorExcel('K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila)), 2, 'Fase');
      end
      else
      begin
        zFases.First;
        lEncuentra := False;
                {Buscamos los frentes de trabajo..}
        while not zFases.Eof do
        begin
          if Frente = zFases.FieldValues['sNumeroOrden'] then
            lEncuentra := True;
          zFases.Next;
        end;

        if lEncuentra = False then
          ColoresErrorExcel('C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila)), 2, 'Frente');

        zReprog.First;
        lEncuentra := False;
                {Buscamos los frentes de trabajo..}
//        while not zReprog.Eof do
//        begin
//          if sConvenio = zReprog.FieldValues['sIdConvenio'] then
//            lEncuentra := True;
//          zReprog.Next;
//        end;
//
//        if lEncuentra = False then
//          ColoresErrorExcel('B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila)), 2, 'Convenio');
      end;

      zAnexos.First;
      lEncuentra := False;
            {Buscamos los anexos..}
      while not zAnexos.Eof do
      begin
        if Anexo = zAnexos.FieldValues['sAnexo'] then
          lEncuentra := True;
        zAnexos.Next;
      end;

      if dParamAnexo <> 'AnexoC' then
         lEncuentra := True;

      if lEncuentra = False then
      begin
        if dParamAnexo = 'AnexoC' then
          ColoresErrorExcel('N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila)), 2, 'Anexo')
        else
          ColoresErrorExcel('M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila)), 2, 'Anexo')
      end;

    end;
    Fila := Fila + 1;
  end;

  if lExiste = False then
  begin
    MessageDlg('No se encontraron Datos para Importar!', mtInformation, [mbOk], 0);
    ValidaAnexosC := True;
    lValidaAnexo := True;
  end
  else
  begin
     if Ponderado < 100.00 then
        ColoresErrorExcel('I' + Trim(IntToStr(2)), 'I' + Trim(IntToStr(Fila-1)), 2, 'PonderadoMenor');

       if Ponderado > 100.00 then
        ColoresErrorExcel('I' + Trim(IntToStr(2)), 'I' + Trim(IntToStr(Fila-1)), 2, 'PonderadoMayor');
  end;

  if lValidaAnexo = False then
    if PartidasRepetidas(dParamAnexo) then
      ValidaAnexosC := True;

  if dParamAnexo = 'AnexoC' then
    CuadroColores('R', 'S', 'T', 'W')
  else
    CuadroColores('O', 'P', 'Q', 'T');

    {Generamos cuadro de codigo de colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
    ValidaAnexosC := True;
    lValidaAnexo := True;

    ExcelApplication1.UserControl := False;
    ExcelApplication1.Interactive[flcid] := True;
    ExcelApplication1.Disconnect;
  end;

  Fila := 2;
  lContinua := False;
  if dParamAnexo = 'AnexoC' then
    Nivel := StrToInt(ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2)
  else
    Nivel := StrToInt(ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2);

  if Nivel = 0 then
    lContinua := True;

  if lContinua = False then
  begin
    messageDLG('No se encontro en la Fila 2 el Nivel 0 del Paquete Principal en el Archivo de Excel.', mtInformation, [mbOk], 0);
    ValidaAnexosC := True;
    lValidaAnexo := True;
  end;

    {Temrina Validacion inicial..}
end;
{$ENDREGION}
{$REGION 'ELIMINA PERSONAL, EQUIPO, MATERIALES, HERRAMIENTAS, BASICOS (VALIDACION)'}

function TfrmImportaciondedatos.ValidaDeleteAnexosP(dParamTabla, dParamId, dParamBuscaTabla, dParamBuscaTabla2: string): boolean;
var
  Fila: integer;
  Id, cadena, cadena2, sValue: string;
begin
    //Validaciones antes de insertar..
  ValidaDeleteAnexosP := False;
  Fila := 2;
  sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
  while (sValue <> '') do
  begin
          {Validamos que los ids no se encuentren reportados...}
    Id := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

    if (dParamTabla <> 'basicos') and (dParamTabla <> 'herramientas') and (dParamTabla <> 'isometricos') then
    begin
      if dParamBuscaTabla = 'bitacorademateriales' then
        dParamId := 'sIdMaterial';

      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('select ' + dParamId + ' from ' + dParamBuscaTabla + ' Where sContrato = :contrato and ' + dParamId + ' =:Id limit 1 ');
      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
      Connection.zCommand.Params.ParamByName('Id').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Id').Value := Id;
      connection.zCommand.Open;

      if connection.zCommand.RecordCount > 0 then
        cadena := cadena + Id + ' , ';

      if dParamBuscaTabla = 'bitacorademateriales' then
        dParamId := 'sIdInsumo';
    end;

    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select ' + dParamId + ' from ' + dParamBuscaTabla2 + ' Where sContrato = :contrato and ' + dParamId + ' =:Id limit 1 ');
    Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
    Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
    Connection.zCommand.Params.ParamByName('Id').DataType := ftString;
    Connection.zCommand.Params.ParamByName('Id').Value := Id;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
      cadena2 := cadena2 + Id + ' , ';

    Fila := Fila + 1;
    sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
  end;

  if cadena <> '' then
  begin
    messageDLG('No se puede continuar, Existen Ids de ' + dParamTabla + ' Reportados. ' + #13 + 'Id : ' + cadena, mtInformation, [mbOk], 0);
    ValidaDeleteAnexosP := True;
  end
  else
    if cadena2 <> '' then
    begin
      if dParamTabla <> 'isometricos' then
        messageDLG('No se puede continuar, Existen Ids de ' + dParamTabla + ' en Recursos por Partida. ' + #13 + 'Id : ' + cadena2, mtInformation, [mbOk], 0)
      else
        messageDLG('No se puede continuar, Existen Ids de ' + dParamTabla + ' en Generadores de Obra. ' + #13 + 'Id : ' + cadena2, mtInformation, [mbOk], 0);
      ValidaDeleteAnexosP := True;
    end
    else
    begin
      Fila := 2;
      if dParamTabla = 'isometricos' then
        dParamId := 'sIsometrico';

      sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
      while (sValue <> '') do
      begin
        Id := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
              // Se elimina el catalogo de Anexo
        try
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('DELETE FROM ' + dParamTabla + ' Where sContrato = :contrato and ' + dParamId + ' =:Id ');
          Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          Connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato;
          Connection.zCommand.Params.ParamByName('Id').DataType := ftString;
          Connection.zCommand.Params.ParamByName('Id').Value := Id;
          connection.zCommand.ExecSQL();
        except
          on e: exception do begin
            UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Importación de Plantilla de Anexos', 'Al eliminar registro', 0);
            ValidaDeleteAnexosP := True;
            exit;
          end;
        end;
        Fila := Fila + 1;
        sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
      end;
    end;
end;
{$ENDREGION}
{$REGION 'MENSAJES CODIGO COLORES'}

procedure TfrmImportaciondeDatos.ColoresErrorExcel(sFila: string; sColumna: string; iTipo: Integer; sMensaje: string);
var
  color: array[1..5] of integer;
begin
  color[1] := 2; {Blanco}
  color[2] := 6; {Amarillo}
  color[3] := 3; {Rojo}
  color[4] := 5; {Azul}
  color[5] := 6; {no se}

  if (iTipo = 3) or (iTipo = 4) then
    ExcelApplication1.Range[sFila, sColumna].font.Color := clWhite
  else
    ExcelApplication1.Range[sFila, sColumna].font.Color := clBlack;
  ExcelApplication1.Range[sFila, sColumna].font.Name := 'Arial';
  ExcelApplication1.Range[sFila, sColumna].Interior.ColorIndex := color[iTipo];

    {Llenamos los mensajes al Array..}
  if sMensaje = 'Contrato' then
    if pos('CONTRATOS NO ENCONTRADOS EN INTELIGENT', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'CONTRATOS NO ENCONTRADOS EN INTELIGENT, ';

  if sMensaje = 'Convenio' then
    if pos('CONVENIO O REPROGRAMACION NO ENCONTRADOS EN INTELIGENT', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'CONVENIO O REPROGRAMACION NO ENCONTRADOS EN INTELIGENT, ';

  if sMensaje = 'Actividad' then
    if pos('PARTIDAS NO ENCONTRADAS EN CATALOGO DE ANEXO C', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'PARTIDAS NO ENCONTRADAS EN CATALOGO DE ANEXO C, ';

  if sMensaje = 'DT' then
    if pos('PARTIDAS CON DISTRIBUCIONES DIFERENTES A LA CANTIDAD DE ANEXO', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'PARTIDAS CON DISTRIBUCIONES DIFERENTES A LA CANTIDAD DE ANEXO, ';

  if sMensaje = 'DMOEA' then
    if pos('RECURSOS CON DISTRIBUCIONES DIFERENTES A LA CANTIDAD TOTAL', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'RECURSOS CON DISTRIBUCIONES DIFERENTES A LA CANTIDAD TOTAL, ';

  if sMensaje = 'Grupo' then
    if pos('GRUPOS NO ENCONTRADOS EN CATALOGOS DE PLANOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'GRUPOS NO ENCONTRADOS EN CATALOGOS DE PLANOS, ';

  if sMensaje = 'Plataforma' then
    if pos('PLATAFORMAS NO ENCONTRADAS EN CATALOGOS DE PLATAFORMAS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'PLATAFORMAS NO ENCONTRADAS EN CATALOGOS DE PLATAFORMAS, ';

  if sMensaje = 'Tipo' then
    if pos('TIPO DE PARTIDA NO PERMITIDO', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'TIPO DE PARTIDA NO PERMITIDO, ';

  if sMensaje = 'TipoPersonal' then
    if pos('IDS TIPOS DE PERSONAL NO ENCONTRADOS EN ADMINISTRACION DE CATALOGOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS TIPOS DE PERSONAL NO ENCONTRADOS EN ADMINISTRACION DE CATALOGOS, ';

  if sMensaje = 'TipoEquipo' then
    if pos('IDS TIPOS DE EQUIPO NO ENCONTRADOS EN ADMINISTRACION DE CATALOGOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS TIPOS DE EQUIPO NO ENCONTRADOS EN ADMINISTRACION DE CATALOGOS, ';

  if sMensaje = 'TipoPaq' then
    if pos('NO SE PERMITEN TIPOS(PU/ADM) EN PAQUETES', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'NO SE PERMITEN TIPOS(PU/ADM) EN PAQUETES, ';

  if sMensaje = 'Anexo' then
    if pos('IDS DE ANEXOS NO ENCONTRADOS EN ADMINISTRACION DE CATALOGOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE ANEXOS NO ENCONTRADOS EN ADMINISTRACION DE CATALOGOS, ';

  if sMensaje = 'Frente' then
    if pos('FRENTES NO ENCONTRADOS EN REGISTRO DE FRENTES DE TRABAJO', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'FRENTES NO ENCONTRADOS EN REGISTRO DE FRENTES DE TRABAJO, ';

  if sMensaje = 'Fase' then
    if pos('ID DE FASES NO ENCONTRADAS EN ADMINISTRACION DE CATALOGOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'ID DE FASES NO ENCONTRADAS EN ADMINISTRACION DE CATALOGOS, ';

  if sMensaje = 'FasePaq' then
    if pos('LOS PAQUETES NO DEBEN TENER FASES DE PROYECTO ASIGNADAS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'LOS PAQUETES NO DEBEN TENER FASES DE PROYECTO ASIGNADAS, ';

  if sMensaje = 'Iguales' then
    if pos('PARTIDAS O PAQUETES DUPLICADOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'PARTIDAS O PAQUETES DUPLICADOS, ';

  if sMensaje = 'Extraordinaria' then
    if pos('ESPECIFICACION INCORRECTA DE PARTIDA EXTRAORDINARIA COLOCAR(Si/No)', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'ESPECIFICACION INCORRECTA DE PARTIDA EXTRAORDINARIA COLOCAR(Si/No), ';

    {Cantidades..}
  if sMensaje = 'dNulo' then
    if pos('CANTIDADES NULL', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'CANTIDADES NULL, ';

  if sMensaje = 'dIncorrecto' then
    if pos('CANTIDADES INCORRECTAS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'CANTIDADES INCORRECTAS, ';

  if sMensaje = 'dNegativo' then
    if pos('CANTIDADES NEGATIVAS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'CANTIDADES NEGATIVAS, ';

  if sMensaje = 'Jornada' then
    if pos('JORNADAS MAYORES A 24', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'JORNADAS MAYORES A 24, ';

    {Textos..}
  if sMensaje = 'tNulo' then
    if pos('NO SE ACEPTAN TEXTOS NULL', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'NO SE ACEPTAN TEXTOS NULL, ';

  if sMensaje = 'tMedidaError' then
    if pos('UNIDADES DE MEDIDA INCORRECTAS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'UNIDADES DE MEDIDA INCORRECTAS, ';

    {Numeros..}
  if sMensaje = 'nNulo' then
    if pos('NO SE ACEPTAN NUMEROS NULL', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'NO SE ACEPTAN NUMEROS NULL, ';

  if sMensaje = 'nIncorrecto' then
    if pos('NUMEROS INCORRECTOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'NUMEROS INCORRECTOS, ';

  if sMensaje = 'nNegativo' then
    if pos('NUMEROS ENTEROS NEGATIVOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'NUMEROS ENTEROS NEGATIVOS, ';

  if sMensaje = 'nNivel' then
    if pos('SALTO DE NIVELES NO PERMITIDOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'SALTO DE NIVELES NO PERMITIDOS, ';

  if sMensaje = 'nDecimal' then
    if pos('CANTIDADES DECIMALES EN NUMEROS ENTEROS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'CANTIDADES DECIMALES EN NUMEROS ENTEROS, ';

  if sMensaje = 'Alcances' then
    if pos('SUMATORIA DE ALCANCES POR PARTIDA NO PERMITIDA (DIFERENTE AL 100%)', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'SUMATORIA DE ALCANCES POR PARTIDA NO PERMITIDA (DIFERENTE AL 100%), ';

  if sMensaje = 'PonderadoMenor' then
    if pos('SUMATORIA DE PONDERADOS MENOR A 100%', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'SUMATORIA DE PONDERADOS MENOR A 100%, ';

  if sMensaje = 'PonderadoMayor' then
    if pos('SUMATORIA DE PONDERADOS MAYOR A 100%', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'SUMATORIA DE PONDERADOS MAYOR A 100%, ';

    {Fechas..}
  if sMensaje = 'fNulo' then
    if pos('NO SE ACEPTAN FECHAS NULL', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'NO SE ACEPTAN FECHAS NULL, ';

  if sMensaje = 'fIncorrecto' then
    if pos('FECHAS INCORRECTAS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'FECHAS INCORRECTAS, ';

  if sMensaje = 'fMenores' then
    if pos('FECHAS DE TERMINO MENORES A FECHAS DE INICIO', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'FECHAS DE TERMINO MENORES A FECHAS DE INICIO, ';

  if sMensaje = 'dtFechaMayor' then
    if pos('FECHAS FUERA DEL RANGO DEL CONTRATO', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'FECHAS FUERA DEL RANGO DEL CONTRATO, ';

    {Insumos..}
  if sMensaje = 'Personal' then
    if pos('IDS DE PERSONAL NO ENCONTRADOS EN CATALOGOS DE PERSONAL', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE PERSONAL NO ENCONTRADOS EN CATALOGOS DE PERSONAL, ';

  if sMensaje = 'Equipo' then
    if pos('IDS DE EQUIPO NO ENCONTRADOS EN CATALOGOS DE EQUIPO', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE EQUIPO NO ENCONTRADOS EN CATALOGOS DE EQUIPO, ';

  if (sMensaje = 'Insumo') or (sMensaje = 'Material') then
    if pos('IDS DE MATERIAL NO ENCONTRADOS EN CATATALOGOS DE MATERIALES/CONSUMIBLES', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE MATERIAL NO ENCONTRADOS EN CATATALOGOS DE MATERIALES/CONSUMIBLES, ';

  if sMensaje = 'Herramienta' then
    if pos('IDS DE HERRAMIENTA NO ENCONTRADOS EN CATALOGOS DE HERRAMIENTA', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE HERRAMIENTA NO ENCONTRADOS EN CATALOGOS DE HERRAMIENTA, ';

  if sMensaje = 'Basico' then
    if pos('IDS DE BASICO NO ENCONTRADOS EN CATALOGOS DE BASICO', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE BASICO NO ENCONTRADOS EN CATALOGOS DE BASICO, ';

  if sMensaje = 'Proveedor' then
    if pos('IDS DE PROVEEDOR NO ENCONTRADOS EN CATALOGOS DE PORVEEDORES', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE PROVEEDOR NO ENCONTRADOS EN CATALOGOS DE PORVEEDORES, ';

  if sMensaje = 'Familia' then
    if pos('IDS DE FAMILIA NO ENCONTRADOS EN CATALOGOS DE FAMILIA DE PRODUCTOS', CodigoColor[iTipo]) = 0 then
      CodigoColor[iTipo] := CodigoColor[iTipo] + 'IDS DE FAMILIA NO ENCONTRADOS EN CATALOGOS DE FAMILIA DE PRODUCTOS, ';

end;
 {$ENDREGION}
{$REGION 'CUADRO CODIGO COLORES'}

procedure TfrmImportaciondeDatos.CuadroColores(sCodigoC1: string; sCodigoC2: string; sErrorC1: string; sErrorC2: string);
var
  Ren: integer;
begin
    {Creacion cuadro colores}
  if (CodigoColor[2] <> '') or (CodigoColor[3] <> '') or (CodigoColor[4] <> '') then
  begin
        {Encabezados}
    ExcelApplication1.Range[sErrorC2 + '5', sErrorC2 + '5'].value := '';
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC1 + '5'].value := '';
        {Textos}
    ExcelApplication1.Range[sErrorC1 + '6', sErrorC1 + '6'].value := '';
    ExcelApplication1.Range[sErrorC1 + '11', sErrorC1 + '11'].value := '';
    ExcelApplication1.Range[sErrorC1 + '16', sErrorC1 + '16'].value := '';
        {Titulos}
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC1 + '5'].ColumnWidth := 8.57;
    ExcelApplication1.Range[sCodigoC2 + '5', sCodigoC2 + '5'].ColumnWidth := 7.45;
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC2 + '5'].ColumnWidth := 10;
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC1 + '5'].Value := 'CODIGO COLORES';
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC2 + '5'].MergeCells := True;
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC2 + '5'].WrapText := True;
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC2 + '5'].font.Color := clWhite;
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC2 + '5'].font.Name := 'Calibri';
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC2 + '5'].font.Size := 11;
    ExcelApplication1.Range[sCodigoC1 + '5', sCodigoC2 + '5'].Interior.ColorIndex := 56;

    ExcelApplication1.Range[sErrorC1 + '5', sErrorC1 + '5'].Value := 'ERRORES ENCONTRADOS';
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC2 + '5'].font.Color := clWhite;
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC2 + '5'].MergeCells := True;
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC2 + '5'].VerticalAlignment := xlHAlignCenter;
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC2 + '5'].HorizontalAlignment := xlHAlignCenter;
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC2 + '5'].WrapText := True;
    ExcelApplication1.Range[sErrorC1 + '5', sErrorC2 + '5'].Interior.ColorIndex := 56;
    ExcelApplication1.Range[sCodigoC1 + '5', sErrorC2 + '5'].borders.LineStyle := 1;

    Ren := 1;
    if CodigoColor[2] <> '' then
    begin
      Ren := 6;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren)].Interior.ColorIndex := 6;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren)].MergeCells := True;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren + 1), sCodigoC2 + IntToStr(Ren + 4)].MergeCells := True;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren + 4)].borders.LineStyle := 1;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].WrapText := True;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].VerticalAlignment := xlHAlignCenter;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].HorizontalAlignment := xlHAlignCenter;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC1 + IntToStr(Ren)].Value := CodigoColor[2];
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].MergeCells := True;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].borders.LineStyle := 1;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].Interior.ColorIndex := 15;
    end;

    if CodigoColor[3] <> '' then
    begin
      Ren := Ren + 5;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren)].Interior.ColorIndex := 3;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren)].MergeCells := True;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren + 1), sCodigoC2 + IntToStr(Ren + 4)].MergeCells := True;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren + 4)].borders.LineStyle := 1;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].WrapText := True;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].VerticalAlignment := xlHAlignCenter;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].HorizontalAlignment := xlHAlignCenter;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC1 + IntToStr(Ren)].Value := CodigoColor[3];
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].MergeCells := True;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].borders.LineStyle := 1;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].Interior.ColorIndex := 15;
    end;

    if CodigoColor[4] <> '' then
    begin
      Ren := Ren + 5;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren)].Interior.ColorIndex := 5;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren)].MergeCells := True;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren + 1), sCodigoC2 + IntToStr(Ren + 4)].MergeCells := True;
      ExcelApplication1.Range[sCodigoC1 + IntToSTr(Ren), sCodigoC2 + IntToStr(Ren + 4)].borders.LineStyle := 1;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].WrapText := True;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].VerticalAlignment := xlHAlignCenter;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].HorizontalAlignment := xlHAlignCenter;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC1 + IntToStr(Ren)].Value := CodigoColor[4];
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].MergeCells := True;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].borders.LineStyle := 1;
      ExcelApplication1.Range[sErrorC1 + IntToSTr(Ren), sErrorC2 + IntToStr(Ren + 4)].Interior.ColorIndex := 15;
    end;
    messageDLG('Se encontraton Datos Erróneos en la Plantilla de Excel!', mtWarning, [mbOk], 0);
  end;
end;
 {$ENDREGION}
{$REGION 'VALIDA TIPOS DE DATOS (TEXTOS, NUMEROS, FECHAS)'}

procedure TfrmImportaciondeDatos.ValidaCampo(sTipo: string; Columna: string; fila: Integer; Campo: string; lFechas: Boolean; sColAnt: string);
var
  Nivel, iDato, i: integer;
  Actividad, Cadena,
    AnexoAux, TipoDato,
    grupo, tipo, frente,
    sValor, sAux, sNivel: string;
  dIdFecha, dIdFechaF: tDate;
  Cantidad, Costos: double;
begin
  try
    if sTipo = 'Numero' then
    begin
      iDato := 3;
           {Datos vacios..}
      try
        cadena := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
        if (trim(Cadena) = '') or (Cadena = 'Null') then
        begin
          sValor := 'nNulo';
          ExcelApplication1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].value := 'Null';

          raise exception.Create('-Excepcion por espacio en blanco..');
        end;
      except
      end;
      sValor := 'nIncorrecto';
      Nivel := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
      Cantidad := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
      if Nivel < 0 then
      begin
        sValor := 'nNegativo';
        raise exception.Create('-Excepcion por numero negativo..');
      end;

      if campo = 'Nivel' then
      begin
        if (Cantidad - Nivel) > 0 then
        begin
          sValor := 'nDecimal';
          raise exception.Create('-Excepcion por numero decimal..');
        end;
              {Validacion de niveles con valores muy altos..}
        if pos(IntToStr(Nivel) + ',', sDatoNivel) = 0 then
        begin
          sAux := sDatoNivel;
          while (sAux <> '') do
          begin
            sNivel := copy(sAux, 0, pos(',', sAux) - 1);
            sAux := copy(sAux, pos(',', sAux) + 1, length(sAux));
          end;

          if sDatoNivel <> '' then
            if (Nivel - StrToInt(sNivel)) >= 2 then
            begin
              sValor := 'nNivel';
              sDatoNivel := sDatoNivel + IntToStr(Nivel) + ',';
              raise exception.Create('-Excepcion por nivel alto..');
            end;

          sDatoNivel := sDatoNivel + IntToStr(Nivel) + ',';
        end;
      end;
    end;

    if sTipo = 'Texto' then
    begin
      iDato := 2;
      if campo = 'Medida' then
      begin
        try
          sValor := 'tMedidaError';
          Nivel := StrToInt(ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2);
          if (Nivel < 0) or (Nivel >= 0) then
          begin
            sValor := 'tMedidaError';
            ColoresErrorExcel(columna + Trim(IntToStr(fila)), columna + Trim(IntToStr(fila)), iDato, sValor);
          end;
        except
                   //No hace nada.. es correcta la uniad de medida..
        end;
      end
      else
      begin
        Actividad := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
        if (trim(Actividad) = '') or (Actividad = 'Null') then
        begin
          sValor := 'tNulo';
          ExcelApplication1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].value := 'Null';
          raise exception.Create('-Excepcion por espacio en blanco..');
        end;
      end;
    end;

    if sTipo = 'Decimal' then
    begin
      iDato := 3;
           {Datos vacios..}
      try
        sValor := 'dIncorrecto';
        cadena := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
        if (trim(Cadena) = '') or (Cadena = 'Null') then
        begin
          sValor := 'dNulo';
          ExcelApplication1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].value := 'Null';
          raise exception.Create('-Excepcion por espacio en blanco..');
        end;
      except
      end;
      Cantidad := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
      if Cantidad < 0 then
      begin
        sValor := 'dNegativo';
        raise exception.Create('-Excepcion por cantidades negativas..');
      end;
    end;

    if sTipo = 'Fecha' then
    begin
      iDato := 4;
           {Datos vacios..}
      try
        sValor := 'fIncorrecto';
        cadena := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
        if (trim(Cadena) = '') or (Cadena = 'Null') then
        begin
          sValor := 'fNulo';
          ExcelApplication1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].value := 'Null';
          raise exception.Create('-Excepcion por espacio en blanco..');
        end;
      except
      end;
      dIdFecha := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
      if dIdFecha = 0 then
      begin
        sValor := 'fIncorrecto';
        raise exception.Create('-Excepcion por fechas nulas..');
      end;

      if lFechas then
      begin
        dIdFecha := ExcelWorksheet1.Range[sColAnt + Trim(IntToStr(Fila)), sColAnt + Trim(IntToStr(Fila))].Value2;
        if dIdFecha = 0 then
          raise exception.Create('-Excepcion por fechas nulas..');

        dIdFechaF := ExcelWorksheet1.Range[columna + Trim(IntToStr(Fila)), columna + Trim(IntToStr(Fila))].Value2;
        if dIdFechaF = 0 then
          raise exception.Create('-Excepcion por fechas nulas..');

              {Validando fechas Finales menores a las de incio..}
        if dIdFechaF < dIdFecha then
        begin
          sValor := 'fMenores';
          ColoresErrorExcel(sColAnt + Trim(IntToStr(fila)), sColAnt + Trim(IntToStr(fila)), iDato, sValor);
          ColoresErrorExcel(columna + Trim(IntToStr(fila)), columna + Trim(IntToStr(fila)), iDato, sValor);
        end;
      end;
    end;
  except
    ColoresErrorExcel(columna + Trim(IntToStr(fila)), columna + Trim(IntToStr(fila)), iDato, sValor);
  end;
end;
 {$ENDREGION}
{$REGION 'ELIMINA CUADRO COLORES'}

procedure TfrmImportaciondeDatos.EliminaCuadro(sPosicion: string; iIndice: Integer);
var
  cadena: string;
begin
  if iIndice = 0 then
    cadena := 'Z'
  else
    cadena := columnas[iIndice + 20];
    {Para no confundir al usuario ponemos todas las celdas en blanco...}
  ExcelApplication1.Range['A2', cadena + '10000'].Interior.ColorIndex := 2;
  ExcelApplication1.Range['A2', cadena + '10000'].font.Color := clBlack;
  ExcelApplication1.Range['A2', cadena + '10000'].MergeCells := False;
    {Quitamos todas las propiedades y datos..}
  ExcelApplication1.Range[sPosicion + '2', cadena + '10000'].borders.LineStyle := 0;
  ExcelApplication1.Range[sPosicion + '2', cadena + '10000'].value := '';
  sDatoNivel := '';
end;
 {$ENDREGION}
{$REGION 'ConstruyeExplosion'}
procedure TfrmImportaciondeDatos.ConstruyeExplosion;
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// GENERA PROGRAMA DE TRABAJO //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      Ren, nivel, i: integer;
      Progreso, TotalProgreso: real;
    begin
      Ren := 2;

      Excel.ActiveWindow.Zoom := 100;

      Excel.Columns['A:F'].ColumnWidth := 20;

      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := 'Actividad';
      FormatoEncabezado;
      Hoja.Range['B1:B1'].Select;
      Excel.Selection.Value := 'Clave';
      FormatoEncabezado;
      Hoja.Range['C1:C1'].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['D1:D1'].Select;
      Excel.Selection.Value := 'Unidad';
      FormatoEncabezado;
      Hoja.Range['E1:E1'].Select;
      Excel.Selection.Value := 'Cantidad';
      FormatoEncabezado;
      Hoja.Range['F1:F1'].Select;
      Excel.Selection.Value := 'Tipo';
      FormatoEncabezado;

      i := 1;
      ProgressBar1.Max := registro;
      while i < registro do begin
        if (recursos[i, 3] <> '') or (recursos[i, 1] <> '') then begin
          Hoja.Cells[Ren, 1].Select;
          Excel.Selection.Value := recursos[i, 1];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 2].Select;
          Excel.Selection.Value := recursos[i, 2];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 3].Select;
          Excel.Selection.Value := recursos[i, 3];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 4].Select;
          Excel.Selection.Value := recursos[i, 4];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 5].Select;
          Excel.Selection.Value := recursos[i, 5];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 6].Select;
          Excel.Selection.Value := recursos[i, 6];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          inc(Ren);
        end;
        inc(i);
        ProgressBar1.Position := ProgressBar1.Position + 1;
      end;
      ProgressBar1.Max := 0;
    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'EXPLOSION DE INSUMOS ';
      except
        Hoja.Name := 'EXPLOSION DE INSUMOS ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
      end;
    end;

    Result := Resultado;
  end;

begin
    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
  if not SaveDialog1.Execute then
    Exit;

    // Generar el ambiente de excel
  try
    Excel := CreateOleObject('Excel.Application');
  except
    FreeAndNil(Excel);
    showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := True;
  end
  else
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := False;
  end;

  Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

    // Verificar si cuenta con las hojas necesarias
  while Libro.Sheets.Count < 2 do
    Libro.Sheets.Add;

    // Verificar si se pasa de hojas necesarias
  Libro.Sheets[1].Select;
  while Libro.Sheets.Count > 1 do
    Excel.ActiveWindow.SelectedSheets.Delete;

    // Proceder a generar la hoja REPORTE
  CadError := '';

  if GenerarPlantilla then
  begin
        // Grabar el archivo de excel con el nombre dado
    messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
  end;

  Excel := '';

  if CadError <> '' then
    showmessage(CadError);
end;
 {$ENDREGION}
{$REGION 'PARTIDAS IGUALES'}

function TfrmImportaciondedatos.PartidasRepetidas(sParamTipo: string): boolean;
var
  i, t, fila, iNivel, x: integer;
  sValue, ImpsContrato, ImpsConvenio, ImpsNumeroActividad,
    ImpsMedida, ImpsAnexo, sTipo, sWbs, ImpsWbsAnterior, ImpsFolio: string;
  paquete: array[1..3000, 1..3] of string;
  lActualiza: boolean;
begin
    //Creamos la tabla temporal.
  PartidasRepetidas := False;
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('CREATE TEMPORARY TABLE IF NOT EXISTS `actividadesxanexo_temp` ( ' +
    '`sContrato` VARCHAR(15) COLLATE latin1_swedish_ci NOT NULL DEFAULT "" COMMENT "Contrato", ' +
    '`sIdConvenio` VARCHAR(5) COLLATE latin1_swedish_ci NOT NULL DEFAULT "" COMMENT "Convenio", ' +
    '`sNumeroOrden` VARCHAR(35) COLLATE latin1_swedish_ci NOT NULL DEFAULT "" COMMENT "sNumeroOrden", ' +
    '`sWbs` VARCHAR(100) COLLATE latin1_swedish_ci NOT NULL DEFAULT "" COMMENT "sWbs", ' +
    '`sNumeroActividad` VARCHAR(20) COLLATE latin1_swedish_ci NOT NULL DEFAULT "" COMMENT "Numero de Actividad", ' +
    '`sTipoActividad` ENUM("Paquete","Actividad") NOT NULL DEFAULT "Actividad" COMMENT "Tipo de Actividad", ' +
    'PRIMARY KEY (`sContrato`, `sIdConvenio`, `sWbs`, `sNumeroOrden`,`sNumeroActividad`, `sTipoActividad`), ' +
    'KEY `actividadesxanexo_fk` (`sIdConvenio`), ' +
    'KEY `sContrato` (`sContrato`, `sWbs`), ' +
    'KEY `sContrato_2` (`sContrato`, `sIdConvenio`, `sWbs`), ' +
    'KEY `sContrato_3` (`sContrato`, `sWbs`, `sNumeroActividad`) ' +
    ')ENGINE=InnoDB ' +
    'CHARACTER SET "latin1" COLLATE "latin1_swedish_ci" ' +
    'COMMENT="Actividades x Anexo"');
  connection.QryBusca.ExecSQL;

  I := 0;
  t := 1;
  Fila := 2;
  sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

    //Recorremos el archivo de Excel
  while (sValue <> '') do
  begin

    ImpsContrato := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;

    inc(I);
    if sParamTipo = 'AnexoC' then
    begin
      iNivel := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
      ImpsNumeroActividad := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
      ImpsMedida := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
      ImpsAnexo := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;
    end
    else
    begin
      iNivel := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
      ImpsNumeroActividad := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
      ImpsMedida := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
      ImpsAnexo := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;
      ImpsConvenio := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
      ImpsFolio    := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
      zq_Esp.Locate('sIdFolio', ImpsFolio, [loCaseInsensitive]);
      ImpsFolio := zq_Esp.FieldByName('sNumeroOrden').AsString;
    end;

    if Trim(ImpsMedida) = '' then
      sTipo := 'Paquete'
    else
      sTipo := 'Actividad';

    sWbs := '';
    if iNivel <> 0 then
    begin
      for x := 1 to t - 1 do
      begin
        if iNivel - 1 >= strToint(paquete[x][1]) then
        begin
          if (sTipo = 'Actividad') and (ImpsAnexo <> '') then
            sWbs := paquete[x][2] + '.' + ImpsAnexo + '.'
          else
            sWbs := paquete[x][2] + '.';
          ImpsWbsAnterior := paquete[x][2];
        end;
      end;
      sWbs := sWbs + ImpsNumeroActividad;
    end
    else
    begin
      ImpsWbsAnterior := '';
      sWbs := ImpsNumeroActividad;
    end;

    if sTipo = 'Paquete' then
    begin
      paquete[t][1] := inttostr(iNivel);
      paquete[t][2] := sWbs;
      t := t + 1;
    end;

        //Intentamos insertar registros,,
    try
      connection.zCommand.Active := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('INSERT INTO actividadesxanexo_temp ( sContrato, sIdConvenio, sTipoActividad, sWbs, sNumeroActividad, sNumeroOrden) ' +
        'VALUES (:contrato, :convenio, :tipo, :wbs, :actividad, :Folio)');
      Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
      Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
      Connection.zCommand.Params.ParamByName('convenio').DataType := ftString;
      if sParamTipo = 'AnexoC' then
         Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio
      else
         Connection.zCommand.Params.ParamByName('convenio').value := ImpsConvenio;
      Connection.zCommand.Params.ParamByName('tipo').DataType := ftString;

      if Trim(ImpsMedida) = '' then
        Connection.zCommand.Params.ParamByName('tipo').value := 'Paquete'
      else
        Connection.zCommand.Params.ParamByName('tipo').value := 'Actividad';
      Connection.zCommand.Params.ParamByName('wbs').DataType := ftString;

      if Trim(ImpsWbsAnterior) <> '' then
        Connection.zCommand.Params.ParamByName('wbs').value := sWbs
      else
        Connection.zCommand.Params.ParamByName('wbs').value := Trim(ImpsNumeroActividad);

      Connection.zCommand.Params.ParamByName('actividad').DataType := ftString;
      Connection.zCommand.Params.ParamByName('actividad').value := Trim(ImpsNumeroActividad);
      Connection.zCommand.Params.ParamByName('Folio').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Folio').value := Trim(ImpsFolio);
      connection.zCommand.ExecSQL;
    except
      lActualiza := False;
      if sParamTipo = 'AnexoC' then
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'O' + Trim(IntToStr(Fila)), 2, 'Iguales')
      else
        ColoresErrorExcel('A' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila)), 2, 'Iguales');
      PartidasRepetidas := True;
    end;

    fila := fila + 1;
    sValue := trim(ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2);
  end;

    //Finalmente borramos la información..
  connection.zCommand.Active := False;
  Connection.zCommand.SQL.Clear;
  Connection.zCommand.SQL.Add('delete from actividadesxanexo_temp where sContrato =:Contrato and sIdConvenio =:Convenio ');
  Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
  Connection.zCommand.Params.ParamByName('contrato').value := ImpsContrato;
  if sParamTipo = 'AnexoC' then
     Connection.zCommand.Params.ParamByName('convenio').value := Global_Convenio
   else
     Connection.zCommand.Params.ParamByName('convenio').value := ImpsConvenio;
  connection.zCommand.ExecSQL;

end;
 procedure TfrmImportaciondeDatos.rAnexoPersonalClick(Sender: TObject);
begin

end;

{$ENDREGION}


end.

