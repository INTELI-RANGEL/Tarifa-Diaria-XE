unit frm_contratos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, DB, ADODB, 
  StdCtrls, DBCtrls, Mask, frm_connection, frm_barra, Global,
  Menus, jpeg, ExtCtrls, ExtDlgs, RXDBCtrl, RxLookup,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, 
  unitexcepciones, udbgrid, unittbotonespermisos, UnitValidaTexto, unitactivapop,
  rxToolEdit, rxCurrEdit;

type
  TfrmContratos = class(TForm)
    grid_contratos: TDBGrid;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    tsContrato: TDBEdit;
    tmComentarios: TDBMemo;
    tmDescripcion: TDBMemo;
    frmBarra1: TfrmBarra;
    tmCliente: TDBMemo;
    OpenPicture: TOpenPictureDialog;
    GroupBox1: TGroupBox;
    bImagen: TImage;
    tlStatus: TDBCheckBox;
    Label19: TLabel;
    tsIdResidencia: TDBLookupComboBox;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    N3: TMenuItem;
    Salir1: TMenuItem;
    Label20: TLabel;
    tsUbicacion: TDBEdit;
    dsActivos: TDataSource;
    tsIdActivo: TRxDBLookupCombo;
    ds_contratos: TDataSource;
    contratos: TZQuery;
    Residencias: TZReadOnlyQuery;
    Activos: TZReadOnlyQuery;
    dsResidencias: TDataSource;
    Label1: TLabel;
    tsCodigo: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    tsLicitacion: TDBEdit;
    tmTitulo: TDBMemo;
    Label8: TLabel;
    tsTipoObra: TDBComboBox;
    chkVigenciaPersonal: TDBCheckBox;
    chkVigenciaEquipo: TDBCheckBox;
    Label10: TLabel;
    Anexos: TZReadOnlyQuery;
    ds_anexos: TDataSource;
    tsAnexo: TDBLookupComboBox;
    Label11: TLabel;
    tsNumeroPOT: TDBEdit;
    Label12: TLabel;
    txtProrrateo: TDBEdit;
    txtCapacidadTripulacion: TRxDBCalcEdit;
    lblCapacidadTripulacion: TLabel;
    chkjorpu: TDBCheckBox;
    Label13: TLabel;
    tsNumeroOrden: TDBEdit;
    DBCheckBox1: TDBCheckBox;
    Label14: TLabel;
    dbObra: TDBComboBox;
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure tsContratoKeyPress(Sender: TObject; var Key: Char);
    procedure tsActivoKeyPress(Sender: TObject; var Key: Char);
    procedure grid_contratosCellClick(Column: TColumn);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Insertar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsContratoEnter(Sender: TObject);
    procedure tsContratoExit(Sender: TObject);
    procedure tmDescripcionEnter(Sender: TObject);
    procedure tmDescripcionExit(Sender: TObject);
    procedure tmClienteEnter(Sender: TObject);
    procedure tmClienteExit(Sender: TObject);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure bImagenClick(Sender: TObject);
    procedure grid_contratosEnter(Sender: TObject);
    procedure grid_contratosKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_contratosKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tsIdResidenciaKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdResidenciaEnter(Sender: TObject);
    procedure tsIdResidenciaExit(Sender: TObject);
    procedure tsUbicacionEnter(Sender: TObject);
    procedure tsUbicacionExit(Sender: TObject);
    procedure tsUbicacionKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdActivoEnter(Sender: TObject);
    procedure tsIdActivoExit(Sender: TObject);
    procedure tsIdActivoKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure tsCodigoEnter(Sender: TObject);
    procedure tsCodigoExit(Sender: TObject);
    procedure tsCodigoKeyPress(Sender: TObject; var Key: Char);
    procedure tsLicitacionEnter(Sender: TObject);
    procedure tsLicitacionExit(Sender: TObject);
    procedure tsLicitacionKeyPress(Sender: TObject; var Key: Char);
    procedure tmTituloEnter(Sender: TObject);
    procedure tmTituloExit(Sender: TObject);
    procedure tmTituloKeyPress(Sender: TObject; var Key: Char);
    procedure tsTipoObraKeyPress(Sender: TObject; var Key: Char);
    procedure tsTipoObraEnter(Sender: TObject);
    procedure tsTipoObraExit(Sender: TObject);
    procedure tsTipoObraChange(Sender: TObject);
    procedure contratosAfterScroll(DataSet: TDataSet);
    procedure tsAnexoEnter(Sender: TObject);
    procedure tsAnexoExit(Sender: TObject);
    procedure tsAnexoKeyPress(Sender: TObject; var Key: Char);
    procedure grid_contratosMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_contratosMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_contratosTitleClick(Column: TColumn);
    procedure tsNumeroPOTEnter(Sender: TObject);
    procedure tsNumeroPOTExit(Sender: TObject);
    procedure ActualizaContrato;
    procedure txtProrrateoEnter(Sender: TObject);
    procedure txtProrrateoExit(Sender: TObject);
    procedure txtCapacidadTripulacionEnter(Sender: TObject);
    procedure txtCapacidadTripulacionExit(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmContratos: TfrmContratos;
  sientra: Byte;
  utgrid: ticdbgrid;
  botonpermiso: tbotonespermisos;
  ContratoActual, ContratoAnterior: string;
implementation

//uses dlg_Contratos;

{$R *.dfm}

procedure TfrmContratos.frmBarra1btnAddClick(Sender: TObject);
var
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;
begin
  activapop(frmContratos, popupprincipal);
  frmBarra1.btnAddClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;
  OpcButton := 'New';

  contratos.Append;
  contratos.FieldValues['mComentarios'] := '*';
  contratos.FieldValues['sCentroGestor'] := '*';
  contratos.FieldValues['sCentroBeneficio'] := '*';
  contratos.FieldValues['sPosicionFinanciera'] := '*';
  contratos.FieldValues['sElementoPEP'] := '*';
  contratos.FieldValues['sCentroCosto'] := '*';
  contratos.FieldValues['sFondo'] := '*';
  contratos.FieldValues['slicitacion'] := '*';
  contratos.FieldValues['stitulo'] := '*';
  contratos.FieldValues['sCuentaMayor'] := '*';
  contratos.FieldValues['sTipoObra'] := 'PROGRAMADA';
  contratos.FieldValues['sPoliza'] := '*';
  contratos.FieldValues['mComentarios'] := '*';
  contratos.FieldValues['sUbicacion'] := '*';
  contratos.FieldValues['lStatus'] := 'Activo';
  contratos.FieldValues['sCodigo'] := global_Contrato_Barco;
  contratos.FieldValues['mCliente'] := '*';
  contratos.FieldValues['mDescripcion'] := '*';
  contratos.FieldValues['sMascara'] := '*';
  contratos.FieldValues['sIdResidencia'] := '02';
  contratos.FieldValues['lCobraPersonal'] := 'No';
  contratos.FieldValues['lCobraEquipo'] := 'No';
  contratos.FieldValues['lJorPu'] := 'No';
  contratos.FieldValues['LmuestraPerRD'] := 'Si';
  //Cargamos por dafualt la imagen del contrato de barco..
  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('select bImagen from contratos where sContrato =:codigo');
  connection.zCommand.ParamByName('codigo').AsString := global_Contrato_Barco;
  connection.zCommand.Open;

  if connection.zCommand.RecordCount > 0 then
  begin
      BlobField := connection.zCommand.FieldByName('bImagen');
      BS := connection.zCommand.CreateBlobStream(BlobField, bmRead);
      if bs.Size > 1 then
      begin
        try
          Pic := TJpegImage.Create;
          try
            bImagen.Picture.Graphic.SaveToStream(bS);
            bImagen.Picture.SaveToFile(global_ruta + 'ImagenContrato_barco.jpg');
          finally
            Pic.Free;
          end;
        finally
          bS.Free
        end
      end
     // contratos.FieldByName('bImagen').AsWideString := connection.zCommand.FieldValues['bImagen'];
  end;

  tlStatus.Checked := True;
  tsContrato.SetFocus;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  tsContrato.Enabled := true;
end;

procedure TfrmContratos.frmBarra1btnEditClick(Sender: TObject);
begin
  frmBarra1.btnEditClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;
  OpcButton := 'Edit';
  try
    activapop(frmContratos, popupprincipal);
    contratos.Edit;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Contratos', 'Al editar registro', 0);
      frmBarra1.btnCancelClick(Sender);
    end;
  end;
  tsContrato.SetFocus;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  contratoAnterior := contratos.fieldvalues['sContrato'];
   // if contratos.FieldValues['sContrato'] = contratos.FieldValues['sCodigo'] then
  tsContrato.Enabled := false;
  MessageDlg('Aqui no podra cambiar el Contrato, solo la informacion adicional.' +
    'Para cambiar el "Contrato" vaya al menu Herramientas > Importacion de datos > (Click en el boton)Cambiar Contrato', mtInformation, [mbOk], 0);
end;

procedure TfrmContratos.frmBarra1btnPostClick(Sender: TObject);
var
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;
  cadena: string;
begin
  {Validacion del campo ID (caracteres validos)}
  if not validaTexto(nil, nil, 'Contrato', tsContrato.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
  {Continua insercion de datos}
  try
      {$REGION 'VALIDACIONES ANTES DE GUARDAR DATOS'}
      if tsContrato.Text = '' then
        cadena := cadena + #13 + '» Contrato';

      if tsCodigo.Text = '' then
        cadena := cadena + #13 + '» Codif. Empresa/Auxiliar';

      if tsIdResidencia.Text = '' then
        cadena := cadena + #13 + '» Residencia';

      if tsTipoObra.Text = '' then
        cadena := cadena + #13 + '» Tipo de Obra';

      if tmDescripcion.Text = '' then
        cadena := cadena + #13 + '» Descripcion';

      if tmCliente.Text = '' then
        cadena := cadena + #13 + '» Cliente';

      if tsIdActivo.Text = '' then
        cadena := cadena + #13 + '» Activo';

      if tsUbicacion.Text = '' then
        cadena := cadena + #13 + '» Ubicacion';

      if tsLicitacion.Text = '' then
        cadena := cadena + #13 + '» Licitacion';

      if cadena <> '' then
      begin
          MessageDlg('Existen Datos Vacíos Favor de Llenar.' + cadena, mtInformation, [mbOk], 0);
          exit;
      end;

      contratoActual := tsContrato.Text;

      if tsAnexo.Enabled then
        contratos.FieldByName('sIdAnexo').Value := tsAnexo.KeyValue
      else
        contratos.FieldByName('sIdAnexo').Value := '';

      if contratos.FieldByName('sIdAnexo').IsNull then
        contratos.FieldByName('sIdAnexo').Value := '';


      if (contratos.FieldValues['bImagen'] = '') or (sientra = 10) then
      begin
        if OpenPicture.FileName <> '' then
        begin
          try
            BlobField := contratos.FieldByName('bImagen');
            BS := contratos.CreateBlobStream(BlobField, bmWrite);
            try
              Pic := TJpegImage.Create;
              try
                Pic.LoadFromFile(OpenPicture.FileName);
                Pic.SaveToStream(Bs);
              finally
                Pic.Free;
              end;
            finally
              bS.Free
            end
          except

          end
        end
        else
        begin
            if connection.zCommand.RecordCount > 0 then
            begin
                try
                  BlobField := contratos.FieldByName('bImagen');
                  BS := contratos.CreateBlobStream(BlobField, bmWrite);
                  try
                    Pic := TJpegImage.Create;
                    try
                      Pic.LoadFromFile(global_ruta + 'ImagenContrato_barco.jpg');
                      Pic.SaveToStream(Bs);
                    finally
                      Pic.Free;
                    end;
                  finally
                    bS.Free
                  end
                except

                end
            end
            else
            begin
                MessageDlg('Agrega una Imagen al Nuevo Contrato ! ', mtInformation, [mbOk], 0);
                exit;
            end;
        end
      end;

      if chkVigenciaPersonal.Checked = True then
         Global_Personal := 'Si'
      else
         Global_Personal := 'No';

      if chkVigenciaEquipo.Checked = True then
         Global_Equipo := 'Si'
      else
         Global_Equipo := 'No';

      if chkJorpu.Checked = True then
         Global_PuJor := 'Si'
      else
         Global_PuJor := 'No';

      {$ENDREGION}

      desactivapop(popupprincipal);
      contratos.Post;

      //Buscamos el contrato de Barco...
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('select scontrato from contratos where sTipoObra = "BARCO" and sContrato = sCodigo ');
      Connection.QryBusca.Open;

      if Connection.QryBusca.RecordCount > 0 then
        global_contrato_barco := Connection.QryBusca.FieldByName('sContrato').AsString
      else
        global_contrato_barco := '';

      {$REGION 'INSERCION DE DATOS BASICOS DEL CONTRATO U ORDEN'}               // JJF by Ivan 3 Nov 2013
      if OpcButton = 'New' then
      begin
          Connection.QryBusca.Active := False;
          Connection.QryBusca.SQL.Clear;
          Connection.QryBusca.SQL.Add('select sContrato from configuracion where sContrato = :contrato');
          Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.QryBusca.Params.ParamByName('Contrato').Value := contratos.FieldValues['sContrato'];
          Connection.QryBusca.Open;
          if Connection.QryBusca.RecordCount = 0 then
          begin
              //Primero validamos si el codigo está ligado al contrato de barco, sino solo insertamos los mov. básicos y configuraciones. JJF by ivan 3 Nov 2013
              if contratos.FieldValues['sCodigo'] = global_Contrato_Barco then
              begin
                  {$REGION 'LLENADO DE DATOS BASICOS SI EL CODIGO = CONTRATO BARCO'}

                  //Consultamos los tipos de movimiento..
                  connection.QryBusca.Active := False;
                  connection.QryBusca.SQL.Clear;
                  connection.QryBusca.SQL.Add('select sNombrecorto, sNombre, sRfc, sDireccion1, sDireccion2, sCiudad, sSlogan, sPiePagina, sTelefono, sFax, sWeb, sEmail, sReportesCIA, '+
                                              'bImagen, bImagenAux1, bImagenAux2, bImagenAux3, bImagenAux4 from configuracion where sContrato =:contrato ');
                  connection.QryBusca.ParamByName('Contrato').AsString := global_Contrato_Barco;
                  connection.QryBusca.Open;

                  if connection.QryBusca.RecordCount > 0 then
                  begin
                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zcommand.SQL.Add('INSERT INTO configuracion ( sContrato, sTipoContrato, sFormato, sNombrecorto, sNombre, sRfc, sDireccion1, sDireccion2, sCiudad, sSlogan, '+
                                                 'sPiePagina, sTelefono, sFax, sWeb, sEmail, sReportesCIA, bImagen, bImagenAux1, bImagenAux2, bImagenAux3, bImagenAux4 ) VALUES '+
                                                 '(:contrato, :tipo, :formato, :Nombrecorto, :Nombre, :Rfc, :Direccion1, :Direccion2, :Ciudad, :Slogan, :PiePagina, :Telefono, :Fax, :Web, :Email, :ReportesCIA, '+
                                                 ':Imagen, :ImagenAux1, :ImagenAux2, :ImagenAux3, :ImagenAux4 )');
                      connection.zCommand.Params.ParamByName('contrato').DataType    := ftString;
                      connection.zCommand.Params.ParamByName('contrato').value       := contratos.FieldValues['sContrato'];
                      connection.zCommand.Params.ParamByName('tipo').DataType        := ftString;
                      connection.zCommand.Params.ParamByName('tipo').value           := 'Precio Unitario';
                      connection.zCommand.Params.ParamByName('formato').DataType     := ftString;
                      connection.zCommand.Params.ParamByName('formato').value        := concat(contratos.FieldValues['sContrato'], '-');
                      connection.zCommand.Params.ParamByName('Nombrecorto').DataType := ftString;
                      connection.zCommand.Params.ParamByName('Nombrecorto').value    := connection.QryBusca.FieldValues['sNombreCorto'];
                      connection.zCommand.Params.ParamByName('Nombre').DataType      := ftString;
                      connection.zCommand.Params.ParamByName('Nombre').value         := connection.QryBusca.FieldValues['sNombre'];
                      connection.zCommand.Params.ParamByName('Rfc').DataType        := ftString;
                      connection.zCommand.Params.ParamByName('Rfc').value           := connection.QryBusca.FieldValues['sRfc'];
                      connection.zCommand.Params.ParamByName('Direccion1').DataType := ftString;
                      connection.zCommand.Params.ParamByName('Direccion1').value    := connection.QryBusca.FieldValues['sDireccion1'];
                      connection.zCommand.Params.ParamByName('Direccion2').DataType := ftString;
                      connection.zCommand.Params.ParamByName('Direccion2').value    := connection.QryBusca.FieldValues['sDireccion2'];
                      connection.zCommand.Params.ParamByName('Ciudad').DataType     := ftString;
                      connection.zCommand.Params.ParamByName('Ciudad').value        := connection.QryBusca.FieldValues['sCiudad'];
                      connection.zCommand.Params.ParamByName('Slogan').DataType     := ftString;
                      connection.zCommand.Params.ParamByName('Slogan').value        := connection.QryBusca.FieldValues['sSlogan'];
                      connection.zCommand.Params.ParamByName('PiePagina').DataType  := ftString;
                      connection.zCommand.Params.ParamByName('PiePagina').value     := connection.QryBusca.FieldValues['sPiePagina'];
                      connection.zCommand.Params.ParamByName('Telefono').DataType   := ftString;
                      connection.zCommand.Params.ParamByName('Telefono').value      := connection.QryBusca.FieldValues['sTelefono'];
                      connection.zCommand.Params.ParamByName('Fax').DataType        := ftString;
                      connection.zCommand.Params.ParamByName('Fax').value           := connection.QryBusca.FieldValues['sFax'];
                      connection.zCommand.Params.ParamByName('Web').DataType        := ftString;
                      connection.zCommand.Params.ParamByName('Web').value           := connection.QryBusca.FieldValues['sWeb'];
                      connection.zCommand.Params.ParamByName('Email').DataType      := ftString;
                      connection.zCommand.Params.ParamByName('Email').value         := connection.QryBusca.FieldValues['sEmail'];
                      connection.zCommand.Params.ParamByName('ReportesCIA').DataType:= ftString;
                      connection.zCommand.Params.ParamByName('ReportesCIA').value   := connection.QryBusca.FieldValues['sReportesCIA'];
                      connection.zCommand.Params.ParamByName('Imagen').DataType     := ftBlob;
                      connection.zCommand.Params.ParamByName('Imagen').value        := connection.QryBusca.FieldValues['bImagen'];
                      connection.zCommand.Params.ParamByName('ImagenAux1').DataType := ftBlob;
                      connection.zCommand.Params.ParamByName('ImagenAux1').value    := connection.QryBusca.FieldValues['bImagenAux1'];
                      connection.zCommand.Params.ParamByName('ImagenAux2').DataType := ftBlob;
                      connection.zCommand.Params.ParamByName('ImagenAux2').value    := connection.QryBusca.FieldValues['bImagenAux2'];
                      connection.zCommand.Params.ParamByName('ImagenAux3').DataType := ftBlob;
                      connection.zCommand.Params.ParamByName('ImagenAux3').value    := connection.QryBusca.FieldValues['bImagenAux3'];
                      connection.zCommand.Params.ParamByName('ImagenAux4').DataType := ftBlob;
                      connection.zCommand.Params.ParamByName('ImagenAux4').value    := connection.QryBusca.FieldValues['bImagenAux4'];
                      connection.zCommand.ExecSQL;
                  end;

                  //Consultamos los tipos de movimiento..
                  connection.QryBusca.Active := False;
                  connection.QryBusca.SQL.Clear;
                  connection.QryBusca.SQL.Add('select * from tiposdemovimiento where sContrato =:contrato and sClasificacion <> "Movimiento de Barco" ');
                  connection.QryBusca.ParamByName('Contrato').AsString := global_Contrato_Barco;
                  connection.QryBusca.Open;

                  while not connection.QryBusca.Eof do
                  begin
                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                        '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                      connection.zCommand.Params.ParamByName('contrato').DataType    := ftString;
                      connection.zCommand.Params.ParamByName('contrato').value       := contratos.FieldValues['sContrato'];
                      connection.zCommand.Params.ParamByName('tipo').DataType        := ftString;
                      connection.zCommand.Params.ParamByName('tipo').value           := connection.QryBusca.FieldValues['sIdTipoMovimiento'];
                      connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                      connection.zCommand.Params.ParamByName('descripcion').value    := connection.QryBusca.FieldValues['sDescripcion'];
                      connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                      connection.zCommand.Params.ParamByName('clasificacion').value  := connection.QryBusca.FieldValues['sClasificacion'];
                      connection.zCommand.Params.ParamByName('orden').DataType       := ftInteger;
                      connection.zCommand.Params.ParamByName('orden').value          := connection.QryBusca.FieldValues['iOrden'];;
                      connection.zCommand.ExecSQL;
                      connection.QryBusca.Next;
                  end;
                  {$ENDREGION}
              end
              else
                  {$REGION 'LLENADO DE DATOS BASICOS SI EL CODIGO <> CONTRATO BARCO'}
              begin
                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO configuracion ( sContrato, sTipoContrato, sFormato ) VALUES (:contrato, :tipo, :formato )');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'Precio Unitario';
                  connection.zCommand.Params.ParamByName('formato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('formato').value := concat(contratos.FieldValues['sContrato'], '-');
                  connection.zCommand.ExecSQL;

                   // Inserta los tipos de movimiento ....
                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                    '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'E';
                  connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('descripcion').value := 'VOLUMEN DE OBRA';
                  connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('clasificacion').value := 'Tiempo en Operacion';
                  connection.zCommand.Params.ParamByName('orden').DataType := ftInteger;
                  connection.zCommand.Params.ParamByName('orden').value := 1;
                  connection.zCommand.ExecSQL;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                    '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'A';
                  connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('descripcion').value := 'ALCANCES POR PARTIDA';
                  connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('clasificacion').value := 'Tiempo en Operacion';
                  connection.zCommand.Params.ParamByName('orden').DataType := ftInteger;
                  connection.zCommand.Params.ParamByName('orden').value := 1;
                  connection.zCommand.ExecSQL;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                    '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'N';
                  connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('descripcion').value := 'COMENTARIOS';
                  connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('clasificacion').value := 'Notas';
                  connection.zCommand.Params.ParamByName('orden').DataType := ftInteger;
                  connection.zCommand.Params.ParamByName('orden').value := 4;
                  connection.zCommand.ExecSQL;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                    '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'AE';
                  connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('descripcion').value := 'RECEPCION DE MATERIALES';
                  connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('clasificacion').value := 'Notas';
                  connection.zCommand.Params.ParamByName('orden').DataType := ftInteger;
                  connection.zCommand.Params.ParamByName('orden').value := 4;
                  connection.zCommand.ExecSQL;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                    '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'M-1';
                  connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('descripcion').value := 'MAL TIEMPO';
                  connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('clasificacion').value := 'Tiempo Muerto';
                  connection.zCommand.Params.ParamByName('orden').DataType := ftInteger;
                  connection.zCommand.Params.ParamByName('orden').value := 3;
                  connection.zCommand.ExecSQL;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                    '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'M-9';
                  connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('descripcion').value := 'OTROS TIEMPOS INACTIVOS';
                  connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('clasificacion').value := 'Tiempo Muerto';
                  connection.zCommand.Params.ParamByName('orden').DataType := ftInteger;
                  connection.zCommand.Params.ParamByName('orden').value := 3;
                  connection.zCommand.ExecSQL;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO tiposdemovimiento ( sContrato, sIdTipoMovimiento, sDescripcion, sClasificacion, iOrden, lGrafica, iColor, dVentaMN, dVentaDLL ) VALUES ' +
                    '(:contrato , :tipo, :descripcion, :clasificacion, :orden, "No", 1, 0, 0)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('tipo').DataType := ftString;
                  connection.zCommand.Params.ParamByName('tipo').value := 'TMDS';
                  connection.zCommand.Params.ParamByName('descripcion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('descripcion').value := 'AJUSTES DE JORNADAS';
                  connection.zCommand.Params.ParamByName('clasificacion').DataType := ftString;
                  connection.zCommand.Params.ParamByName('clasificacion').value := 'Tiempo Muerto';
                  connection.zCommand.Params.ParamByName('orden').DataType := ftInteger;
                  connection.zCommand.Params.ParamByName('orden').value := 5;
                  connection.zCommand.ExecSQL;
              end;

              connection.zCommand.SQL.Clear;
              connection.zcommand.SQL.Add('INSERT INTO turnos ( sContrato, sIdTurno, sDescripcion ) VALUES (:contrato , "A", "UNICO")');
              connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
              connection.zCommand.ExecSQL;

              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zcommand.SQL.Add('INSERT INTO convenios ( sContrato, sIdConvenio, sNumeroOrden, sDescripcion, dFecha, dFechaInicio, dFechaFinal, iGerencialInicio, iGerencialFinal, sHorarioInicio, sHorarioFinal) VALUES ' +
                '(:contrato , "", "", "PLAZO DE EJECUCION CONTRATADO", :fecha, :fechai, :fechaf, 0,0, "00:00", "00:00")');
              connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
              connection.zCommand.Params.ParamByName('fecha').DataType := ftDate;
              connection.zCommand.Params.ParamByName('fecha').value := date;
              connection.zCommand.Params.ParamByName('fechai').DataType := ftDate;
              connection.zCommand.Params.ParamByName('fechai').value := date;
              connection.zCommand.Params.ParamByName('fechaf').DataType := ftDate;
              connection.zCommand.Params.ParamByName('fechaf').value := date;
              connection.zCommand.ExecSQL;

              if global_usuario <> 'INTEL-CODE' then
              begin
                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zcommand.SQL.Add('INSERT INTO contratosxusuario ( sContrato, sIdUsuario ) VALUES (:contrato , :usuario)');
                  connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                  connection.zCommand.Params.ParamByName('usuario').DataType := ftString;
                  connection.zCommand.Params.ParamByName('usuario').value := global_usuario;
                  connection.zCommand.ExecSQL;

                  //Ahora buscamos los usuarios que tengan habilitada la opcion de asignar contratos en autoo..
                  connection.QryBusca2.Active := False;
                  connection.QryBusca2.SQL.Clear;
                  connection.QryBusca2.SQL.Add('select sIdUsuario from usuarios where lAsignaFrentes = "Si"');
                  connection.QryBusca2.Open;

                  if connection.QryBusca2.RecordCount > 0 then
                  begin
                      while not connection.QryBusca2.Eof do
                      begin
                          try
                            //Se inserta el nuevo contrato a los usuarios...
                            connection.zCommand.Active := False;
                            connection.zCommand.SQL.Clear;
                            connection.zcommand.SQL.Add('INSERT INTO contratosxusuario ( sContrato, sIdUsuario ) VALUES ' +
                              '(:contrato , :usuario)');
                            connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                            connection.zCommand.Params.ParamByName('contrato').value := contratos.FieldValues['sContrato'];
                            connection.zCommand.Params.ParamByName('usuario').DataType := ftString;
                            connection.zCommand.Params.ParamByName('usuario').value := connection.QryBusca2.FieldValues['sIdUsuario'];
                            connection.zCommand.ExecSQL;
                          except

                          end;
                          connection.QryBusca2.Next;
                      end;
                  end;
              end;
              {$ENDREGION}
              MessageDlg('Los Datos se Guardaron Correctamente !', mtInformation, [mbOk], 0);
          end
          else
          begin
              messageDLG('El Contrato ya existe. Favor de Verificar', mtInformation, [mbOk], 0);
              exit;
          end;
      end;
      {$ENDREGION}
      Insertar1.Enabled := True;
      Editar1.Enabled := True;
      Registrar1.Enabled := False;
      Can1.Enabled := False;
      Eliminar1.Enabled := True;
      Refresh1.Enabled := True;
      Salir1.Enabled := True;
      frmBarra1.btnPostClick(Sender);
  except
    on e: exception do //cpl>>
    begin
       // MessageDlg('Ocurrio un error al actualizar el registro.', mtInformation, [mbOk], 0);
      //soad -> Si existe el error se procede a eliminar toda la basura creada....

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('delete from configuracion where sContrato = :contrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := contratos.FieldValues['sContrato'];
      Connection.QryBusca.ExecSQL;

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('delete from turnos where sContrato = :contrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := contratos.FieldValues['sContrato'];
      Connection.QryBusca.ExecSQL;

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('delete from turnos where sContrato = :contrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := contratos.FieldValues['sContrato'];
      Connection.QryBusca.ExecSQL;

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('delete from ordenesdetrabajo where sContrato = :contrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := contratos.FieldValues['sContrato'];
      Connection.QryBusca.ExecSQL;

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('delete from tiposdemovimiento where sContrato = :contrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := contratos.FieldValues['sContrato'];
      Connection.QryBusca.ExecSQL;

      //A lo ultimo el contrato el contrato...
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('delete from convenios where sContrato = :contrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := contratos.FieldValues['sContrato'];
      Connection.QryBusca.ExecSQL;

      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Contratos', 'Al salvar registro', 0);
      frmBarra1.btnCancel.Click;

    end;
  end;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  tsContrato.Enabled := true;
end;

procedure TfrmContratos.frmBarra1btnCancelClick(Sender: TObject);
begin
  frmBarra1.btnCancelClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  desactivapop(popupprincipal);
  contratos.Cancel;
  if tsTipoObra.Text = 'PROGRAMADA' then
    tsAnexo.Enabled := true
  else
    tsAnexo.Enabled := false;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  tsContrato.Enabled := true;
end;

procedure TfrmContratos.frmBarra1btnDeleteClick(Sender: TObject);
begin
  if contratos.RecordCount > 0 then
    MessageDlg('No se puede eliminar el contrato, notifique al administrador del sistema.', mtInformation, [mbOk], 0);
end;

procedure TfrmContratos.frmBarra1btnRefreshClick(Sender: TObject);
begin
  Activos.Refresh;
  Residencias.refresh;
  contratos.refresh;
end;

procedure TfrmContratos.frmBarra1btnExitClick(Sender: TObject);
begin
  frmBarra1.btnExitClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  close
end;

procedure TfrmContratos.tsContratoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsCodigo.SetFocus
end;

procedure TfrmContratos.tsActivoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsUbicacion.SetFocus
end;

procedure TfrmContratos.tsAnexoEnter(Sender: TObject);
begin
  tsAnexo.color := global_color_entrada
end;

procedure TfrmContratos.tsAnexoExit(Sender: TObject);
begin
  tsAnexo.color := global_color_salida
end;

procedure TfrmContratos.tsAnexoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tmDescripcion.SetFocus;
end;

procedure TfrmContratos.grid_contratosCellClick(Column: TColumn);
var
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;
begin
  if frmBarra1.btnCancel.Enabled = True then
     frmBarra1.btnCancel.Click;

  if contratos.RecordCount > 0 then
  begin
      if contratos.FieldValues['ljorpu'] = 'Si' then
         chkjorpu.Checked := True
      else
         chkjorpu.Checked := False;

      if contratos.FieldValues['lCobraPersonal'] = 'Si' then
         chkvigenciapersonal.Checked := True
      else
         chkvigenciapersonal.Checked := False;

      if contratos.FieldValues['lCobraEquipo'] = 'Si' then
         chkvigenciaequipo.Checked := True
      else
         chkvigenciaequipo.Checked := False;

      BlobField := contratos.FieldByName('bImagen');
      BS := contratos.CreateBlobStream(BlobField, bmRead);
      if bs.Size > 1 then
      begin
        try
          Pic := TJpegImage.Create;
          try
            Pic.LoadFromStream(bS);
            bImagen.Picture.Graphic := Pic;
          finally
            Pic.Free;
          end;
        finally
          bS.Free
        end
      end
      else
        if fileExists(global_ruta + 'MiImagen.jpg') then
           bImagen.Picture.LoadFromFile(global_ruta + 'MiImagen.jpg')
        else
           bImagen.Picture := nil
  end
end;

procedure TfrmContratos.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'adContratos', PopupPrincipal);
  UtGrid := TicdbGrid.create(grid_contratos);
  sientra := 20;
  contratos.Active := False;
  contratos.SQL.Clear;
  contratos.SQL.Add('Select * From contratos Order By sContrato');
  contratos.Open;

  Activos.Active := False;
  Activos.Active := True;

  Residencias.Active := False;
  Residencias.Open;

  Anexos.Active := False;
  Anexos.Open;

  OpcButton := '';
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  frmBarra1.btnCancel.Click;
  contratos.Refresh;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmContratos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  contratos.Cancel;
  connection.contrato.Active := False;
  connection.contrato.Open;
  action := cafree;
  utgrid.destroy;
  botonpermiso.Free;
end;

procedure TfrmContratos.Insertar1Click(Sender: TObject);
begin
  frmBarra1.btnAdd.Click
end;

procedure TfrmContratos.Editar1Click(Sender: TObject);
begin
  frmBarra1.btnEdit.Click
end;

procedure TfrmContratos.Registrar1Click(Sender: TObject);
begin
  frmBarra1.btnPost.Click
end;

procedure TfrmContratos.Can1Click(Sender: TObject);
begin
  frmBarra1.btnCancel.Click
end;

procedure TfrmContratos.contratosAfterScroll(DataSet: TDataSet);
begin
  if contratos.FieldByName('sTipoObra').AsString = 'PROGRAMADA' then
  begin
    tsAnexo.Enabled := True;
    tsAnexo.KeyValue := contratos.FieldByName('sIdAnexo').AsString;
  end
  else begin
    tsAnexo.Enabled := false;
  end;

  if contratos.FieldByName('sTipoObra').AsString <> 'BARCO' then
  begin
    txtCapacidadTripulacion.Value := 0;
    txtCapacidadTripulacion.Visible := false;
    lblCapacidadTripulacion.Visible := false;
  end
  else
  begin
    txtCapacidadTripulacion.Visible := true;
    lblCapacidadTripulacion.Visible := true;
  end;
end;

procedure TfrmContratos.Eliminar1Click(Sender: TObject);
begin
  frmBarra1.btnDelete.Click
end;

procedure TfrmContratos.Refresh1Click(Sender: TObject);
begin
  frmBarra1.btnRefresh.Click
end;

procedure TfrmContratos.Salir1Click(Sender: TObject);
begin
  frmBarra1.btnExit.Click
end;

procedure TfrmContratos.tsContratoEnter(Sender: TObject);
begin
  tsContrato.Color := global_color_entrada
end;

procedure TfrmContratos.tsContratoExit(Sender: TObject);
begin
  tsContrato.Color := global_color_salida
end;

procedure TfrmContratos.tmDescripcionEnter(Sender: TObject);
begin
  tmDescripcion.Color := global_color_entrada
end;

procedure TfrmContratos.tmDescripcionExit(Sender: TObject);
begin
  tmDescripcion.Color := global_color_salida
end;

procedure TfrmContratos.tmClienteEnter(Sender: TObject);
begin
  tmCliente.Color := global_color_entrada
end;

procedure TfrmContratos.tmClienteExit(Sender: TObject);
begin
  tmCliente.Color := global_color_salida
end;

procedure TfrmContratos.tmComentariosEnter(Sender: TObject);
begin
  tmComentarios.Color := global_color_entrada
end;

procedure TfrmContratos.tmComentariosExit(Sender: TObject);
begin
  tmComentarios.Color := global_color_salida
end;

procedure TfrmContratos.bImagenClick(Sender: TObject);  
begin
  if (contratos.State = dsInsert) or (contratos.State = dsEdit) then
  begin
    OpenPicture.Title := 'Inserta Imagen';
    if OpenPicture.Execute then
    begin
      try
        sientra := 10;
        bImagen.Picture.LoadFromFile(OpenPicture.FileName);
      except
        if fileExists(global_ruta + 'MiImagen.jpg') then
          bImagen.Picture.LoadFromFile(global_ruta + 'MiImagen.jpg')
        else
          bImagen.Picture := nil
      end
    end
  end
end;

procedure TfrmContratos.grid_contratosEnter(Sender: TObject);
var
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;
begin
  if frmBarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if contratos.RecordCount > 0 then
  begin
    BlobField := contratos.FieldByName('bImagen');
    BS := contratos.CreateBlobStream(BlobField, bmRead);
    if bs.Size > 1 then
    begin
      try
        Pic := TJpegImage.Create;
        try
          Pic.LoadFromStream(bS);
          bImagen.Picture.Graphic := Pic;
        finally
          Pic.Free;
        end;
      finally
        bS.Free
      end
    end
    else
      if fileExists(global_ruta + 'MiImagen.jpg') then
        bImagen.Picture.LoadFromFile(global_ruta + 'MiImagen.jpg')
      else
        bImagen.Picture := nil
  end
end;

procedure TfrmContratos.grid_contratosKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
var
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;
begin
  if frmBarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if contratos.RecordCount > 0 then
  begin
    BlobField := contratos.FieldByName('bImagen');
    BS := contratos.CreateBlobStream(BlobField, bmRead);
    if bs.Size > 1 then
    begin
      try
        Pic := TJpegImage.Create;
        try
          Pic.LoadFromStream(bS);
          bImagen.Picture.Graphic := Pic;
        finally
          Pic.Free;
        end;
      finally
        bS.Free
      end
    end
    else
      if fileExists(global_ruta + 'MiImagen.jpg') then
        bImagen.Picture.LoadFromFile(global_ruta + 'MiImagen.jpg')
      else
        bImagen.Picture := nil
  end
end;

procedure TfrmContratos.grid_contratosKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;
begin
  if frmBarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if contratos.RecordCount > 0 then
  begin
    BlobField := contratos.FieldByName('bImagen');
    BS := contratos.CreateBlobStream(BlobField, bmRead);
    if bs.Size > 1 then
    begin
      try
        Pic := TJpegImage.Create;
        try
          Pic.LoadFromStream(bS);
          bImagen.Picture.Graphic := Pic;
        finally
          Pic.Free;
        end;
      finally
        bS.Free
      end
    end
    else
      if fileExists(global_ruta + 'MiImagen.jpg') then
        bImagen.Picture.LoadFromFile(global_ruta + 'MiImagen.jpg')
      else
        bImagen.Picture := nil
  end
end;

procedure TfrmContratos.grid_contratosMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmContratos.grid_contratosMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmContratos.grid_contratosTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmContratos.tsIdResidenciaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tmdescripcion.SetFocus
end;

procedure TfrmContratos.tsIdResidenciaEnter(Sender: TObject);
begin
  tsIdResidencia.Color := global_color_entrada
end;

procedure TfrmContratos.tsIdResidenciaExit(Sender: TObject);
begin
  tsIdResidencia.Color := global_color_salida
end;

procedure TfrmContratos.tsUbicacionEnter(Sender: TObject);
begin
  tsUbicacion.color := global_color_entrada
end;

procedure TfrmContratos.tsUbicacionExit(Sender: TObject);
begin
  tsUbicacion.color := global_color_salida
end;

procedure TfrmContratos.tsUbicacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    TSLICITACION.SetFocus
end;

procedure TfrmContratos.txtCapacidadTripulacionEnter(Sender: TObject);
begin
  txtCapacidadTripulacion.Color := global_color_entrada
end;

procedure TfrmContratos.txtCapacidadTripulacionExit(Sender: TObject);
begin
  txtCapacidadTripulacion.Color := global_color_salida
end;

procedure TfrmContratos.txtProrrateoEnter(Sender: TObject);
begin
  txtProrrateo.Color := global_color_entrada
end;

procedure TfrmContratos.txtProrrateoExit(Sender: TObject);
begin
  txtProrrateo.Color := global_color_salida
end;

procedure TfrmContratos.tsIdActivoEnter(Sender: TObject);
begin
  tsIdActivo.Color := global_color_entrada
end;

procedure TfrmContratos.tsIdActivoExit(Sender: TObject);
begin
  tsIdActivo.Color := global_color_salida
end;

procedure TfrmContratos.tsIdActivoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsUbicacion.SetFocus
end;

procedure TfrmContratos.tdFechaFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tmComentarios.SetFocus
end;

procedure TfrmContratos.tsCodigoEnter(Sender: TObject);
begin
  tsCodigo.Color := global_color_entrada
end;

procedure TfrmContratos.tsCodigoExit(Sender: TObject);
begin
  tsCodigo.Color := global_color_salida
end;

procedure TfrmContratos.tsCodigoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tstipoobra.SetFocus
end;

procedure TfrmContratos.tsLicitacionEnter(Sender: TObject);
begin
  tslicitacion.color := global_color_entrada
end;

procedure TfrmContratos.tsLicitacionExit(Sender: TObject);
begin
  tslicitacion.color := global_color_salida
end;

procedure TfrmContratos.tsLicitacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tmTitulo.SetFocus;

end;


procedure TfrmContratos.tsNumeroPOTEnter(Sender: TObject);
begin
  tsNumeroPOT.color := global_color_entrada
end;

procedure TfrmContratos.tsNumeroPOTExit(Sender: TObject);
begin
  tsNumeroPOT.color := global_color_salida
end;

procedure TfrmContratos.tsTipoObraChange(Sender: TObject);
begin
  if tsTipoObra.Text = 'PROGRAMADA' then
  begin
    tsAnexo.Enabled := true;
    if Anexos.RecordCount > 0 then
      tsAnexo.KeyValue := Anexos.FieldByName('sAnexo').AsString
  end
  else begin
    tsAnexo.Enabled := false;
  end;
end;

procedure TfrmContratos.tsTipoObraEnter(Sender: TObject);
begin
  tsTipoObra.color := global_color_entrada
end;

procedure TfrmContratos.tsTipoObraExit(Sender: TObject);
begin
  tsTipoObra.color := global_color_salida
end;

procedure TfrmContratos.tsTipoObraKeyPress(Sender: TObject; var Key: Char);
begin
//  If Key = #13 then
//    if tsAnexo.Visible then
//      tsAnexo.SetFocus
//    else
//      tmDescripcion.SetFocus
  if Key = #13 then
    tsidresidencia.SetFocus
end;

procedure TfrmContratos.tmTituloEnter(Sender: TObject);
begin
  tmTitulo.Color := global_color_entrada
end;

procedure TfrmContratos.tmTituloExit(Sender: TObject);
begin
  tmTitulo.Color := global_color_salida;
end;

procedure TfrmContratos.tmTituloKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tmComentarios.SetFocus;
end;

procedure TfrmContratos.ActualizaContrato;
var
  base, tabla, campo, cad: string;
  datos: array[1..300] of string;
  i, x: Integer;
begin
  connection.qryBusca.Active := False;
  connection.qryBusca.SQL.Clear;
  connection.qryBusca.SQL.Add('Show tables');
  connection.qryBusca.Open;
  base := 'Tables_in_' + global_db;
  i := 1;
  while not connection.QryBusca.Eof do
  begin
    tabla := connection.QryBusca.FieldValues[base];
    connection.qryBusca2.Active := False;
    connection.qryBusca2.SQL.Clear;
    connection.qryBusca2.SQL.Add('describe ' + tabla + ' ');
    connection.qryBusca2.Open;

    if connection.QryBusca2.RecordCount > 0 then
    begin
      while not connection.QryBusca2.Eof do
      begin
        if connection.QryBusca2.FieldValues['Field'] <> 'sNumeroOrden' then
        begin
          if connection.QryBusca2.FieldValues['Field'] = 'sContrato' then
          begin
            datos[i] := tabla;
            i := i + 1;
          end;
        end;
        connection.QryBusca2.Next;
      end;
    end;
    connection.QryBusca.Next;
  end;

     // Actualiza todos los registros..
  if connection.QryBusca.RecordCount > 0 then
  begin
    for x := 1 to i - 1 do
    begin
      tabla := datos[x];
      connection.qryBusca.Active := False;
      connection.qryBusca.SQL.Clear;
      connection.qryBusca.SQL.Add('update ' + tabla + ' set sContrato = :Nuevo where sContrato = :Contrato ');
      connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      connection.qryBusca.Params.ParamByName('Contrato').Value := ContratoAnterior;
      connection.qryBusca.Params.ParamByName('Nuevo').DataType := ftString;
      connection.qryBusca.Params.ParamByName('Nuevo').Value := ContratoActual;
      connection.qryBusca.ExecSQL;
    end;
  end;
  messageDLG('Proceso Terminado con Exito', mtInformation, [mbOk], 0);
end;

end.

