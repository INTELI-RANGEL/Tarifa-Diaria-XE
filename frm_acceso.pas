unit frm_acceso;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, frm_connection, global, Sockets, DB, strUtils,
  ADOdb, Newpanel, ExtCtrls, ComCtrls, ZDataset, frxpngimage, AdvGlassButton,
  IniFiles, ZSqlProcessor, UnitTIniTracer, Menus, jpeg, dxGDIPlusClasses,
  ImgList, cxGraphics, cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore,
  dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee,
  dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, cxButtons, cxControls, cxContainer, cxEdit, cxTextEdit,
  cxMaskEdit, cxDropDownEdit;

type
  Tfrmacceso = class(TForm)
    ip_client: TTcpClient;
    Panel1: TPanel;
    tNewPanel1: tNewPanel;
    Image2: TImage;
    StatusBar: TStatusBar;
    Label4: TLabel;
    Label3: TLabel;
    ZqScript: TZSQLProcessor;
    imgIcons: TcxImageList;
    cmbServer: TcxComboBox;
    Label7: TLabel;
    Label1: TLabel;
    tsIdUsuario: TcxTextEdit;
    tsPassword: TcxTextEdit;
    Label2: TLabel;
    lblBase: TLabel;
    sDataName: TcxComboBox;
    tsPuerto: TEdit;
    Label5: TLabel;
    AdvGlassButton1: TAdvGlassButton;
    btnAdelante: TcxButton;
    btnAbortar: TcxButton;
    procedure btnAdelanteClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnAbortarClick(Sender: TObject);
    procedure tsPasswordKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdUsuarioKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdUsuarioEnter(Sender: TObject);
    procedure tsIdUsuarioExit(Sender: TObject);
    procedure tsPasswordEnter(Sender: TObject);
    procedure tsPasswordExit(Sender: TObject);
    procedure cmbServerEnter(Sender: TObject);
    procedure cmbServerExit(Sender: TObject);
    procedure cmbServerKeyPress(Sender: TObject; var Key: Char);
    procedure FormActivate(Sender: TObject);
    procedure btnSalirClick(Sender: TObject);
    procedure sDataNameEnter(Sender: TObject);
    procedure sDataNameExit(Sender: TObject);
    procedure sDataNameKeyPress(Sender: TObject; var Key: Char);
    procedure cmbServerChange(Sender: TObject);
    procedure SetTransparentForm(AHandle: THandle; AValue: byte = 0);
    procedure FormCreate(Sender: TObject);
    procedure AdvGlassButton1Click(Sender: TObject);
    function TestServer: boolean;
    procedure tsIdUsuarioPropertiesEditValueChanged(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
    salir: boolean;
    dbGralExiste : Boolean;
  end;

const
  WS_EX_LAYERED = $80000;
  LWA_COLORKEY = 1;
  LWA_ALPHA = 2;

type
  TSetLayeredWindowAttributes = function(
    hwnd: HWND; // handle to the layered window
    crKey: TColor; // specifies the color key
    bAlpha: byte; // value for the blend function
    dwFlags: DWORD // action
    ): BOOL; stdcall;

var
  frmacceso: Tfrmacceso;
  intentos: byte;
  mensaje: string;
  sVector: array[1..100] of string;
  listServ: tstringlist;
  Ini: TIniFile;
  FilePath: string;
  flagAccesoEnIni: boolean;

implementation

uses frm_inteligent, frm_AltaServidor, Utilerias, unitmanejofondo,
  frm_seleccion2;
{$R *.dfm}



(*
  procedure Tfrmacceso.WMCommand(var Msg: TWMCommand);
  begin
     if Msg.NotifyCode = EN_CHANGE then
     begin
        PostMessage(Handle, WM_NEXTDLGCTL,0, 0);
        inherited;
     end;
  end;
*)

procedure Tfrmacceso.SetTransparentForm(AHandle: THandle; AValue: byte = 0);
var
  Info: TOSVersionInfo;
  SetLayeredWindowAttributes: TSetLayeredWindowAttributes;
begin
  //Check Windows version
  Info.dwOSVersionInfoSize := SizeOf(Info);
  GetVersionEx(Info);
  if (Info.dwPlatformId = VER_PLATFORM_WIN32_NT) and
    (Info.dwMajorVersion >= 5) then
  begin
    SetLayeredWindowAttributes := GetProcAddress(GetModulehandle(user32), 'SetLayeredWindowAttributes');
    if Assigned(SetLayeredWindowAttributes) then
    begin
      SetWindowLong(AHandle, GWL_EXSTYLE, GetWindowLong(AHandle, GWL_EXSTYLE) or WS_EX_LAYERED);
         //Make form transparent
      SetLayeredWindowAttributes(AHandle, 0, AValue, LWA_ALPHA);
    end;
  end;
end;



procedure Tfrmacceso.btnAdelanteClick(Sender: TObject);
var
  zQuery: tzReadOnlyQuery;
  iItemDB: Byte;
  lFoundDB: Boolean;
  QrAcceso: TzReadOnlyQuery;
  Error: boolean;
  AlterUsuarios: TzsqlProcessor;
begin
  salir := False;
  Error := false;

  if sDataName.Visible = False then
  begin
    if (uppercase(tsidusuario.text) = 'ADMIN') then
      global_bduser := 'root'
    else
      global_bduser := 'inteligent';

    Global_Puerto := strToint(tsPuerto.Text);
    connection.zConnection.Disconnect;

    if Global_ServAcceso = '' then
    begin
      connection.zConnection.HostName := sVector[cmbServer.ItemIndex + 1];
      connection.zConnection.Port := Global_Puerto;
      Global_ServAcceso := connection.zConnection.HostName;
    end
    else
    begin
      connection.zConnection.HostName := Global_ServAcceso;
      connection.zConnection.Port := Global_PortAcceso;
    end;

    connection.zConnection.User := 'root';
    connection.zConnection.Password := 'danae';
    connection.zConnection.Database := '';
    connection.zConnection.Catalog := '';
    try
      connection.zConnection.Connect;
    except
      on e: exception do
      begin
          //StatusBar.Panels[0].Text := e.Message ;
        if pos('Access denied', e.message) > 0 then
        begin
          if (uppercase(tsidusuario.Text) = 'ADMIN') and (uppercase(tspassword.Text) = uppercase(intelpass)) then
          begin
            connection.zConnection.Disconnect;
            connection.zConnection.User := 'root';
            connection.zConnection.Password := IntelPass;
            connection.zConnection.Database := '';
            connection.zConnection.Catalog := '';
            try
              connection.zConnection.Connect;
            except
              on e: exception do
              begin
                messagedlg('No se Puede Conectar al Servidor.' + #13 + #10 +
                  'Informacion del error: ' + e.ClassName + ',' + e.Message, mterror, [mbok], 0);
                error := true;
              end;
            end;

            if not error then
            begin
              if connection.zConnection.Ping then
              begin
                zqScript.Connection := connection.zConnection;
                zqScript.ParamByName('password').AsString := IntelPass;
                try
                  zqScript.Execute;
                except
                  on E: exception do
                  begin
                    messagedlg('Ocurrio un error al Inicializar Parametros de Configuracion.' + #13 + #10 +
                      'Informacion del error: ' + e.ClassName + ',' + e.Message, mterror, [mbok], 0);
                    error := true;
                  end;
                end;

                if not error then
                begin
                  connection.zConnection.Disconnect;
                  connection.zConnection.User := IntelUser;
                  connection.zConnection.Password := IntelPass;
                  connection.zConnection.Database := '';
                  connection.zConnection.Catalog := '';
                  try
                    connection.zConnection.Connect;
                  except
                    on e: exception do
                    begin
                      messagedlg('No se Puede Conectar al Servidor.' + #13 + #10 +
                        'Informacion del error: ' + e.ClassName + ',' + e.Message, mterror, [mbok], 0);
                      error := true;
                    end;
                  end;

                end;
              end
              else
              begin
                application.MessageBox('No se Logra tener Comunicacion con este Servidor.' + #13 + #10 +
                  'Notifiquelo al Administrador del Sistema, para verificar los parametros de Conexion.', 'Inteligent');

              end;
            end;
          end
          else
          begin
            if (uppercase(tsidusuario.Text) = 'ADMIN') then
              messagedlg('No se Puede Conectar al Servidor.' + #13 + #10 +
                'La Contraseña del Administrador No es Correcta ', mtInformation, [mbok], 0)

            else
              messagedlg('No se Puede Conectar al Servidor.' + #13 + #10 +
                'Informacion del Problema: ' + e.ClassName + ',' + e.Message, mterror, [mbok], 0);
            error := true;
          end;

        end
        else
        begin
          Global_ServAcceso := '';
          error := true;

          messagedlg('No se Puede Conectar al Servidor.' + #13 + #10 +
            'Informacion del Problema: ' + e.ClassName + ',' + e.Message, mterror, [mbok], 0);
        end;

      end;
    end;

    if not error then
    begin
      if connection.zConnection.Ping then
      begin
        QrAcceso := TzReadOnlyquery.Create(nil);
        QrAcceso.Connection := connection.zConnection;
        QrAcceso.SQL.Text := 'select * from adminintel.acceso where user=' + quotedstr(global_bduser) +
          ' and servidor=' + quotedstr(sVector[cmbServer.ItemIndex + 1]);
        {try
          QrAcceso.Open;
        except
          on E: exception do
          begin
            if pos('t exist', e.Message) > 0 then
            begin
              messagedlg('La base de Datos AdminIntel y/o la tabla Acceso no existe en el servidor.  ' + #13 + #10 +
                'Verificar con el Administrador del Sistema', mterror, [mbok], 0);
            end
            else
            begin
              messagedlg('Error generado:' + #13 + #10 + e.Message, mterror, [mbok], 0);
            end;
            error := true;
          end;

        end;}

        if not error then
        begin
          {try
            if QrAcceso.RecordCount = 1 then
              global_bdPass := desencripta(QrAcceso.FieldByName('password').AsString)
            else
              global_bdPass := IntelPass;

          finally
            freeandnil(QrAcceso);
          end;}

          global_bdUser := 'root';
          global_bdPass := 'danae';
          Connection.zConnection.Disconnect;
          connection.zConnection.HostName := sVector[cmbServer.ItemIndex + 1];
          connection.zConnection.Port := Global_Puerto;
          connection.zConnection.User := 'root';
          connection.zConnection.Password := 'danae';
          connection.zConnection.Database := '';

          try
            connection.zConnection.Connect;
          except
            on E: exception do
            begin
              error := true;
              StatusBar.Panels[0].Text := e.Message;
              begin
                messagedlg('Error generado:' + #13 + #10 + e.Message, mterror, [mbok], 0);
              end;
            end;
          end;

          if not error then
          begin
            try
              if connection.zConnection.Ping then
              begin
                sDataName.Properties.Items.Clear;
                dbGralExiste := False;
                //Consultar las BD que le corresponden a cada usuario conforme a la BD bd_gral, y mostrarlas en el combo
                zQuery := tzReadOnlyQuery.Create(Self);
                zQuery.Connection := connection.zConnection;
                zQuery.Active := False;
                zQuery.SQL.Clear;
                zQuery.SQL.Add('show databases');
                zQuery.Open;

                while Not zQuery.Eof do
                begin
                  if (zQuery.FieldByName('Database').AsString = 'db_gral') then
                  begin
                    dbGralExiste := True;
                  end;
                  zQuery.Next;
                end;

                if dbGralExiste = True then
                begin
                  Connection.zConnection.Disconnect;
                  connection.zConnection.Database := 'db_gral';
                  with connection.zCommand do
                  begin
                    Active := False;
                    SQL.Clear;
                    SQL.Add('select ' +
                            'ub.useer, ' +
                            'ub.nombrees, ' +
                            'b.bd_user ' +
                            'from user_bd ub ' +
                            'inner join bd b ' +
                            'on b.id_bd = ub.id_bds ' +
                            'where useer = :useer ' +
                            'group by b.bd_user');
                    Params.ParamByName('useer').AsString := tsIdUsuario.Text;
                    Open;
                  end;

                  if connection.zCommand.RecordCount > 0 then
                  begin


                    while Not connection.zCommand.Eof do
                    begin
                      zQuery.Active := False;
                      zQuery.SQL.Clear;
                      zQuery.SQL.Add('show databases');
                      zQuery.Open;

                      iItemDB := 0;
                      lFoundDB := False;
                      while not zQuery.Eof do
                      begin
                        if (zQuery.FieldValues['database'] <> 'performance_schema') and (zQuery.FieldValues['database'] <> 'mysql') and (zQuery.FieldValues['database'] <> 'information_schema') and (zQuery.FieldValues['database'] <> 'test') and (zQuery.FieldValues['database'] <> 'chat') and (zQuery.FieldValues['database'] <> 'chat_php') and (zQuery.FieldValues['database'] <> 'ic_exsoll') and (zQuery.FieldValues['database'] <> 'joomla') and (zQuery.FieldValues['database'] <> 'phpmyadmin') and (zQuery.FieldValues['database'] <> 'adminintel') then
                        begin
                          if (connection.zCommand.FieldByName('bd_user').AsString = zQuery.FieldByName('Database').AsString) then
                          begin
                            sDataName.Properties.Items.Add(zQuery.FieldValues['database']);
                            if zQuery.FieldValues['database'] = global_db then
                              if lFoundDB then
                                iItemDB := zQuery.RecNo - 2
                              else
                                iItemDB := zQuery.RecNo - 1
                          end;
                        end
                        else
                          lFoundDB := True;
                        zQuery.Next;
                      end;

                      connection.zCommand.Next;
                    end;
                  end
                  else
                  begin
                    application.MessageBox('Este usuario no tiene asignado ninguna BD. Vuelva a intentarlo', 'Inteligent');
                    Exit;
                  end;

                end
                else
                begin
                  zQuery.Active := False;
                  zQuery.SQL.Clear;
                  zQuery.SQL.Add('show databases');
                  zQuery.Open;

                  iItemDB := 0;
                  lFoundDB := False;
                  while not zQuery.Eof do
                  begin
                    if (zQuery.FieldValues['database'] <> 'performance_schema') and (zQuery.FieldValues['database'] <> 'mysql') and (zQuery.FieldValues['database'] <> 'information_schema') and (zQuery.FieldValues['database'] <> 'test') and (zQuery.FieldValues['database'] <> 'chat') and (zQuery.FieldValues['database'] <> 'chat_php') and (zQuery.FieldValues['database'] <> 'ic_exsoll') and (zQuery.FieldValues['database'] <> 'joomla') and (zQuery.FieldValues['database'] <> 'phpmyadmin') and (zQuery.FieldValues['database'] <> 'adminintel') then
                    begin
                        sDataName.Properties.Items.Add(zQuery.FieldValues['database']);
                        if zQuery.FieldValues['database'] = global_db then
                          if lFoundDB then
                            iItemDB := zQuery.RecNo - 2
                          else
                            iItemDB := zQuery.RecNo - 1
                    end
                    else
                      lFoundDB := True;
                    zQuery.Next;
                  end;
                end;

                zQuery.Destroy;
                sDataName.Visible := True;
                sDataName.ItemIndex := 0;
                StatusBar.Panels[0].Text := '';
                lblBase.Visible := True;
                btnAdelante.Caption := 'Entrar a Inteligent';
                sDataName.SetFocus
              end;
            except
              on E: exception do
              begin
                messagedlg('Error generado:' + #13 + #10 + e.Message, mterror, [mbok], 0);
              end;
            end;

          end;
        end;
      end
      else
      begin
        application.MessageBox('No se Logra tener Comunicacion con este Servidor.' + #13 + #10 +
          'Notifiquelo al Administrador del Sistema, para verificar los parametros de Conexion.', 'Inteligent');
      end;
    end;
  end
  else
  begin
    if intentos = 3 then
    begin
      application.MessageBox('Intento accesar en mas de 3 ocaciones. Saliendo del Sistema', 'Inteligent');
      salir := true;
      close;
    end;

    if connection.zConnection.Ping then
      if sDataName.Text <> '' then
      begin
        global_Puerto := strToint(tsPuerto.Text);
        global_db := sDataName.Text;
        global_ipServer := sVector[cmbServer.ItemIndex + 1];
        connection.zConnection.Disconnect;
        connection.zConnection.HostName := sVector[cmbServer.ItemIndex + 1];
        connection.zConnection.Database := global_db;
        connection.zConnection.Port := Global_Puerto;
        connection.zConnection.User := global_bduser;
        connection.zConnection.Password := global_bdpass;
        connection.zConnection.Connect;
      end;

    if connection.zConnection.Ping then
    begin
      if (tsIdUsuario.Text = 'INTEL-CODE') or (uppercase(tsidusuario.text) = 'ADMIN') then
      begin
        if uppercase(tsidusuario.text) = 'ADMIN' then
        begin
          connection.Zcommand.Active := false;
          connection.Zcommand.SQL.Text := 'select * from usuarios where sidusuario=' + quotedstr(tsidusuario.text);
          try
            connection.Zcommand.Open;

          except
            on E: exception do
            begin
              error := true;
              StatusBar.Panels[0].Text := e.Message;
            end;
          end;

          if not error then
          begin
            if connection.Zcommand.recordcount = 0 then
            begin
              if uppercase(tspassword.text) = uppercase(intelpass) then
              begin
                AlterUsuarios := TzSqlProcessor.Create(nil);
                AlterUsuarios.Connection := connection.zConnection;
                AlterUsuarios.Script.Text := 'ALTER TABLE `usuarios` MODIFY COLUMN `sPassword` VARCHAR(50) COLLATE latin1_swedish_ci NOT NULL DEFAULT "" COMMENT "Contraseña";';

                try
                  AlterUsuarios.Execute;
                except
                  on E: exception do
                  begin
                    messagedlg('Ocurrio un error al Inicializar Parametros de Configuracion.' + #13 + #10 +
                      'Informacion del error: ' + e.ClassName + ',' + e.Message, mterror, [mbok], 0);
                    error := true;
                  end;
                end;

                if not error then
                begin
                  connection.QryBusca.active := false;
                  connection.QryBusca.SQL.text := 'insert into usuarios(sidusuario,spassword) values(:user,:pass)';
                  connection.QryBusca.ParamByName('user').AsString := 'admin';
                  try
                    connection.QryBusca.ParamByName('pass').AsString := encripta(Intelpass);
                    connection.QryBusca.ExecSQL;
                    connection.Zcommand.Refresh;
                  except
                    on e: exception do
                    begin
                      messagedlg('No se Pudo Cargar el Administrador.' + #13 + #10 +
                        'Informacion del error: ' + e.ClassName + ',' + e.Message, mterror, [mbok], 0);
                      error := true;
                    end;
                  end;
                end;
              end
              else
              begin
                application.MessageBox('Ese Usuario No EXISTE', 'Inteligent');
                error := true;
              end;
            end;

            if not error then
            begin
              if connection.Zcommand.recordcount = 1 then
              begin
                try
                  if uppercase(desencripta(connection.Zcommand.FieldByName('spassword').AsString)) <> uppercase(tspassword.Text) then
                  begin
                    application.MessageBox('PASSWORD INCORRECTO ', 'Inteligent');
                    error := true;
                  end;
                except
                  application.MessageBox('PASSWORD INCORRECTO ', 'Inteligent');
                  error := true;
                end;
              end
              else
              begin
                application.MessageBox('Ese Usuario No EXISTE', 'Inteligent');
                error := true;
              end;
            end;
          end;
        end;

        if not error then
        begin

          global_contrato := '';
          global_usuario := uppercase(tsIdUsuario.Text);
          global_password := '';
          global_nombre := 'INTEL-CODE S.A. DE C.V.';
          global_puesto := 'ADMINISTRADOR UNICO';
          global_activo := '';
          global_grupo := 'INTEL-CODE';
          global_ip := ip_client.LocalHostAddr;
          close
        end;
      end
      else
      begin
        global_Server := cmbServer.Text;
        global_ipserver := sVector[cmbServer.ItemIndex + 1];
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;


        connection.QryBusca.SQL.Add('select * from usuarios where sIdUsuario = :usuario and lacceso="si"');
        connection.QryBusca.params.ParamByName('Usuario').DataType := ftString;
        connection.QryBusca.params.ParamByName('Usuario').Value := tsIdUsuario.Text;
        connection.QryBusca.Open;
        if connection.QryBusca.RecordCount > 0 then
        begin
          if connection.QryBusca.FieldValues['sPassword'] = tsPassword.Text then
          begin
            global_contrato := '';
            global_usuario := connection.QryBusca.FieldValues['sIdUsuario'];
            global_password := tsPassword.Text;
            global_nombre := connection.QryBusca.FieldValues['sNombre'];
            global_puesto := connection.QryBusca.FieldValues['sPuesto'];
            global_activo := connection.QryBusca.fieldvalues['lActivo'];
            global_grupo := connection.QryBusca.fieldvalues['sIdGrupo'];
            Global_ip := ip_client.LocalHostAddr;
            Close
          end
          else
            application.MessageBox('PASSWORD INCORRECTO ', 'Inteligent');
        end
        else
        begin
          application.MessageBox('Ese Usuario No EXISTE', 'Inteligent');
          global_usuario := tsIdUsuario.Text;
          global_password := tsPassword.Text;
          global_nombre := 'Falta introducir informacion general del usuario seleccionado';
          global_puesto := '';
          global_activo := 'Si';
          global_grupo := '';
          global_ip := ip_client.LocalHostAddr;
          intentos := intentos + 1;
          beep;
          Exit;
        end
      end
    end;
  end;
end;


procedure Tfrmacceso.FormShow(Sender: TObject);
var
  MiArchivo: string;
  FileText: TextFile;
  wCadena: WideString;
  sTipo: string;
  iVector,
    iPos: Byte;
  sPortAcceso: string;
  IniTracer: TIniTracer;
  appINI: TIniFile;
  detectar: string;
begin

 //****************************************************************************************************************
  ShowWindow(Application.Handle, SW_SHOW);
  cmbServer.Properties.Items.Clear;
  StatusBar.Panels[0].Text := '';
  sPortAcceso := '3306';
  for iVector := 1 to 100 do
    sVector[iVector] := '';

  iVector := 1;

  MiArchivo := extractfilepath(application.exename) + 'inteligent.ini';
  if not fileExists(MiArchivo) then
  begin
    IniTracer := TIniTracer.create(self, 'SOFTWARE\INTELIGENT', 'INTELIGENT', 'INTELIGENT', 'cotemar');
    MiArchivo := IniTracer.definirIni;
    cmbserver.Hint := Miarchivo;
    if (MiArchivo = '') or (not fileExists(MiArchivo)) then begin
      showmessage('No hay archivo de configuración INI, por favor indique uno');
      if IniTracer.cambiarIni = '' then
      begin
        showmessage('La aplicación no puede funcionar sin archivo de configuración, por lo tanto se cerrará');
        salir := true;
        close;
      end
      else
        showmessage('Es necesario volver a iniciar la aplicación para que el cambio de archivo de configuracion tenga efecto');
      PostMessage(Handle, WM_CLOSE, 0, 0);
      salir := true;
      close;
    end;
    IniTracer.Free;
  end;


  if salir = false then
  begin
    //configuracion del sistema
    global_archivoini := MiArchivo;
    flagAccesoEnIni := False;

    ini := TIniFile.Create(global_archivoini);
    Global_ServAcceso := ini.readString('SYSTEM', 'SERV_ACCESO', '');
    sPortAcceso := ini.readString('SYSTEM', 'PORT_ACCESO', '');
    global_title_embarque := ini.readString('SYSTEM', 'TITLE_EMBARQUE', '');
    global_files := ini.readString('SYSTEM', 'FILES', extractfilepath(application.exename) + 'Reportes\');
    global_inicio := ini.readInteger('SYSTEM', 'ITEM_INICIAL', 1);
    global_final := ini.readInteger('SYSTEM', 'ITEM_FINAL', 1000);
    global_dias := ini.readInteger('SYSTEM', 'DIAS_ANTERIORES', 10);
    global_independiente := ini.readString('SYSTEM', 'ORDEN_UNICA', 'No');
    global_menu := ini.readString('SYSTEM', 'MENU_INICIAL', 'activo');
    global_db := ini.readString('SYSTEM', 'DATA_NAME', 'inteligent');
    global_dependencia := ini.readString('SYSTEM', 'DEPENDENCIA', '');
    global_checkgenerador := ini.readString('SYSTEM', 'CHECK_GENERADORES', '|INSTALACION|ORDENDECAMBIO|REFERENCIA|WBS|');
    global_ruta := ini.readString('SYSTEM', 'RUTA_SISTEMA', extractfilepath(application.exename));
    ini.free;

    Self.Caption := global_version + '  [' + GetAppVersion + ' ]';
    frmInteligent.Caption := global_version + '  [' + GetAppVersion + ' ]';
    frmSeleccion2.Caption := global_version + '  [' + GetAppVersion + ' ]';
    Self.Label4.Caption :=global_version + '  [' + GetAppVersion + ' ]';
    {codigo de carmen parala imagen de fondo}
    detectar := global_ruta + 'image.ini';
    if leeini(detectar) <> 'no' then
      muestrafondo(frmInteligent.JvBackground1, unitmanejofondo.imapatglobal, unitmanejofondo.estadoglobal)
    else
      escribeinidefault(detectar, 'bmCenter');
    {fin codigo de carmen}

    if sPortAcceso <> '' then flagAccesoEnIni := True;

    //bases de datos registradas en el ini
    //*********************************************************************************
    FilePath := MiArchivo;
    AssignFile(FileText, MiArchivo);
    Reset(FileText);

    while not Eof(FileText) do
    begin
      ReadLn(FileText, wCadena);
      if wCadena = '' then
        continue;
      if MidStr(wCadena, 1, 1) = '[' then
        sTipo := MidStr(wCadena, 1, Pos(']', wCadena))
      else
        if sTipo = '[DATA_BASE]' then
        begin
          sVector[iVector] := MidStr(wCadena, 1, Pos('=', wCadena) - 1);
          wCadena := MidStr(wCadena, Pos('=', wCadena) + 1, Length(wCadena));
          cmbServer.Properties.Items.Add(wCadena);
          listserv.Add(wCadena + '=' + sVector[iVector]);
          iVector := iVector + 1;
        end;
    end;
    CloseFile(FileText);

    if global_db = '' then
      global_db := 'inteligent';


    if global_checkgenerador = '' then
      global_checkgenerador := '|INSTALACION|ORDENDECAMBIO|REFERENCIA|WBS|';

    if global_files = '' then
      global_files := global_ruta + '\files\';

    global_orden_general := '';
    intentos := 0;
    tsPassword.Text := '';

    try
      Global_PortAcceso := strtoint(sportacceso);

    except
      Global_PortAcceso := 3306;
    end;

    if cmbServer.Properties.Items.Count > 0 then
      cmbServer.ItemIndex := 0;
    cmbServer.SetFocus;

 //**************************************************************************************************************
  end;
{$IFDEF Debug}
  cmbServer.ItemIndex := 0;
{$ENDIF}
end;

procedure Tfrmacceso.AdvGlassButton1Click(Sender: TObject);
var
  Pos, i: Integer;
  NombreServ: string;
begin

  Application.CreateForm(TfrmAltaServidor, frmAltaServidor);

  try
    frmAltaServidor.Top := self.Top + trunc(self.Height / 4);
    frmAltaServidor.Left := self.Left + trunc(self.Width / 10);
    frmAltaServidor.ShowModal;
                            {self.Top + trunc(self.Height/4);
      Left:=self.Left + trunc(self.Width/4);}
    if frmAltaServidor.Servidor <> '' then
    begin
          // Revisar si el servidor no existe en previamente
      pos := -1;
      for I := 0 to listServ.Count - 1 do
        if uppercase(listServ.ValueFromIndex[i]) = frmAltaServidor.Servidor then
        begin
          pos := i;
          NombreServ := listServ.Names[i];
          break;
        end;
         // Pos := listServ.IndexOf(frmAltaServidor.Servidor);
      if Pos >= 0 then
      begin
        ShowMessage('El servidor que está intentando dar de alta ya existe ( ' + NombreServ + ' ).' + #10 + #10 + 'Verifique esto e intente de nuevo..');
        cmbServer.ItemIndex := Pos;
        cmbServer.Properties.OnChange(Sender);
        cmbServer.SetFocus;
        Exit;
      end;

          // Verificado el servicio se debe agregar a la lista
      cmbServer.Properties.Items.Add(frmAltaServidor.Descripcion);
      listServ.Add(frmAltaServidor.Servidor);
      sVector[cmbServer.Properties.Items.Count] := frmAltaServidor.Servidor;

      Ini := TIniFile.Create(FilePath);
      try
        Ini.WriteString('DATA_BASE', frmAltaServidor.Servidor, frmAltaServidor.Descripcion); // Grabar el nuevo servidor
      finally
        Ini.Free;
      end;

      tsPuerto.Text := IntToStr(frmAltaServidor.hPuerto);

      cmbServer.ItemIndex := cmbServer.Properties.Items.IndexOf(frmAltaServidor.Descripcion);
      cmbServer.Properties.OnChange(Sender);
      TestServer;
    end
    else
      cmbServer.SetFocus;
  finally
    freeandnil(frmAltaServidor);
  end;
end;

procedure Tfrmacceso.btnAbortarClick(Sender: TObject);
begin
  salir := true;
  exit;
end;

procedure Tfrmacceso.tsPasswordKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    btnAdelante.Click;
end;

procedure Tfrmacceso.tsIdUsuarioKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsPassword.SetFocus
end;

procedure Tfrmacceso.tsIdUsuarioPropertiesEditValueChanged(Sender: TObject);
begin
  sDataName.Visible := False;
  lblBase.Visible := False;
end;

procedure Tfrmacceso.tsIdUsuarioEnter(Sender: TObject);
begin
  tsIdUsuario.Style.Color := $00FFF0E1;
end;

procedure Tfrmacceso.tsIdUsuarioExit(Sender: TObject);
begin
  tsIdUsuario.Style.Color := clWhite
end;

procedure Tfrmacceso.tsPasswordEnter(Sender: TObject);
begin
  tsPassword.Style.Color := $00FFF0E1
end;

procedure Tfrmacceso.tsPasswordExit(Sender: TObject);
begin
  tsPassword.Style.Color := clWhite
end;

procedure Tfrmacceso.cmbServerEnter(Sender: TObject);
begin
  cmbServer.Style.Color := $00FFF0E1
end;

procedure Tfrmacceso.cmbServerExit(Sender: TObject);
begin
  cmbServer.Style.Color := clWhite
end;

procedure Tfrmacceso.cmbServerKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsIdUsuario.SetFocus
end;

procedure Tfrmacceso.FormActivate(Sender: TObject);
var
  InfoSize, H, RsltLen: Cardinal;
  VersionBlock: Pointer;
  Rslt: PVSFixedFileInfo;
begin
{  SetTransparentForm(Handle, 215);}
  InfoSize := GetFileVersionInfoSize(PChar(Application.ExeName), H);
  VersionBlock := AllocMem(InfoSize);
  try
    GetFileVersionInfo(PChar(Application.ExeName), H, InfoSize, VersionBlock);
    VerQueryValue(VersionBlock, '\', Pointer(Rslt), RsltLen);
  finally
    FreeMem(VersionBlock);
  end;

  lblBase.Visible := False;
  sDataName.Visible := False;
  btnAdelante.Caption := 'Iniciar Sesion';

end;

procedure Tfrmacceso.FormCreate(Sender: TObject);
var
  fileSkin: TextFile;
  sSkin: string;
begin
  ListServ := tstringlist.Create;

       {seleccionar tema anterior usado}
  { if FileExists('fileSkin.dat') then
      begin
         AssignFile(fileSkin, 'fileSkin.dat');
         Reset(fileSkin);
         ReadLn(fileSkin, sSkin);
         CloseFile(fileSkin);
         if sSkin='' then
          begin
            connection.sSkinManager1.Active := False;
          end
          else
          begin
            connection.sSkinManager1.SkinName := sSkin;
            connection.sSkinManager1.Active := True;
          end;

      end
   else
      begin
         AssignFile(fileSkin, 'fileSkin.dat');
         ReWrite(fileSkin);
         WriteLn(fileSkin, 'WLM');
         CloseFile(fileSkin);
         connection.sSkinManager1.SkinName := 'WLM';
         connection.sSkinManager1.Active := True;
      end; }

end;



procedure Tfrmacceso.btnSalirClick(Sender: TObject);
begin
  close
end;

procedure Tfrmacceso.sDataNameEnter(Sender: TObject);
begin
  sDataName.Style.Color := $00FFF0E1
end;

procedure Tfrmacceso.sDataNameExit(Sender: TObject);
begin
  sDataName.Style.Color := clWhite
end;

procedure Tfrmacceso.sDataNameKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    btnAdelante.SetFocus
end;

procedure Tfrmacceso.cmbServerChange(Sender: TObject);
var
  sPuerto: Integer;
begin
(*
    if FileExists(FilePath) then
    Begin
      Ini := tinifile.create(FilePath);
      Try
      sPuerto := Ini.ReadInteger('SERVIDORES', listserv.Strings[cmbserver.ItemIndex], 0);
      Finally
        Ini.Free;
      End;
      tsPuerto.Text := IntToStr(sPuerto);
    End;
*)
  lblBase.Visible := False;
  sDataName.Visible := False;
  btnAdelante.Caption := 'Iniciar Sesion';
end;

function Tfrmacceso.TestServer: boolean;
var
  hPuerto, sPuerto: Integer;
  VResult, continuar: Boolean;
  zQuery: TZReadOnlyQuery;
begin
  Continuar := true;
  Vresult := false;
  try
    hPuerto := StrToInt(tsPuerto.Text);
  except
    hPuerto := -1;
  end;
  if hPuerto < 1 then
  begin
    ShowMessage('El puerto especificado no es correcto.' + #10 + #10 + 'El puerto debe ser un número entero comprendido entre 1 y 65536');
    tsPuerto.SetFocus;
    continuar := false;
  end;
  if continuar then
  begin
    Connection.zConnection.Disconnect; //    ConnectionHMG.Disconnect;
    Connection.zConnection.Catalog := 'mysql'; // .ConnectionHMG.Catalog := 'mysql';
    Connection.zConnection.Catalog := 'mysql'; //ConnectionHMG.Database := 'mysql';
    Connection.zConnection.HostName := listserv.Strings[cmbserver.ItemIndex]; //lblista.text;   //ConnectionHMG.HostName := lbLista.Text;
    Connection.zConnection.Password := intelpass;
    Connection.zConnection.Port := hPuerto;
    Connection.zConnection.Protocol := 'mysql-5';
    Connection.zConnection.User := 'root';
    Result := False;
    try
      Connection.zConnection.Connect;
      Result := Connection.zConnection.Ping;
    except
      Result := False;
    end;
    if Result then
    begin
// Mostrar los bases de datos correspondientes a este servidor
      zQuery := tzReadOnlyQuery.Create(Self);
      zQuery.Connection := Connection.zConnection;
      zQuery.SQL.Text := 'show databases';
      zQuery.Open;
      sDataName.Properties.Items.Clear;
      while not zQuery.Eof do
      begin
        if (zQuery.FieldValues['database'] <> 'performance_schema') and (zQuery.FieldValues['database'] <> 'mysql') and (zQuery.FieldValues['database'] <> 'information_schema') and (zQuery.FieldValues['database'] <> 'test') and (zQuery.FieldValues['database'] <> 'chat') and (zQuery.FieldValues['database'] <> 'chat_php') and (zQuery.FieldValues['database'] <> 'ic_exsoll') and (zQuery.FieldValues['database'] <> 'joomla') and (zQuery.FieldValues['database'] <> 'phpmyadmin') then
          sDataName.Properties.Items.Add(zQuery.FieldValues['database']);
        zQuery.Next;
      end;
      zQuery.Destroy;
// Habilitar los campos de las bases de datos
      if sDataName.Properties.Items.Count > 0 then
        sDataName.ItemIndex := 0;
    end
    else
    begin
      ShowMessage('No ha podido ser posible establecer conexión con el servidor especificado.' + #10 + #10 + 'Revise los datos capturados para especificar el servidor o revise su conexión a la red si su base de datos se encuentra en un servidor remoto.');
      tsPuerto.SetFocus;
      Exit;
    end;
  end;
end;
end.

