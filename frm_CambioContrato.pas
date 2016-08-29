unit frm_CambioContrato;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, global, ComCtrls, 
  StdCtrls, ExtCtrls, DBCtrls, db, Menus, OleCtrls,
  Buttons, 
  ZAbstractRODataset, ZDataset, 
  rxToolEdit, UnitExcepciones;

type
  TfrmCambioContrato = class(TForm)
    Progress: TProgressBar;
    btnOk: TBitBtn;
    btnExit: TBitBtn;
    Memo1: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    txtContratoActual: TEdit;
    txtContratoNuevo: TEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnExitClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure FormShow(Sender: TObject);

  private
    { Private declarations }
  public

    { Public declarations }
  end;

var
  frmCambioContrato: TfrmCambioContrato;
  sCadena: string;
  qryAnexo: TZReadOnlyQuery;
  iTabla: Byte;

implementation

{$R *.dfm}

procedure TfrmCambioContrato.btnExitClick(Sender: TObject);
begin
  close;
end;



procedure TfrmCambioContrato.btnOkClick(Sender: TObject);
var
  iAnexo: Byte;
  i, x: integer;
  base, tabla, campo, cad: string;
  datos: array[1..200] of string;
begin
  try
    qryAnexo := tzReadOnlyQuery.Create(Self);
    qryAnexo.Connection := Connection.zConnection;

    if txtContratoNuevo.Text <> '' then
    begin
//      if not UnitTablasImpactadas.boolRelaciones(connection.zConnection) then
//      begin
        connection.qryBusca.Active := False;
        connection.qryBusca.SQL.Clear;
        connection.qryBusca.SQL.Add('SET FOREIGN_KEY_CHECKS=0');
        connection.qryBusca.ExecSQL;


        connection.qryBusca.Active := False;
        connection.qryBusca.SQL.Clear;
        connection.qryBusca.SQL.Add('select sContrato from contratos where sContrato =:Contrato');
        connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
        connection.qryBusca.Params.ParamByName('Contrato').Value := txtContratoActual.Text;
        connection.qryBusca.Open;

        if connection.QryBusca.RecordCount > 0 then
        begin
          connection.qryBusca.Active := False;
          connection.qryBusca.SQL.Clear;
          connection.qryBusca.SQL.Add('Show tables');
          connection.qryBusca.Open;
          base := 'Tables_in_' + global_db;
          i := 1;
          progress.Min := 1;
          progress.Position := 1;
          progress.Max := connection.QryBusca.RecordCount + 1;
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
                if connection.QryBusca2.FieldValues['Field'] = 'sContrato' then
                begin
                  datos[i] := tabla;
                  i := i + 1;
                  //memo1.Lines.Add(tabla);
                end;
                connection.QryBusca2.Next;
              end;
            end;
            connection.QryBusca.Next;
            progress.Position := progress.Position + 1;
          end;

            //resetear la barra de progreso
          progress.Position := 1;
          progress.Max := i;
            // Actualiza todos los registros..
          for x := 1 to i - 1 do
          begin
            tabla := datos[x];
            memo1.Lines.Add('Actualizando en : ' + tabla);
            connection.qryBusca.Active := False;
            connection.qryBusca.SQL.Clear;
            connection.qryBusca.SQL.Add('update ' + tabla + ' set sContrato = :ContratoNuevo where sContrato = :Contrato ');
            connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
            connection.qryBusca.Params.ParamByName('Contrato').Value := txtContratoActual.Text;
            connection.qryBusca.Params.ParamByName('ContratoNuevo').DataType := ftString;
            connection.qryBusca.Params.ParamByName('ContratoNuevo').Value := txtContratoNuevo.Text;
            connection.qryBusca.ExecSQL;
            progress.Position := progress.Position + 1;
          end;
          connection.qryBusca.Active := False;
          connection.qryBusca.SQL.Clear;
          connection.qryBusca.SQL.Add('SET FOREIGN_KEY_CHECKS=1');
          connection.qryBusca.ExecSQL;
          MessageDlg('Proceso Terminado Con Exito, para cargar los cambios salga y entre del programa.', mtInformation, [mbOk], 0);
        end
        else
          MessageDlg('No existe el Contrato Actual !', mtInformation, [mbOk], 0);
//      end
//      else
//      begin
//        connection.qryBusca.Active := False;
//        connection.qryBusca.SQL.Clear;
//        connection.qryBusca.SQL.Add('update contratos set sContrato = :ContratoNuevo where sContrato = :Contrato ');
//        connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
//        connection.qryBusca.Params.ParamByName('Contrato').Value := txtContratoActual.Text;
//        connection.qryBusca.Params.ParamByName('ContratoNuevo').DataType := ftString;
//        connection.qryBusca.Params.ParamByName('ContratoNuevo').Value := txtContratoNuevo.Text;
//        connection.qryBusca.ExecSQL;
//      end;
    end
    else
      MessageDlg('Debe escribir un Contrato Nuevo..!', mtInformation, [mbOk], 0);
  except
    on e: exception do begin
      connection.qryBusca.Active := False;
      connection.qryBusca.SQL.Clear;
      connection.qryBusca.SQL.Add('SET FOREIGN_KEY_CHECKS=1');
      connection.qryBusca.ExecSQL;
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Cambio de Nombre de Contrato', 'Al cambiar nombre de contrato', 0);
    end;
  end;
end;



procedure TfrmCambioContrato.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree;
end;

procedure TfrmCambioContrato.FormShow(Sender: TObject);
begin
  txtContratoActual.Text := Global_Contrato;
  txtContratoNuevo.SetFocus;
end;

end.

