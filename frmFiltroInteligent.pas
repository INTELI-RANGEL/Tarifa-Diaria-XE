unit frmFiltroInteligent;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, RXDBCtrl, StdCtrls, frm_Connection, global, DB, StrUtils,
  Menus, ZAbstractRODataset, ZDataset, UdbGrid, UnitExcepciones;

type
  TfrmFiltros = class(TForm)
    GridResult: TRxDBGrid;
    grFiltro: TGroupBox;
    tsCampo: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    tsCondicion: TComboBox;
    tsValor: TEdit;
    Button1: TButton;
    Label3: TLabel;
    dsQryResult: TDataSource;
    btnEliminar: TButton;
    tmFiltro: TMemo;
    popExportar: TPopupMenu;
    Exportar1: TMenuItem;
    SaveSql: TSaveDialog;
    QryResult: TZReadOnlyQuery;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GridResultTitleBtnClick(Sender: TObject; ACol: Integer;
      Field: TField);
    procedure GridResultColEnter(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnEliminarClick(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure GridResultMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure GridResultMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GridResultTitleClick(Column: TColumn);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  UtGrid:TicDbGrid;
  frmFiltros: TfrmFiltros;
  sQuery     : String ;
  isNumeric  : Boolean ;
  sCondicion : String ;

implementation

{$R *.dfm}

procedure TfrmFiltros.FormShow(Sender: TObject);
begin
    UtGrid:=TicdbGrid.create(gridResult);
    sQuery      := 'Select e.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador, e2.sNumeroActividad, e2.sIsometrico, e2.sPrefijo, e2.sIsometricoReferencia, ' +
                   'e2.sInstalacion, e2.iOrdenCambio, e2.dCantidad, e2.mComentarios From estimacionxpartida e2 ' +
                   'INNER JOIN estimaciones e ON (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e.sNumeroGenerador) Where ' ;
    sCondicion  := 'e2.sContrato = :Contrato ' ;
    tmFiltro.Lines.Clear ;
    tmFiltro.Lines.Add(sCondicion) ;

    tsCampo.Text := 'iNumeroEstimacion' ;
    isNumeric := False ;
    QryResult.Active := False ;
    QryResult.SQL.Clear ;
    QryResult.SQL.Add(sQuery + tmFiltro.Text ) ;
    QryResult.Params.ParamByName('Contrato').DataType := ftString ;
    QryResult.Params.ParamByName('Contrato').Value := global_contrato ;
    QryResult.Open ;
end;

procedure TfrmFiltros.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  utgrid.Destroy;
  action := cafree ;
end;

procedure TfrmFiltros.GridResultTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
Var
   sOrder : String ;
   iLine     : Byte ;
   sConsulta : tStrings ;
   sCampo    : String ;
begin
   sConsulta := TStringList.Create;
   sOrder := 'Order By ' + Field.FieldName ;
   For iLine := 0 To tmFiltro.Lines.Count Do
        If (iLine >= 1) And (tmFiltro.Lines.Strings[iLine] <> '') Then
            sConsulta.Add ( ' And ' + tmFiltro.Lines.Strings[iLine] )
        Else
            sConsulta.Add ( tmFiltro.Lines.Strings[iLine] ) ;
   QryResult.Active := False ;
   QryResult.SQL.Clear ;
   QryResult.SQL.Add(sQuery + sConsulta.Text + sOrder) ;
   QryResult.Params.ParamByName('Contrato').DataType := ftString ;
   QryResult.Params.ParamByName('Contrato').Value := global_contrato ;
   QryResult.Open ;
end;

procedure TfrmFiltros.GridResultTitleClick(Column: TColumn);
begin
   UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmFiltros.GridResultColEnter(Sender: TObject);
begin
  try
    with GridResult.SelectedField do
    Begin
        If DataType = ftFloat Then
            isNumeric := True
        Else
            isNumeric := False ;
        tsCampo.Text := DisplayLabel ;
        tsValor.Text := Value
    End
  except

  end;
end;

procedure TfrmFiltros.GridResultMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmFiltros.GridResultMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmFiltros.Button1Click(Sender: TObject);
Var
    UtGrid:TicDbGrid;
    sCadena   : String ;
    iLine     : Byte ;
    sConsulta : tStrings ;
    sCampo    : String ;
begin
  try
    If isNumeric Then
        sCadena := tsValor.Text
    Else
        sCadena := '"' + tsValor.Text + '"' ;

    sConsulta := TStringList.Create;
    If tsCampo.Text = 'iNumeroEstimacion' Then
        sCampo := 'e.' + tsCampo.Text
    Else
        sCampo := 'e2.' + tsCampo.Text ;

    If tsCampo.Text <> '' Then
    Begin
        If tsCondicion.Text = 'IGUAL A' Then
            tmFiltro.Lines.Add(sCampo + ' = ' + sCadena )
        Else
            If tsCondicion.Text = 'MAYOR QUE' Then
                tmFiltro.Lines.Add(sCampo + ' > ' + sCadena )
            Else
                If tsCondicion.Text = 'MENOR QUE' Then
                    tmFiltro.Lines.Add(sCampo + ' < ' + sCadena )
                Else
                    If tsCondicion.Text = 'MAYOR IGUAL QUE' Then
                        tmFiltro.Lines.Add(sCampo + ' >= ' + sCadena )
                    Else
                        If tsCondicion.Text = 'MENOR IGUAL QUE' Then
                            tmFiltro.Lines.Add(sCampo + ' <= ' + sCadena )
                        Else
                            If tsCondicion.Text = 'DIFERENTE DE' Then
                                  tmFiltro.Lines.Add(sCampo + ' <> ' + sCadena )
                            Else
                                  If tsCondicion.Text = 'DENTRO DE' Then
                                      If NOT isNumeric Then
                                          tmFiltro.Lines.Add( sCampo + ' LIKE "%' + tsValor.Text + '%"' ) ;
        For iLine := 0 To tmFiltro.Lines.Count Do
            If (iLine >= 1) And (tmFiltro.Lines.Strings[iLine] <> '') Then
                sConsulta.Add ( ' And ' + tmFiltro.Lines.Strings[iLine] )
            Else
                sConsulta.Add ( tmFiltro.Lines.Strings[iLine] ) ;
        QryResult.Active := False ;
        QryResult.SQL.Clear ;
        QryResult.SQL.Add(sQuery + sConsulta.Text) ;
        QryResult.Params.ParamByName('Contrato').DataType := ftString ;
        QryResult.Params.ParamByName('Contrato').Value := global_contrato ;
        QryResult.Open ;
    End ;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Filtro de Estimaciones', 'Al agregar filtro', 0);
    end;
  end;
end;

procedure TfrmFiltros.btnEliminarClick(Sender: TObject);
begin
  try
    tmFiltro.Lines.Clear ;
    tmFiltro.Lines.Add(sCondicion) ;
    QryResult.Active := False ;
    QryResult.SQL.Clear ;
    QryResult.SQL.Add(sQuery + tmFiltro.Text ) ;
    QryResult.Params.ParamByName('Contrato').DataType := ftString ;
    QryResult.Params.ParamByName('Contrato').Value := global_contrato ;
    QryResult.Open ;
  except
    on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Filtro de Estimaciones', 'Al eliminar filtro', 0);
    end;
  end;
end;

procedure TfrmFiltros.Exportar1Click(Sender: TObject);
Var
   F1            : TextFile;
   Registro, Code, ValorTemporal  : Integer ;
   Cadena        : String ;
   ValorNumerico : String ;
   ValorFecha    : String ;
   valorCampo     : Variant ;
begin
  try
    SaveSql.Title := 'Guardar Consulta';
    If QryResult.Active = True Then
        If SaveSql.Execute Then
        Begin
            AssignFile(F1, SaveSql.FileName);
            rewrite(F1) ;
            QryResult.First ;
            While QryResult.Eof <> True Do
            Begin
                Cadena := '' ;
                ValorCampo := '' ;
                for Registro := 0 to QryResult.FieldCount - 1 do
                Begin
                      ValorCampo :=  QryResult.Fields[Registro].Value ;
                      If VarIsnull(ValorCampo)  Then
                      Begin
                          ValorCampo :='' ;
                          Cadena := Cadena + '"' + ValorCampo + '"' + ',' ;
                      end
                      Else
                      Begin
                          If (QryResult.Fields[Registro].DataType = ftString)	OR (QryResult.Fields[Registro].DataType = ftMemo) Then
                          Begin
                              ValorCampo :=  AnsiReplaceText ( ValorCampo , '"' , '""' ) ;
                              If Length(ValorCampo) > 0 Then
                                   ValorCampo := '"' + ValorCampo + '"' ;
                              Cadena := Cadena + ValorCampo + ',' ;
                          End
                          Else
                              If QryResult.Fields[Registro].DataType = ftDate	Then
                              Begin
                                  ValorFecha := '"' + DateToStr(ValorCampo) + '"'  ;
                                  Cadena := Trim(Cadena + ValorFecha + ',') ;
                              End
                              Else
                              If QryResult.Fields[Registro].DataType <> ftBlob	Then
                              Begin
                                   ValorNumerico := '"' + Trim(ValorCampo) + '"' ;
                                  Cadena := Trim(Cadena + ValorNumerico +',') ;
                              End ;
                      End
                end ;
                WriteLn(F1, Cadena ) ;
                QryResult.Next ;
            End ;
            closefile(F1) ;
        End
  except
    on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Filtro de Estimaciones', 'Al exportar a excel', 0);
    end;
  end;
end;

end.
