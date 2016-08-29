unit Frm_ImportaProject;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, NxScrollControl, NxCustomGridControl, NxCustomGrid, NxGrid,
  NxColumns, NxColumnClasses, StdCtrls,ComCtrls,StrUtils,RxMemDS,ComObj,
  frm_connection, global,db,Utilerias, Grids, DBGrids, JvExDBGrids, JvDBGrid,
  JvDBUltimGrid, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint,
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
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, cxDropDownEdit, cxGridCustomTableView, cxGridTableView,
  cxGridBandedTableView, cxGridCustomView, cxClasses, cxGridLevel, cxGrid,
  NxPropertyItemClasses,NxInspector, cxSplitter, cxTL, cxTLdxBarBuiltInMenu,
  cxInplaceContainer, Mask, JvExMask, JvToolEdit, cxContainer, cxLabel,
  cxProgressBar, JvExControls, JvDBLookup, ZAbstractRODataset, ZDataset,cxMemo;

type
  TFrmImportaProject = class(TForm)
    Panel2: TPanel;
    ArchivoMsP: TFileOpenDialog;
    Panel3: TPanel;
    NxInsColumnas: TNextInspector;
    cxSplitter1: TcxSplitter;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    CxTlsPrograma: TcxTreeList;
    Panel1: TPanel;
    Panel4: TPanel;
    JFedtArchivo: TJvFilenameEdit;
    cxLabel1: TcxLabel;
    btnImportar: TButton;
    CxPbAvance: TcxProgressBar;
    QrContratos: TZReadOnlyQuery;
    dsContratos: TDataSource;
    dsFolios: TDataSource;
    QrFolios: TZReadOnlyQuery;
    Label2: TLabel;
    Label3: TLabel;
    jDblCmbFolio: TJvDBLookupCombo;
    jDblCmbContrato: TJvDBLookupCombo;
    Panel5: TPanel;
    StyleRepository: TcxStyleRepository;
    cxStyle2: TcxStyle;
    cxStyle3: TcxStyle;
    cxStyle4: TcxStyle;
    cxStyle5: TcxStyle;
    cxStyle6: TcxStyle;
    cxStyle7: TcxStyle;
    cxStyle8: TcxStyle;
    cxStyle9: TcxStyle;
    cxStyle10: TcxStyle;
    cxStyle11: TcxStyle;
    cxStyle12: TcxStyle;
    cxStyle13: TcxStyle;
    cxStyle14: TcxStyle;
    stlGroupNode: TcxStyle;
    stlFixedBand: TcxStyle;
    TreeListStyleSheetDevExpress: TcxTreeListStyleSheet;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure JFedtArchivoAfterDialog(Sender: TObject; var AName: string;
      var AAction: Boolean);
    procedure FormShow(Sender: TObject);
    procedure jDblCmbContratoChange(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
    procedure CxTlsProgramaStylesGetContentStyle(Sender: TcxCustomTreeList;
      AColumn: TcxTreeListColumn; ANode: TcxTreeListNode; var AStyle: TcxStyle);
  private
    const
      Db_Campos:Array[1..13,1..4] of string=(('No. Actividad','1','Texto1','sNumeroActividad'),('Tipo Actividad','0','','sTipoActividad'),('Tipo Anexo','0','','sTipoAnexo'),('Fecha de Inicio','1','Comienzo','dFechaInicio'),
                                             ('Fecha de Termino','1','Fin','dFechaFinal'),('Duración','0','','dDuracion'),('Ponderado','0','','dPonderado'),('Unid. Medida','0','','sMedida'),('Cantidad','0','','dCantidad'),
                                             ('Costo Mn','0','','dVentaMN'),('Costo Dll','0','','dVentaDLL'),('Id Fase','0','','sIdFase'),('Anexo','0','','sAnexo'));



    { Private declarations }
    function FindMyNode(TreeList: TcxTreeList; const ID: Variant; ColIndex: Integer): TcxTreeListNode;
    Procedure PreImportOTProject(sFileProject:TFileName;Grid:TcxTreeList;Custom:TNextInspector;PgbAvance: TcxProgressBar=nil);
    Function LoadHeaderProject(var MsProject: Variant;var ActProject:Variant;Grid:TcxGridTableView;Custom:TNextInspector;PgbAvance: TcxProgressBar=nil):Boolean;Overload;
    Function LoadHeaderProject(var MsProject: Variant;var ActProject:Variant;Grid:TcxTreeList;Custom:TNextInspector;PgbAvance: TcxProgressBar=nil):Boolean;Overload;

    Function PreviewProject(var MsProject: Variant;var ActProject:Variant;Grid:TcxGridTableView):Boolean; Overload;
    Function PreviewProject(var MsProject: Variant;var ActProject:Variant;Grid:TcxTreeList;PgbAvance: TcxProgressBar=nil):Boolean; Overload;

    Procedure ImportOTProject(ParamContrato,ParamConvenio,ParamFolio:String;sFileProject:TFileName;Grid:TcxTreeList;Custom:TNextInspector;PgbAvance: TcxProgressBar=nil);


    Function ValidaOtProject(ParamContrato,ParamConvenio,ParamFolio:String;Grid:TcxTreeList;var MsProject: Variant;ListaParams:TStringList;PgbAvance: TcxProgressBar=nil):Boolean;

  public
    { Public declarations }
  end;




var
  FrmImportaProject: TFrmImportaProject;

implementation

uses Frm_PopUpImportacionPP,UFunctionsGHH,UnitExcepciones;

{$R *.dfm}


Function TFrmImportaProject.ValidaOtProject(ParamContrato,ParamConvenio,ParamFolio:String;Grid:TcxTreeList;var MsProject: Variant;ListaParams:TStringList;PgbAvance: TcxProgressBar=nil):Boolean;
var
  ActProject,
  Task:Variant;
  ListaPdas:TStringList;
  i,Reng:Integer;
  sPda,sAnexo,sCampo,sTmpCampo,sAnd:String;
  posVar:Integer;
  QrConsulta:tzReadOnlyQuery;
  projectField: Integer;
  Valor:Variant;
  sTipoVal:String;
 // Inicio:TDateTime;
  isBad:Boolean;
  AMemo: TcxMemo;
  Nodo:TcxTreeListNode;
  sWbs,sCadError:String;
begin
  IsBad:=false;
  ActProject:=MsProject.ActiveProject;
  QrConsulta:=TzReadOnlyQuery.Create(nil);
  try

    if Grid.ColumnCount < ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+2 then
    begin
      with Grid.CreateColumn do
      begin
        Caption.text:='Observaciones de Validación';
        Width:=250;
        Caption.AlignHorz:=TaCenter;
        Caption.AlignVert:=TcxAlignmentVert(TaCenter);
        Caption.MultiLine:=true;
        Editing:=false;
        PropertiesClass:= TcxMemoProperties;
        Name:='ObsSAMG';
        Styles.Header := cxStyle1;
      end;
    end;


    QrConsulta.Connection:=Connection.zConnection;
    for Reng:=0 to ActProject.Tasks.Count-1 do
    begin
      Task:=ActProject.Tasks.item[Reng+1];
      sPda:='';
      sAnexo:='';
      sWbs:=Task.wbs;
      Nodo:=nil;
      sCadError:='';
      if sWbs<>'' then
        Nodo := FindMyNode(Grid,sWbs,0);

      if Nodo<>nil then
        nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1]:='';

      for i := 1 to ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count do
      begin
        sCampo:=MsProject.FieldConstantToFieldName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);
        sTmpCampo:=sCampo;
        posVar:=ListaParams.IndexOfName(sCampo);

        if PosVar<>-1 then
        begin
          sCampo:=ListaParams.ValueFromIndex[PosVar];
          if AnsiCompareText(sCampo,'sNumeroActividad')=0 then
            sPda:=Task.GetField(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);

          if AnsiCompareText(sCampo,'sAnexo')=0 then
            sAnexo:=Task.GetField(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);

          valor:=GetPrjField(Task,ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);
          sTipoVal:=ShowBasicVariantType(valor);

          if (sTipoVal='') or (sTipoVal='IsEmpty') or
          (sTipoVal='IsError')  then
          begin  //Aqui Hay error
            try
              if sTipoVal='IsError' then
                sCadError:='Campo: '+sTmpCampo+', Error: Tipo de Datos Erroneo.'
              ELSE
                sCadError:='Campo: '+sTmpCampo+', Error: Sin Valor para Almacenar.';

              if Nodo<>nil then
              begin
                if nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1]='' then
                  nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1]:=sCadError
                else
                  nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1]:=
                  nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1] + #13 + #10 +
                  sCadError;

              end;
            except
              On E:Exception do
              begin
                showmessage(e.message + ', ' + e.ClassName);

              end;
            end;
            isBad:=true;
          end;

        end;

        if i=ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count then
        begin
          sAnd:='';
          if sAnexo<>'' then
            sAnd:= 'and sAnexo=' + quotedstr(sAnexo);

          QrConsulta.Active:=false;
          QrConsulta.SQL.Text:= 'Select * from actividadesxorden where sContrato=:Contrato and sIdConvenio=:Convenio '+
                                'and sNumeroOrden=:Orden and sNumeroActividad=:Actividad ' + sAnd;
          QrConsulta.ParamByName('Contrato').AsString:=ParamContrato;
          QrConsulta.ParamByName('Convenio').AsString:=ParamConvenio;
          QrConsulta.ParamByName('Orden').AsString:=ParamFolio;
          QrConsulta.ParamByName('Actividad').AsString:=sPda;
          QrConsulta.Open;
          if QrConsulta.RecordCount>0 then
          begin
            sCadError:='Partida Duplicada en el Folio';

            if Nodo<>nil then
            begin
              if nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1]='' then
                nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1]:=sCadError
              else
                nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1]:=
                nodo.Values[ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count+1] + #13 + #10 +
                sCadError;

            end;


          end;

        end;
      end;
      if PgbAvance<>nil then
      begin
        PgbAvance.Position:=PgbAvance.Position + 1;
        Application.ProcessMessages;
      end;
    end;
  finally
    QrConsulta.Destroy;
  end;
  Result:= Not IsBad;
end;

Procedure TFrmImportaProject.ImportOTProject(ParamContrato,ParamConvenio,ParamFolio:String;sFileProject:TFileName;Grid:TcxTreeList;Custom:TNextInspector;PgbAvance: TcxProgressBar=nil);
{$Region 'Informacion'}
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
{$EndRegion}
var
  ListaParametros: TStringList;
  MsProject:Variant;
  ActProject:Variant;
  Error:Boolean;
  i:Integer;
  InsCustomColumn:TNxComboBoxItem;
  bContinue:Boolean;
begin
  Error:=false;
  bContinue:=true;
  if FileExists(sFileProject) then
  begin
    if AnsiEndsText('.mpp',sFileProject) then
    begin
      ListaParametros:=TstringList.Create;
      try
        try
          for I := 0 to Custom.Items.Count-1 do
          begin
            InsCustomColumn:=TNxComboBoxItem(Custom.Items.Item[I]);
            if (InsCustomColumn.Tag=1) or ((InsCustomColumn.Tag=0) and (InsCustomColumn.ItemIndex>0)) then
              ListaParametros.Add(InsCustomColumn.Value + '=' + InsCustomColumn.Name);
          end;

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
            try
              MsProject.visible:=false;
              MsProject.DisplayAlerts:=false;
              if not error then
              begin
                MsProject.FileOpen(sFileProject);
                ActProject:=MsProject.ActiveProject;

                if PgbAvance<>nil then
                begin
                  PgbAvance.Properties.min:=0;
                  PgbAvance.Properties.Max := ActProject.Tasks.Count*2;
                  PgbAvance.Position := 0;
                  PgbAvance.Visible:=true;
                  Application.ProcessMessages;
                end;

                bContinue:=ValidaOtProject(ParamContrato,ParamConvenio,ParamFolio,Grid,MsProject,ListaParametros,PgbAvance);
                {if LoadHeaderProject(MsProject,ActProject,Grid,Custom,PgbAvance) then
                  if PreviewProject(MsProject,ActProject,Grid,PgbAvance) then
                    btnImportar.Enabled:=true
                  else
                    btnImportar.Enabled:=false; }
              end;
            finally
              MsProject.visible:=true;
              MsProject.DisplayAlerts:=true;
              if bContinue then
                MsProject.Quit;
              if PgbAvance<>nil then
              begin
                PgbAvance.Visible:=False;
                Application.ProcessMessages;
              end;
            end;
          end;
        except

        end;
      finally
        ListaParametros.Destroy;
      end;
    end
    else
      MessageDlg('El Archivo: '+ #13 + #10 + Quotedstr(sFileProject) + ' NO es Valido.',
                  MtError,[MbOk],0);
  end;
end;



procedure TFrmImportaProject.jDblCmbContratoChange(Sender: TObject);
begin
  if (QrFolios.Active) and (QrFolios.RecordCount>0) then
    jDblCmbFolio.KeyValue:=QrFolios.FieldByName('sNumeroOrden').AsString;
end;

procedure TFrmImportaProject.btnImportarClick(Sender: TObject);
begin
  ImportOTProject(jDblCmbContrato.KeyValue,global_Convenio,jDblCmbFolio.KeyValue,JFedtArchivo.FileName,CxTlsPrograma,NxInsColumnas,CxPbAvance);

end;

procedure TFrmImportaProject.CxTlsProgramaStylesGetContentStyle(
  Sender: TcxCustomTreeList; AColumn: TcxTreeListColumn; ANode: TcxTreeListNode;
  var AStyle: TcxStyle);
begin
  if ANode.IsGroupNode then
    AStyle := stlGroupNode;
end;

function TFrmImportaProject.FindMyNode(TreeList: TcxTreeList; const ID: Variant; ColIndex: Integer): TcxTreeListNode;
  function FindChildNode(Node: TcxTreeListNode): TcxTreeListNode;
  var
   J: integer;
  begin
    result := nil;

    if (Node <> nil) then
      for J := 0 to Node.Count - 1 do
      begin
        if (Node.Items[J] = nil) then
          Continue;

        if (Node.Items[J].Values[ColIndex] = ID) then
        begin
          result := Node.Items[J];
          break;
        end;
        if Node.Items[J].HasChildren then
        begin
          result := FindChildNode(Node.Items[J]);
          if (result <> nil) then
            break;
        end;
      end;
  end;

var
  i: Integer;
begin
  result := nil;

  for I := 0 to TreeList.Count - 1 do
  begin
    if (TreeList.Items[I] = nil) then
      Continue;

    if (TreeList.Items[I].Values[ColIndex] = ID) then
    begin
      result := TreeList.Items[I];
      break;
    end
    else
      if TreeList.Items[I].HasChildren then
      begin
        result := FindChildNode(TreeList.Items[I]);
        if (result <> nil) then
          break;
      end;
  end;
end;

Function TFrmImportaProject.LoadHeaderProject(var MsProject: Variant;var ActProject: Variant; Grid: TcxTreeList; Custom: TNextInspector;PgbAvance: TcxProgressBar=nil):Boolean;
var
  InsCustomColumn:TNxComboBoxItem;
  i,x:Integer;
  Resource:Variant;
  ListaCol:TStringList;
  IndexDefault:Integer;
  Res:Boolean;
begin
  Res:=true;
  ListaCol:=TStringList.Create;
  Grid.BeginUpdate;
  try
    try
      Custom.Items.Clear;
      Grid.DeleteAllColumns;
      Grid.Clear;
      with Grid.OptionsView do
      begin
        cellAutoHeight:=true;
        GridLines:=tlglBoth;
        Indicator:=true;
        TreeLineStyle:=tllsDot;
        HeaderAutoHeight:=true;
        DynamicIndent:=true;
      end;

      with Grid.OptionsData do
      begin
        Deleting:=false;
        Editing:=false;
      end;

      with Grid.CreateColumn do
      begin
        Caption.text:='*';
        Width:=100;
        Caption.AlignHorz:=TaCenter;
        Caption.AlignVert:=TcxAlignmentVert(TaCenter);
        Caption.MultiLine:=true;
        Editing:=false;
        Styles.Header := cxStyle1;
      end;

      for i := 1 to ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count do
      begin
        with Grid.CreateColumn do
        begin
          Caption.text:=MsProject.FieldConstantToFieldName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);
          Width:=(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Width)*10;
          Caption.AlignHorz:=TaCenter;
          Caption.AlignVert:=TcxAlignmentVert(TaCenter);
          Caption.MultiLine:=true;
          Editing:=false;
          ListaCol.Add(Caption.text);
          Styles.Header := cxStyle1;
        end;

        if PgbAvance<>nil then
        begin
          PgbAvance.Position := PgbAvance.Position + 1;
          Application.ProcessMessages;
        end;
        
      end;

      for x := 1 to length(Db_Campos) do
      begin
        IndexDefault:=0;
        InsCustomColumn:=TNxComboBoxItem.Create(nil);
        InsCustomColumn.Caption:=Db_Campos[x,1];
        InsCustomColumn.Style:=cbDropDownList;
        InsCustomColumn.Alignment:=taCenter;
        InsCustomColumn.ItemHeight:=20;
        InsCustomColumn.Font.Style:=[FsBold];
        InsCustomColumn.Name:=Db_Campos[x,4];
        InsCustomColumn.Tag:=StrToIntDef(Db_Campos[x,2],0);
        
        if Db_Campos[x,2]='0' then
          InsCustomColumn.Lines.Add('NO Definir');
        for I := 0 to ListaCol.Count - 1 do
        begin
          if Db_Campos[x,3]<>'' then
            if Db_Campos[x,3]=ListaCol[i] then
              if Db_Campos[x,2]='0' then
                IndexDefault:=i + 1
              else
                IndexDefault:=i;

          InsCustomColumn.Lines.Add(ListaCol[i]);
        end;
        InsCustomColumn.ItemIndex:=IndexDefault;
        Custom.Items.AddItem(nil,InsCustomColumn);
      end;
    except
      Res:=False;
    end;
  finally
    Grid.EndUpdate;
    ListaCol.Destroy;
  end;
  Result:=res;
end;


Function TFrmImportaProject.PreviewProject(var MsProject: Variant; var ActProject: Variant; Grid: TcxTreeList;PgbAvance: TcxProgressBar=nil):Boolean;
var
  Reng,I:Integer;
  Task:Variant;
  sWbs,CadTmp:String;
  iPos:Integer;
  Nodo:TcxTreeListNode;
  Padre:TcxTreeListNode;
  Res:Boolean;
begin
  Res:=true;
  Grid.BeginUpdate;
  Grid.Clear;
  try
    try
      for Reng:=0 to ActProject.Tasks.Count-1 do
      begin
        Task:=ActProject.Tasks.item[Reng+1];
        sWbs:=Task.wbs;
        iPos:=LastDelimiter('.',sWbs);
        CadTmp:='';
        if iPos>0 then
          CadTmp:=AnsiMidStr(sWbs,1,iPos-1);

        if CadTmp<>'' then
          Padre := FindMyNode(Grid,CadTmp,0);

        if Padre<>nil then
        begin
          with Grid.AddChild(Padre) do
          begin
            Values[0]:=swbs;
            for i := 1 to ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count do
              Values[i]:=Task.GetField(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);
          end;
        end
        else
        begin
          with Grid.Add do
          begin
            Values[0]:=swbs;
            for i := 1 to ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count do
              Values[i]:=Task.GetField(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);
          end;
        end;

        if PgbAvance<>nil then
        begin
          PgbAvance.Position := PgbAvance.Position + 1;
          Application.ProcessMessages;
        end;

      end;
    except
      Res:=false;
    end;
  finally
    grid.FullExpand;
    Grid.GotoBOF;
    Grid.EndUpdate;
  end;
  Result:=res;
end;


Function TFrmImportaProject.PreviewProject(var MsProject: Variant; var ActProject: Variant; Grid: TcxGridTableView):Boolean;
var
  Reng,I:Integer;
  Task:Variant;
  sWbs:String;
  iPos:Integer;
begin

  {
      //showmessage(MsProject.CustomFieldGetName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field));

      //MsProject.CustomFieldGetName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);

      with Grid.CreateColumn do
      begin
        Caption:=MsProject.FieldConstantToFieldName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);
        Width:=(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Width)*10;
        HeaderAlignmentHorz:=TaCenter;
        HeaderAlignmentVert:=TcxAlignmentVert(TaCenter);
        ListaCol.Add(Caption);
        Styles.Header := cxStyle1;
      end;




    end;}

  Grid.DataController.BeginUpdate;
  Grid.DataController.RecordCount:=0;
  for Reng:=0 to ActProject.Tasks.Count-1 do
  begin
    Grid.DataController.RecordCount:=Grid.DataController.RecordCount + 1;
    Task:=ActProject.Tasks.item[Reng+1];

    sWbs:=Task.wbs;
    iPos:=LastDelimiter('.',sWbs);
    if iPos>0 then
      sWbs:=AnsiMidStr(sWbs,1,iPos-1);
    Grid.DataController.Values[Grid.DataController.RecordCount-1,0]:= sWbs;

    for i := 1 to ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count do
    begin

      //showmessage((MsProject.FieldConstantToFieldName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field)));
      //showmessage(MsProject.FieldNameToFieldConstant(MsProject.FieldConstantToFieldName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field)));
      //showmessage(Task.GetField(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field));
      with Grid.DataController do
      begin

        Values[RecordCount-1,i]:=Task.GetField(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);

      //FieldConstantToFieldName(projectField)

      end;
    end;
  end;
 // Grid.DataController.Groups

  Grid.OptionsView.GroupSummaryLayout:=TcxGridGroupSummaryLayout(0);
  Grid.DataController.Groups.FullExpand;
  Grid.DataController.EndUpdate;
end;

Function TFrmImportaProject.LoadHeaderProject(var MsProject: Variant;var ActProject: Variant; Grid: TcxGridTableView; Custom: TNextInspector;PgbAvance: TcxProgressBar=nil):Boolean;
var
  InsCustomColumn:TNxComboBoxItem;
  i,x:Integer;
  Resource:Variant;
  ListaCol:TStringList;
  IndexDefault:Integer;
begin

  Custom.Items.Clear;
  Grid.ClearItems;

  ListaCol:=TStringList.Create;
  try
    with Grid.CreateColumn do
    begin
      Caption:='*';
      Width:=30;
      HeaderAlignmentHorz:=TaCenter;
      HeaderAlignmentVert:=TcxAlignmentVert(TaCenter);
    end;
    //showmessage(IntToStr(ActProject.TaskTableList.count));

    //showmessage(IntToStr(ActProject.TaskTables.count));
    //showmessage(IntToStr(ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count));

    for i := 1 to ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count do
    begin
      //showmessage(MsProject.CustomFieldGetName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field));

      //MsProject.CustomFieldGetName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);

      with Grid.CreateColumn do
      begin
        Caption:=MsProject.FieldConstantToFieldName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field);
        Width:=(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Width)*10;
        HeaderAlignmentHorz:=TaCenter;
        HeaderAlignmentVert:=TcxAlignmentVert(TaCenter);
        ListaCol.Add(Caption);
        Styles.Header := cxStyle1;
      end;




    end;

    for x := 1 to length(Db_Campos) do
    begin
      IndexDefault:=0;
      InsCustomColumn:=TNxComboBoxItem.Create(nil);
      InsCustomColumn.Caption:=Db_Campos[x,1];
      InsCustomColumn.Style:=cbDropDownList;
      InsCustomColumn.Alignment:=taCenter;
      InsCustomColumn.ItemHeight:=20;
      InsCustomColumn.Font.Style:=[FsBold];
      if Db_Campos[x,2]='0' then
        InsCustomColumn.Lines.Add('NO Definir');
      for I := 0 to ListaCol.Count - 1 do
      begin
        if Db_Campos[x,3]<>'' then
          if Db_Campos[x,3]=ListaCol[i] then
            if Db_Campos[x,2]='0' then
              IndexDefault:=i + 1
            else
              IndexDefault:=i;

        InsCustomColumn.Lines.Add(ListaCol[i]);
      end;
      InsCustomColumn.ItemIndex:=IndexDefault;
      Custom.Items.AddItem(nil,InsCustomColumn);
    end;

     // showmessage(MsProject.FieldConstantToFieldName(ActProject.TaskTables(ActProject.CurrentTable).TableFields[i].Field));

  //  begin
      //if ActProject.TaskTables[i].Name=ActProject.CurrentTable then
      //  for x := 1 to ActProject.TaskTables[i]. do

       // showmessage(ActProject.TaskViewList[i]);
      {CustomColumn:=TNxTextColumn.Create(nil);
      CustomColumn.Width:=200;
      CustomColumn.Display.AsString:=Resource.name;
      Grid.Columns.AddColumn(CustomColumn); }
    //end;
  finally

  end;
end;

procedure TFrmImportaProject.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:=CaFree;
end;

procedure TFrmImportaProject.FormShow(Sender: TObject);
begin
  QrContratos.Active:=false;
  QrContratos.Open;

  jDblCmbContrato.KeyValue:=global_Contrato;

  QrFolios.Active:=false;
  QrFolios.Open;

  jDblCmbContratoChange(Sender);
end;

Procedure TFrmImportaProject.PreImportOTProject(sFileProject:TFileName;Grid:TcxTreeList;Custom:TNextInspector;PgbAvance: TcxProgressBar=nil);
{$Region 'Informacion'}
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
{$EndRegion}
var
  ListaActividades: TStringList;
  MsProject:Variant;
  ActProject:Variant;
  Error:Boolean;
begin
  Error:=false;
  if FileExists(sFileProject) then
  begin
    if AnsiEndsText('.mpp',sFileProject) then
    begin
      ListaActividades:=TstringList.Create;
      try
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
          try
            MsProject.visible:=false;
            MsProject.DisplayAlerts:=false;
            if not error then
            begin
              MsProject.FileOpen(sFileProject);
              ActProject:=MsProject.ActiveProject;

              if PgbAvance<>nil then
              begin
                PgbAvance.Properties.min:=0;
                PgbAvance.Properties.Max := ActProject.Tasks.Count +  
                ActProject.TaskTables(ActProject.CurrentTable).TableFields.Count;
                PgbAvance.Position := 0;
                PgbAvance.Visible:=true;
                Application.ProcessMessages;
              end;

              if LoadHeaderProject(MsProject,ActProject,Grid,Custom,PgbAvance) then
                if PreviewProject(MsProject,ActProject,Grid,PgbAvance) then
                  btnImportar.Enabled:=true
                else
                  btnImportar.Enabled:=false;
            end;
          finally
            MsProject.visible:=true;
            MsProject.DisplayAlerts:=true;
            MsProject.Quit;
            if PgbAvance<>nil then
            begin
              PgbAvance.Visible:=False;
              Application.ProcessMessages;
            end;
          end;
        end;
      finally
        ListaActividades.Destroy;
      end;
    end
    else
      MessageDlg('El Archivo: '+ #13 + #10 + Quotedstr(sFileProject) + ' NO es Valido.',
                  MtError,[MbOk],0);
  end;
end;

procedure TFrmImportaProject.JFedtArchivoAfterDialog(Sender: TObject;
  var AName: string; var AAction: Boolean);
begin
  if AAction then
    PreImportOTProject(AName,CxTlsPrograma,NxInsColumnas,CxPbAvance);
end;

end.
