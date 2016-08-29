program inteligent;
uses
  Forms,
  frm_connection in 'frm_connection.pas' {connection: TDataModule},
  frm_inteligent in 'frm_inteligent.pas' {frmInteligent},
  global in 'global.pas',
  frm_setup in 'frm_setup.pas' {frmSetup},
  frm_SubContratos in 'frm_SubContratos.pas' {frmSubContratos},
  frm_barra in 'frm_barra.pas' {frmBarra: TFrame},
  frm_equipos in 'frm_equipos.pas' {frmEquipos},
  frm_trinomios in 'frm_trinomios.pas' {frmTrinomios},
  frm_turnos in 'frm_turnos.pas' {frmTurnos},
  Utilerias in 'Utilerias.pas',
  frm_gruposxprograma in 'frm_gruposxprograma.pas' {frmGruposxPrograma},
  frm_Consumibles in 'frm_Consumibles.pas' {frmConsumibles},
  frm_acerca in 'frm_acerca.pas' {frmAcerca},
  frm_proveedores in 'frm_proveedores.pas' {frmProveedores},
  frm_deptos in 'frm_deptos.pas' {frmDeptos},
  frm_TipoMovto in 'frm_TipoMovto.pas' {frmMovtos},
  frm_GruposUsuarios in 'frm_GruposUsuarios.pas' {frmGrupos},
  frm_CalculoAvancesxPartida in 'frm_CalculoAvancesxPartida.pas' {frmCalculoAvancesxPartida},
  frm_tramitedepermisos in 'frm_tramitedepermisos.pas' {frmTramitedePermisos},
  frm_DistribucionPrograma in 'frm_DistribucionPrograma.pas' {frmDistribucionPrograma},
  frm_actividadesxgrupo in 'frm_actividadesxgrupo.pas' {frmActividadesxGrupo},
  frm_ConsultadeActividades4 in 'frm_ConsultadeActividades4.pas' {frmConsultaActividad4},
  frm_AjustaOrden in 'frm_AjustaOrden.pas' {frmAjustaOrden},
  frm_CalculoAvancesPaquetes in 'frm_CalculoAvancesPaquetes.pas' {frmCalculoAvancesPaquetes},
  frm_personal in 'frm_personal.pas' {frmPersonal},
  frm_paquetesdeequipo in 'frm_paquetesdeequipo.pas' {frmPaqueteEquipo},
  frm_ordenesPerf in 'frm_ordenesPerf.pas' {frmOrdenesPerf},
  frm_AvancesFinancieros in 'frm_AvancesFinancieros.pas' {frmAvancesFinancieros},
  frm_AvisosAlertas in 'frm_AvisosAlertas.pas' {frmAvisosAlertas},
  frm_Reprogramacion in 'frm_Reprogramacion.pas' {frmReprogramacion},
  frm_comentariosxanexo in 'frm_comentariosxanexo.pas' {frmComentariosxAnexo},
  frm_diasfestivos in 'frm_diasfestivos.pas' {frmDiasFestivos},
  frm_ConsultaxDescripcion in 'frm_ConsultaxDescripcion.pas' {frmConsultaxDescripcion},
  frm_recursosxanexo in 'frm_recursosxanexo.pas' {frmRecursosxAnexo},
  frm_OrdendeCambio in 'frm_OrdendeCambio.pas' {frmOrdendeCambio},
  frm_retecionesypenas in 'frm_retecionesypenas.pas' {frmRetencionesyPenas},
  frm_activos in 'frm_activos.pas' {frmActivos},
  frm_factordecosto in 'frm_factordecosto.pas' {frmFactordeCosto},
  frm_compara2 in 'frm_compara2.pas' {frmComparativo2},
  frm_SqlExportar in 'frm_SqlExportar.pas' {frmSqlExportar},
  frm_SqlManager in 'frm_SqlManager.pas' {frmSqlManager},
  frm_importaciondedatos in 'frm_importaciondedatos.pas' {frmImportaciondeDatos},
  frm_platicas in 'frm_platicas.pas' {frmPlaticas},
  frm_programas in 'frm_programas.pas' {frmProgramas},
  frm_contratosxordenes in 'frm_contratosxordenes.pas' {frmOrdenesxUsuario},
  frm_personalprogramado in 'frm_personalprogramado.pas' {frmPersonalProgramado},
  frm_generado in 'frm_generado.pas' {frmGenerado},
  frm_bitacoradepartamental_2 in 'frm_bitacoradepartamental_2.pas' {frmBitacoraDepartamental_2},
  frm_bitacoraxalcance in 'frm_bitacoraxalcance.pas' {frmBitacoraxAlcance},
  frm_PendientesNew in 'frm_PendientesNew.pas' {frmPendientesNew},
  frm_ConsultadeActividades5 in 'frm_ConsultadeActividades5.pas' {frmConsultaActividad5},
  frm_SqlImportar in 'frm_SqlImportar.pas' {frmSqlImportar},
  frm_ControlDirecto in 'frm_ControlDirecto.pas' {frmControlDirecto},
  frm_ValidaEstimacion in 'frm_ValidaEstimacion.pas' {frmValidaEstimacion},
  frm_ConsultadeActividades2 in 'frm_ConsultadeActividades2.pas' {frmConsultaActividad2},
  frm_ConsultadeActividades in 'frm_ConsultadeActividades.pas' {frmConsultaActividad},
  frm_Kardex in 'frm_Kardex.pas' {frmKardex},
  frm_seguridad in 'frm_seguridad.pas' {frmSeguridad},
  frm_cambiapassword in 'frm_cambiapassword.pas' {frmcambiopassword},
  frm_ReportePeriodo in 'frm_ReportePeriodo.pas' {frmReportePeriodo},
  frm_acceso in 'frm_acceso.pas' {frmacceso},
  frm_Proyeccion in 'frm_Proyeccion.pas' {frmProyeccion},
  frm_FichaTecnica in 'frm_FichaTecnica.pas' {frmFichaTecnica},
  frm_personalconsolidado in 'frm_personalconsolidado.pas' {frmPersonalConsolidado},
  frm_SintesisGerencial in 'frm_SintesisGerencial.pas' {frmSintesisGerencial},
  frm_firmantes in 'frm_firmantes.pas' {frmFirmas},
  frm_ConsultadeActividades3 in 'frm_ConsultadeActividades3.pas' {frmConsultaActividad3},
  frmReporteDiarioGerencial in 'frmReporteDiarioGerencial.pas' {frmReporteGerencial},
  frm_warningdia in 'frm_warningdia.pas' {frmWarningDia},
  frm_tipsdia in 'frm_tipsdia.pas' {frmTipsDia},
  frm_abrereporte in 'frm_abrereporte.pas' {frmAbreReporte},
  frm_usuarios in 'frm_usuarios.pas' {frmUsuarios},
  frm_residencias in 'frm_residencias.pas' {frmResidencias},
  frm_ExportaGeneral in 'frm_ExportaGeneral.pas' {frmExportaGeneral},
  frm_ImportarDiarios in 'frm_ImportarDiarios.pas' {frmImportarDiarios},
  frmFiltroInteligent in 'frmFiltroInteligent.pas' {frmFiltros},
  frm_EstimaProveedor in 'frm_EstimaProveedor.pas' {frmEstimaProveedor},
  frm_ProcRegAvFisico in 'frm_ProcRegAvFisico.pas' {frmProcRegAvFisico},
  frm_Generadores_Barco in 'frm_Generadores_Barco.pas' {frmGeneradoresBarco},
  frm_EstimacionAlbum in 'frm_EstimacionAlbum.pas' {frmEstimacionAlbum},
  UnitEstimacion in 'UnitEstimacion.pas',
  frm_seleccion2 in 'frm_seleccion2.pas' {frmSeleccion2},
  frm_Cuentas in 'frm_Cuentas.pas' {frmCuentas},
  frm_ActividadesAnexo2 in 'frm_ActividadesAnexo2.pas' {frmActividadesAnexo2},
  frm_admonCatalogos in 'frm_admonCatalogos.pas' {frmAdmonCatalogos},
  frm_compara in 'frm_compara.pas' {frmComparativo},
  frm_comparativo in 'frm_comparativo.pas' {frmCompara},
  frm_EstimaInstalado in 'frm_EstimaInstalado.pas' {frmEstimaInstalado},
  inteligent_TLB in 'inteligent_TLB.pas',
  frm_AjustaAnexo in 'frm_AjustaAnexo.pas' {frmAjustaAnexo},
  frm_contratos in 'frm_contratos.pas' {frmContratos},
  frm_Pedidos in 'frm_Pedidos.pas' {frmPedidos},
  StoHtmlHelp in 'StoHtmlHelp.pas',
  frm_ConfiguraMail in 'frm_ConfiguraMail.pas' {frmConfiguraMail},
  frm_detalledeinstalacion in 'frm_detalledeinstalacion.pas' {frmDetalledeInstalacion},
  frm_Proyeccion2 in 'frm_Proyeccion2.pas' {frmProyeccion2},
  frm_empleados in 'frm_empleados.pas' {frmEmpleados},
  frm_contratosxusuario in 'frm_contratosxusuario.pas' {frmContratosxUsuario},
  frm_PrintReportesDiarios in 'frm_PrintReportesDiarios.pas' {frmPrintReportesDiarios},
  frm_tripulacion in 'frm_tripulacion.pas' {frmTripulacion},
  frm_cuadredepersonal in 'frm_cuadredepersonal.pas' {frmCuadredePersonal},
  frm_partidasxisometrico in 'frm_partidasxisometrico.pas' {frmPartidasxIsometrico},
  frm_Conversiones in 'frm_Conversiones.pas' {frmConversiones},
  frm_ReporteDiarioTurno2 in 'frm_ReporteDiarioTurno2.pas' {frmDiarioTurno2},
  frm_prorrateoPernocta in 'frm_prorrateoPernocta.pas' {frmProrrateoPernocta},
  frm_Fases in 'frm_Fases.pas' {frmFases},
  frm_AdmonyTiempos in 'frm_AdmonyTiempos.pas' {frmAdmonyTiempos},
  frm_ReporteDiario_Barco in 'frm_ReporteDiario_Barco.pas' {frmDiarioBarco},
  frm_tiposdeMovimiento in 'frm_tiposdeMovimiento.pas' {frmTiposdeMovimiento},
  frm_ordenes in 'frm_ordenes.pas' {frmOrdenes},
  frm_ordenesGral in 'frm_ordenesGral.pas' {frmOrdenesGeneral},
  Frm_DetalleOficioPersonal in 'Frm_DetalleOficioPersonal.pas' {FrmDetalleOficioPersonal},
  Unit_Barras in 'Unit_Barras.pas',
  frm_CambioContrato in 'frm_CambioContrato.pas' {frmCambioContrato},
  frm_detalletiemposmuertos in 'frm_detalletiemposmuertos.pas' {frmdetallestiemposmuertos},
  frm_anexosCotemar in 'frm_anexosCotemar.pas' {frmAnexosCotemar},
  frm_GruposPersonal in 'frm_GruposPersonal.pas' {frmGruposPersonal},
  frm_Propiedades in 'frm_Propiedades.pas' {FrmPropiedades},
  frm_IntelChart in 'frm_IntelChart.pas' {IntelChart},
  frm_Graficador in 'frm_Graficador.pas' {frmGraficador},
  UnitExcel in 'UnitExcel.pas',
  UTChartMarco in 'UTChartMarco.pas',
  UTChartMouse in 'UTChartMouse.pas',
  frm_grupofamilias in 'frm_grupofamilias.pas' {frmGrupoFamilias},
  frm_Almacenes in 'frm_Almacenes.pas' {frmAlmacenes},
  frm_actividades in 'frm_actividades.pas' {frmActividades},
  frm_bitacoraOptativa in 'frm_bitacoraOptativa.pas' {frmBitacoraOptativa},
  frm_EntradaAlmacen in 'frm_EntradaAlmacen.pas' {frmEntradaAlmacen},
  frm_SalidaAlmacen in 'frm_SalidaAlmacen.pas' {frmSalidaAlmacen},
  frm_AltaServidor in 'frm_AltaServidor.pas' {frmAltaServidor},
  UnitFactorPeriodo in 'UnitFactorPeriodo.pas',
  frm_RequisicionPerf in 'frm_RequisicionPerf.pas' {frmRequisicionPerf},
  UnitTSintaxFormulizer in 'UnitTSintaxFormulizer.pas',
  UnitTSemanticFormulizer in 'UnitTSemanticFormulizer.pas',
  frm_embarcaciones in 'frm_embarcaciones.pas' {frmEmbarcaciones},
  UReporteDiarioMix in 'UReporteDiarioMix.pas',
  UIni in 'UIni.pas',
  frm_DetalleCaptura in 'frm_DetalleCaptura.pas' {DetalleCaptura},
  frm_EditaEstimacion in 'frm_EditaEstimacion.pas' {frmEditaEstimacion},
  frm_OpcionesAvances in 'frm_OpcionesAvances.pas' {frmOpcionesAvances},
  frm_estimaciones in 'frm_estimaciones.pas' {frmEstimaciones},
  frm_EstimacionGeneral in 'frm_EstimacionGeneral.pas' {frmEstimacionGeneral},
  frm_estimacionAnterior in 'frm_estimacionAnterior.pas' {frmEstimacionAnterior},
  frm_valida in 'frm_valida.pas' {frmValida},
  frm_plataformas in 'frm_plataformas.pas' {frmPlataformas},
  frm_AdministrarBd in 'frm_AdministrarBd.pas' {FrmAdministrarBd},
  frm_estimacionOrdenes in 'frm_estimacionOrdenes.pas' {frmEstimacionOrdenes},
  frm_estimacionRecursosPT in 'frm_estimacionRecursosPT.pas' {frmEstimacionRecursosPT},
  frm_EstimacionDetalleAdicional in 'frm_EstimacionDetalleAdicional.pas' {frmEstimacionDetalleAdicional},
  frm_estimacionAdicional in 'frm_estimacionAdicional.pas' {frmEstimacionAdicional},
  frm_ValidaEstimacionGral in 'frm_ValidaEstimacionGral.pas' {frmValidaEstimacionGral},
  frm_BusquedadeNotas in 'frm_BusquedadeNotas.pas' {frmBuscaComentarios},
  frm_cancelacion in 'frm_cancelacion.pas' {frmCancelacion},
  frm_AperturaEstimacionGral in 'frm_AperturaEstimacionGral.pas' {frmAperturaEstimacionGral},
  frm_estimacionAvances in 'frm_estimacionAvances.pas' {frmEstimacionAvances},
  masUtilerias in 'masUtilerias.pas',
  UnitExcepciones in 'UnitExcepciones.pas',
  USelCol in 'USelCol.pas',
  frm_catalogoerrores in 'frm_catalogoerrores.pas' {frmCatalogoErrores},
  frm_Pernoctan in 'frm_Pernoctan.pas' {frmPernoctan},
  UDbGrid in 'UDbGrid.pas',
  frm_ActualizaAvancesRemotos in 'frm_ActualizaAvancesRemotos.pas' {frmActualizacionRemota},
  UFunctionsGHH in 'UFunctionsGHH.pas',
  frm_servidor in 'frm_servidor.pas' {frmdatos},
  Frm_FiltroIsometricos in 'Frm_FiltroIsometricos.pas' {FrmFiltroIsometricos},
  frm_entradaanex in 'frm_entradaanex.pas' {frmentradaanex},
  unitmanejofondo in 'unitmanejofondo.pas',
  UnitValidaTexto in 'UnitValidaTexto.pas',
  frm_GraficaGerencialDX in 'frm_GraficaGerencialDX.pas' {frmGraficaGerencialDX},
  frm_PopUpPaquetes_p in 'frm_PopUpPaquetes_p.pas' {frmPopUpPaquetes_p},
  unitActivaPop in 'unitActivaPop.pas',
  frm_controlEmpleados in 'frm_controlEmpleados.pas' {frmControlEmpleados},
  frm_controlEmpleados2 in 'frm_controlEmpleados2.pas' {frmControlEmpleados2},
  frm_ReporteDiarioTurnoTierra in 'frm_ReporteDiarioTurnoTierra.pas' {frmDiarioTurnoTierra},
  UnitValidacion in 'UnitValidacion.pas',
  frm_graficaexplosion in 'frm_graficaexplosion.pas' {frmGraficaExplosion},
  UnitTablasImpactadas in 'UnitTablasImpactadas.pas',
  frm_DespieceDX in 'frm_DespieceDX.pas' {frmDespieceDX},
  frm_DespieceImagen in 'frm_DespieceImagen.pas' {frmDespieceImagen},
  frm_paquetesdepersonal in 'frm_paquetesdepersonal.pas' {frmPaquetePersonal},
  UnitTBotonesPermisos in 'UnitTBotonesPermisos.pas' {,
  FrmMovtoPersonalxoficio in 'FrmMovtoPersonalxoficio.pas' {FrmMovtosPersonalxoficio},
  FrmMovtoPersonalxoficio in 'FrmMovtoPersonalxoficio.pas' {FrmMovtosPersonalxoficio},
  frm_sincinformes in 'frm_sincinformes.pas' {frmInformeSincronizacion},
  frm_OpcionesReporteProduccion in 'frm_OpcionesReporteProduccion.pas' {frmOpcionesReporteProduccion},
  frm_OpcionesGerencial in 'frm_OpcionesGerencial.pas' {frmOpcionesGerencial},
  frm_lista_personal in 'frm_lista_personal.pas' {frmLista_personal},
  frm_tripulacion_diaria in 'frm_tripulacion_diaria.pas' {frmTripulacionDiaria},
  Frm_CuadreXPartida in 'Frm_CuadreXPartida.pas' {FrmCuadreXPartida},
  UTSuperPanel in 'UTSuperPanel.pas',
  Frm_Moe in 'Frm_Moe.pas' {FrmMoe},
  frm_MovimientosLogisticos in 'frm_MovimientosLogisticos.pas' {frmMovimientosLogisticos},
  frm_ConsumodeCombustible in 'frm_ConsumodeCombustible.pas' {frmConsumodeCombustible},
  frm_NotasGenerales in 'frm_NotasGenerales.pas' {frmNotasGenerales},
  UnitPatrick in 'UnitPatrick.pas',
  frm_gruposdepersonal in 'frm_gruposdepersonal.pas' {frmGruposdePersonal},
  frm_oficiosdemovimientos in 'frm_oficiosdemovimientos.pas' {frmOficiosDeMovimientos},
  frm_oficiosdemovimientos_detalles in 'frm_oficiosdemovimientos_detalles.pas' {frmOficiosDeMovimientos_detalles},
  frm_Condicionesclima in 'frm_Condicionesclima.pas' {frmCondicionesclima},
  frm_moduloadmonpersonal in 'frm_moduloadmonpersonal.pas' {frmModuloAdmonPersonal},
  frm_gruposdeequipo in 'frm_gruposdeequipo.pas' {frmGruposdeEquipo},
  frm_moduloreportegerencial in 'frm_moduloreportegerencial.pas' {frmModuloReporteGerencial},
  frm_bitacora2 in 'frm_bitacora2.pas' {frmBitacora2},
  Frm_ResumenPersonal in 'Frm_ResumenPersonal.pas' {FrmResumenPersonal},
  frm_CatNomFirmantes in 'frm_CatNomFirmantes.pas' {frmcatnomfirmates},
  frm_unificadorequipos in 'frm_unificadorequipos.pas' {FrmUnificadorEquipos},
  Frm_ImportaExportaActiv in 'Frm_ImportaExportaActiv.pas' {FrmImportaExportaActiv},
  frm_ProcesaGenerador in 'frm_ProcesaGenerador.pas' {frmProcesaGenerador},
  frm_formatos in 'frm_formatos.pas' {frmFormatos},
  UnitTarifa_Calidad in 'UnitTarifa_Calidad.pas',
  frm_importacuadre in 'frm_importacuadre.pas' {frmImportaCuadre},
  frm_lista_personalV2 in 'frm_lista_personalV2.pas' {frmListaPersonalV2},
  Frm_BuscaPersonal in 'Frm_BuscaPersonal.pas' {FrmBuscaPersonal},
  Frm_EligeFecha in 'Frm_EligeFecha.pas' {FrmEligeFecha},
  Frm_PopUpImportacionPP in 'Frm_PopUpImportacionPP.pas' {FrmPopUpImportacionPP},
  Frm_ImportaProject in 'Frm_ImportaProject.pas' {FrmImportaProject},
  Frm_NotaCampoObservaciones in 'Frm_NotaCampoObservaciones.pas' {FrmNotaCampoObservaciones},
  unt_Actividades in 'unt_Actividades.pas',
  frm_cuadre in 'frm_cuadre.pas' {frmCuadre},
  UnitTarifa in 'UnitTarifa.pas',
  frm_CuadreXCategoria in 'frm_CuadreXCategoria.pas' {frmCuadreCategoria},
  UnitCuadre in 'UnitCuadre.pas',
  UnitMetodos in 'UnitMetodos.pas',
  frm_ReprogramacionFolio in 'frm_ReprogramacionFolio.pas' {frmReprogramacionFolio},
  Frm_PopUpReprogramacion in 'Frm_PopUpReprogramacion.pas' {FrmPopUpReprogramacion},
  UFrmRecursosTierra in 'UFrmRecursosTierra.pas' {FrmRecursosTierra},
  Frm_NotaCampo in 'Frm_NotaCampo.pas' {FrmNotaCampo},
  Frm_generadores in 'Frm_generadores.pas' {FrmGeneradores},
  frm_ReporteDiarioTurno in 'frm_ReporteDiarioTurno.pas' {frmDiarioTurno},
  frm_bitacoradepartamental_Tierra in 'frm_bitacoradepartamental_Tierra.pas' {frmBitacoraDepartamental_Tierra},
  frm_cuadre_normal in 'frm_cuadre_normal.pas' {frmCuadreNormal},
  Frm_Materiales in 'Frm_Materiales.pas' {FrmAltaMAterial},
  Frm_Consultas in 'Frm_Consultas.pas' {FrmConsultas},
  frm_AjustesDiarios in 'frm_AjustesDiarios.pas' {frmAjustesDiarios};

//frm_PopUpImportacion in 'frm_PopUpImportacion.pas' {FrmPopUpImportacion};

{$R *.TLB}

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Inteligent - Tarifa Diaria';
  Application.HelpFile := '';
  Application.CreateForm(Tconnection, connection);
  Application.CreateForm(TfrmInteligent, frmInteligent);
  Application.CreateForm(Tfrmacceso, frmacceso);
  Application.CreateForm(TfrmSeguridad, frmSeguridad);
  Application.CreateForm(TfrmAcerca, frmAcerca);
  Application.CreateForm(TfrmSeleccion2, frmSeleccion2);
  Application.CreateForm(TfrmSetup, frmSetup);
  Application.Run;
end.



