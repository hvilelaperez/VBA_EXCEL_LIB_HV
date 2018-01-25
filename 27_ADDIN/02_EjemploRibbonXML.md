# Ejemplo de Ribbon XML para Add-In

Definición de personalización del Ribbon de Excel mediante XML para un complemento de gestión de calidad de software.

## Código XML

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
<ribbon startFromScratch="false">
    <tabs>
      <tab id="Tab1" label="Software QLCL">
        <group id="Group1" label="Obras">
          <menu id="Menu1" label="Expediente" imageMso="OpenStartPage" size="large" itemSize="normal">
            <button id="BT1" onAction="CrearNuevaObra" label="Crear nueva obra" imageMso="FileNew"/>
            <button id="BT2" onAction="AbrirObra" label="Abrir obra" imageMso="FileOpen"/>
            <button id="BT3" onAction="GuardarObra" label="Guardar obra" imageMso="FileSave"/>
            <button id="BT4" onAction="ConfiguracionGeneral" label="Configuración general" imageMso="AddInManager"/>
          </menu>
          <menu id="Menu2" label="Información general" imageMso="AccessFormWizard" size="large" itemSize="normal">
            <button id="BT5" onAction="InfoObra" label="Información de la obra" imageMso="QueryShowTable"/>
            <button id="BT6" onAction="DatosEntrada" label="Datos de entrada" imageMso="FormulaMoreFunctionsMenu"/>
          </menu>
          <menu id="Menu3" label="Tabla de códigos aplicados" imageMso="ImportAccess" size="large" itemSize="normal">
            <button id="BT7" onAction="CodigoTrabajo" label="Código de trabajo" imageMso="QueryShowTable"/>
            <button id="BT8" onAction="CodigoEnsayo" label="Código de ensayo" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT9" onAction="EstándarRecepción" label="Estándar de recepción" imageMso="QueryShowTable"/>
            <button id="BT10" onAction="RefEnsayo" label="Referencia de ensayo" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT11" onAction="MetodoVerificacion" label="Método de verificación" imageMso="QueryShowTable"/>
          </menu>
        </group>
        <group id="Group2" label="Bitácora de construcción">
          <button id="BT12" onAction="Bitacora" label="Bitácora" imageMso="FileBackupDatabase" size="large"/>
          <button id="BT13" onAction="Clima" label="Clima, día libre" imageMso="PictureBrightnessGallery" size="large"/>
          <button id="BT14" onAction="VerificarDiaConstruccion" label="Verificar día de construcción" imageMso="FilePrepareMenu" size="normal"/>
          <button id="BT15" onAction="ExportarDatosActa" label="Exportar datos de acta" imageMso="DataValidation" size="normal"/>
          <button id="BT16" onAction="ExportarDatosBitacora" label="Exportar datos de bitácora" imageMso="ReviewTrackChanges" size="normal"/>
        </group>
        <group id="Group3" label="Lista de Actas">
          <button id="BT17" onAction="ListaEnsayoEntrada" label="Ensayo de entrada" imageMso="OrganizationChartSelectLevel" size="normal"/>
          <button id="BT18" onAction="ListaRecepcionTrabajo" label="Recepción de trabajo" imageMso="CreateFormInDesignView" size="normal"/>
          <button id="BT19" onAction="ListaEnsayoSitio" label="Ensayo en sitio" imageMso="ModuleInsert" size="normal"/>
          <button id="BT20" onAction="ListaRecepcionEtapa" label="Recepción por etapa" imageMso="FunctionsFinancialInsertGallery" size="normal"/>
          <menu id="Menu4" label="Otras actas" imageMso="FileSaveToDocumentManagementServer" size="normal" itemSize="normal">
            <button id="BT21" onAction="ListaSeguimientoPrevio" label="Seguimiento previo" imageMso="QueryShowTable"/>
            <button id="BT22" onAction="ListaVerificacionPerimetro" label="Verificación de perímetro" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT23" onAction="ListaVerificacionAltura" label="Verificación de altura" imageMso="QueryShowTable"/>
            <button id="BT24" onAction="ListaVerificacionOtra" label="Otra verificación" imageMso="FormulaMoreFunctionsMenu"/>
          </menu>
        </group>
        <group id="Group4" label="Imprimir Actas">
          <button id="BT25" onAction="ImprimirEnsayoEntrada" label="Ensayo de entrada" imageMso="OrganizationChartSelectLevel" size="normal"/>
          <button id="BT26" onAction="ImprimirRecepcionTrabajo" label="Recepción de trabajo" imageMso="CreateFormInDesignView" size="normal"/>
          <button id="BT27" onAction="ImprimirEnsayoSitio" label="Ensayo en sitio" imageMso="ModuleInsert" size="normal"/>
          <button id="BT28" onAction="ImprimirRecepcionEtapa" label="Recepción por etapa" imageMso="FunctionsFinancialInsertGallery" size="normal"/>
          <menu id="Menu5" label="Otras actas" imageMso="FileSaveToDocumentManagementServer" size="normal" itemSize="normal">
            <button id="BT29" onAction="ImprimirSeguimientoPrevio" label="Seguimiento previo" imageMso="QueryShowTable"/>
            <button id="BT30" onAction="ImprimirVerificacionPerimetro" label="Verificación de perímetro" imageMso="FormulaMoreFunctionsMenu"/>
            <button id="BT31" onAction="ImprimirVerificacionAltura" label="Verificación de altura" imageMso="QueryShowTable"/>
            <button id="BT32" onAction="ImprimirVerificacionOtra" label="Otra verificación" imageMso="FormulaMoreFunctionsMenu"/>
          </menu>
        </group>
        <group id="Group5" label="Bitácora de construcción">
          <button id="BT33" onAction="ConfigurarBitacora" label="Configurar bitácora" imageMso="ViewDocumentActionsPane" size="large"/>
          <button id="BT34" onAction="EscribirBitacora" label="Escribir bitácora" imageMso="FunctionsDateTimeInsertGallery" size="large"/>
        </group>
        <group id="Group6" label="Utilidades varias">
          <menu id="Menu6" label="Manual - Autor" imageMso="FunctionsLogicalInsertGallery" size="large" itemSize="normal">
            <button id="BT35" onAction="ManualUso" label="Manual de uso" imageMso="FunctionsLookupReferenceInsertGallery"/>
            <button id="BT36" onAction="VideoManual" label="Video manual de uso" imageMso="ViewSlideShowView"/>
            <button id="BT37" onAction="AcercaDe" label="Acerca de" imageMso="ContactPictureMenu"/>
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

## Descripción

| Elemento | Descripción |
|----------|-------------|
| `customUI` | Elemento raíz de la interfaz personalizada de Office |
| `ribbon` | Define el Ribbon de Excel (no arranca desde cero) |
| `tabs` | Contenedor de pestañas personalizadas |
| `tab` | Pestaña principal: "Software QLCL" |
| `group` | Grupos de comandos dentro de la pestaña |
| `menu` | Menús desplegables con múltiples opciones |
| `button` | Botones individuales con acciones asociadas |

### Estructura de Grupos

| Grupo | Contenido |
|-------|-----------|
| Obras | Gestión de expedientes: crear, abrir, guardar, configuración general, información y códigos |
| Bitácora de construcción | Registro diario: clima, verificación de días, exportación de datos |
| Lista de Actas | Listados de ensayos, recepciones y verificaciones |
| Imprimir Actas | Impresión de actas por tipo |
| Bitácora de construcción | Configuración y escritura de bitácora |
| Utilidades varias | Manuales, videos de ayuda y acerca de |

### Funciones de callbacks referenciadas

| Callback | Descripción |
|----------|-------------|
| `CrearNuevaObra` | Crea una nueva obra/proyecto |
| `AbrirObra` | Abre una obra existente |
| `GuardarObra` | Guarda el estado actual de la obra |
| `ConfiguracionGeneral` | Accede a la configuración del sistema |
| `InfoObra` | Muestra información general de la obra |
| `DatosEntrada` | Gestiona datos de entrada |
| `Bitacora` | Accede a la bitácora de construcción |
| `Clima` | Registra clima y días libres |
| `ManualUso` | Muestra el manual de usuario |
| `AcercaDe` | Información del programa |

---

*Módulo: 26 - VBA ADDIN | Archivo: Sample xml Ribon.txt*
