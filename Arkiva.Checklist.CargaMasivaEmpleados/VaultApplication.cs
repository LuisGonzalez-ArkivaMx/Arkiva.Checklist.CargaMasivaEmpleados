using System;
using System.Diagnostics;
using MFiles.VAF;
using MFiles.VAF.Common;
using MFiles.VAF.Configuration;
using MFiles.VAF.Core;
using MFilesAPI;
using System.Runtime.Serialization;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Runtime.InteropServices;

namespace Arkiva.Checklist.CargaMasivaEmpleados
{
    [DataContract]
    public class Configuration
    {
        // NOTE: The default value needs to be placed in both the JsonConfEditor
        // (or derived) attribute, and as a default value on the member.
        [DataMember]
        [JsonConfEditor(DefaultValue = "Value 1")]
        public string ConfigValue1 = "Value 1";

    }

    /// <summary>
    /// The entry point for this Vault Application Framework application.
    /// </summary>
    /// <remarks>Examples and further information available on the developer portal: http://developer.m-files.com/. </remarks>
    public partial class VaultApplication
        : ConfigurableVaultApplicationBase<Configuration>
    {
        #region Eventhandler

        [EventHandler(MFEventHandlerType.MFEventHandlerAfterFileUpload, Class = "CL.CreacionMasivaDeEmpleadosProveedor")]
        [EventHandler(MFEventHandlerType.MFEventHandlerAfterFileUpload, Class = "CL.CreacionMasivaDeEmpleadosEmpresaInterna")]
        public void CargaMasivaEmpleados(EventHandlerEnvironment env)
        {
            // Files that we should clean up.
            var filesToDelete = new List<string>();

            try
            {
                int iTipoEmpleado = 0;
                int iObjeto = 0;
                int iClase = 0;
                //var pd_RfcEmpresa = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
                var pd_EstadoProcesamiento = env.Vault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EstadoDeProcesamiento");
                var cl_CargaMasivaEmpleadosProveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.CreacionMasivaDeEmpleadosProveedor");
                var cl_CargaMasivaEmpleadosEmpresaInterna = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.CreacionMasivaDeEmpleadosEmpresaInterna");
                var cl_Proveedor = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.Proveedor");
                var cl_EmpresaInterna = env.Vault.ClassOperations.GetObjectClassIDByAlias("CL.EmpresaInterna");                
                var ot_Proveedor = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.Proveedor");
                var ot_EmpresaInterna = env.Vault.ObjectTypeOperations.GetObjectTypeIDByAlias("OT.EmpresaInterna");

                var oPropertyValues = new PropertyValues();                
                oPropertyValues = env.Vault.ObjectPropertyOperations.GetProperties(env.ObjVer);

                //var iClass = oPropertyValues
                //    .SearchForProperty((int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass)
                //    .TypedValue
                //    .GetLookupID();

                var pdClass = env.Vault.PropertyDefOperations.GetBuiltInPropertyDef(MFBuiltInPropertyDef.MFBuiltInPropertyDefClass);
                var iClassID = oPropertyValues.SearchForPropertyEx(pdClass.ID, true).TypedValue.GetLookupID();

                ObjectClass oObjectClass = env.Vault.ClassOperations.GetObjectClass(iClassID);
                var sClassName = oObjectClass.Name;

                //var sRFC = oPropertyValues
                //    .SearchForPropertyEx(pd_RfcEmpresa, true)
                //    .TypedValue
                //    .Value
                //    .ToString();

                if (iClassID == cl_CargaMasivaEmpleadosEmpresaInterna) // Empleado Interno
                {
                    iTipoEmpleado = 1;

                    iObjeto = ot_EmpresaInterna;
                    iClase = cl_EmpresaInterna;
                }
                else if (iClassID == cl_CargaMasivaEmpleadosProveedor) // Contacto Externo SE
                {
                    iTipoEmpleado = 2;

                    iObjeto = ot_Proveedor;
                    iClase = cl_Proveedor;
                }

                var oObjVerEx = env.ObjVerEx;

                var oObjectFiles = oObjVerEx.Info.Files;

                IEnumerator enumerator = oObjectFiles.GetEnumerator();

                while (enumerator.MoveNext())
                {
                    bool bProcesoExitoso = false;
                    int iEstadoProcesamiento = 0;

                    ObjectFile oFile = (ObjectFile)enumerator.Current;

                    string sFilePath = SysUtils.GetTempFileName(".tmp"); // @"C:\TempMFilesVAF\TempFile.csv"; // 

                    // This must be generated from the temporary path and GetTempFileName. 
                    // It cannot contain the original file name.
                    filesToDelete.Add(sFilePath);

                    // Gets the latest version of the specified file
                    FileVer fileVer = oFile.FileVer;

                    // Download the file to a temporary location
                    env.Vault.ObjectFileOperations.DownloadFile(oFile.ID, fileVer.Version, sFilePath);

                    // Buscar o crear Empresa Interna o Proveedor
                    //BuscarOCrearProveedorOEmpresaInterna(iObjeto, iClase, sRFC, pd_RfcEmpresa);

                    //string sStringFormat = Path.Combine();

                    //string sFileName = Path.GetFileName(sFilePath);

                    //string sNewFilePath = @"C:\TempMFilesVAF\" + sFileName;

                    //File.Copy(sFilePath, sNewFilePath);

                    //string sExcelFile = Path.ChangeExtension(sNewFilePath, ".xlsx");

                    //filesToDelete.Add(sNewFilePath);

                    // Abrir el archivo CSV
                    StreamReader srArchivoCsvEmpleados = new StreamReader(sFilePath);
                    string sDelimitador = ",";
                    string sLinea;

                    // Leer la primera linea del CSV para descartarla porque es el encabezado
                    srArchivoCsvEmpleados.ReadLine();

                    while ((sLinea = srArchivoCsvEmpleados.ReadLine()) != null)
                    {
                        // Datos del empleado
                        string szNombreCompleto = "";
                        //string szApellidoPaterno = "";
                        //string szApellidoMaterno = "";
                        //string szTipoEmpleado = "";
                        string szEmail = "";
                        string szRFCEmpresa = "";
                        string szCURP = "";                        

                        string[] sRegistroPorFila = sLinea.Split(Convert.ToChar(sDelimitador));
                        szNombreCompleto = sRegistroPorFila[0];
                        //szApellidoPaterno = sRegistroPorFila[1];
                        //szApellidoMaterno = sRegistroPorFila[2];
                        //szTipoEmpleado = sRegistroPorFila[3];
                        szEmail = sRegistroPorFila[1];
                        szRFCEmpresa = sRegistroPorFila[2];
                        szCURP = sRegistroPorFila[3];

                        //SysUtils.ReportInfoToEventLog("Linea: 150, " + "Nombre: " + szNombre + ", " + "CURP: " + szCURP);

                        //if (szTipoEmpleado == "Internal")
                        //{
                        //    iTipoEmpleado = 1;

                        //    iObjeto = ot_EmpresaInterna;
                        //    iClase = cl_EmpresaInterna;
                        //}
                        //else // Empleado Externo
                        //{
                        //    iTipoEmpleado = 2;

                        //    iObjeto = ot_Proveedor;
                        //    iClase = cl_Proveedor;
                        //}

                        // Invocar metodo para crear empleado
                        if (CreateEmpleadoInternoOExterno(szNombreCompleto, iTipoEmpleado, szEmail, szRFCEmpresa, szCURP, iObjeto, iClase, sClassName) == true)
                        {
                            bProcesoExitoso = true;
                        }
                    }

                    srArchivoCsvEmpleados.Close();
                    srArchivoCsvEmpleados.Dispose();                             

                    // Abrir archivo excel
                    //Application oExcelFile = new Application();
                    //Workbook wb = oExcelFile.Workbooks.Open(sFilePath);
                    //Worksheet ws = wb.Sheets[1];
                    //ws.Activate();
                    //Range oRangeColumns = ws.UsedRange;

                    //int rowCount = oRangeColumns.Rows.Count;
                    //int colCount = oRangeColumns.Columns.Count;
                    
                    //for (int i = 2; i <= rowCount; i++)
                    //{
                    //    // Datos del empleado
                    //    string szNombre = "";
                    //    string szApellidos = "";
                    //    string szTipoEmpleado = "";
                    //    string szEmail = "";
                    //    string szRFCEmpresa = "";
                    //    string szCURP = "";
                    //    int iTipoEmpleado = 0;

                    //    szNombre = oRangeColumns.Cells[i, 1].Value2.ToString();
                    //    szApellidos = oRangeColumns.Cells[i, 2].Value2.ToString();
                    //    szTipoEmpleado = oRangeColumns.Cells[i, 3].Value2.ToString();
                    //    szEmail = oRangeColumns.Cells[i, 4].Value2.ToString();
                    //    szRFCEmpresa = oRangeColumns.Cells[i, 5].Value2.ToString();
                    //    szCURP = oRangeColumns.Cells[i, 6].Value2.ToString();

                    //    SysUtils.ReportInfoToEventLog("Linea: 145, " + "Nombre: " + szNombre + ", " + "CURP: " + szCURP);

                    //    if (szTipoEmpleado == "Internal")
                    //    {
                    //        iTipoEmpleado = 1;

                    //        iObjeto = ot_EmpresaInterna;
                    //        iClase = cl_EmpresaInterna;
                    //    }                            
                    //    else // Empleado Externo
                    //    {
                    //        iTipoEmpleado = 2;

                    //        iObjeto = ot_Proveedor;
                    //        iClase = cl_Proveedor;
                    //    }                        

                    //    // Invocar metodo para crear empleado
                    //    if (CreateEmpleadoInternoOExterno(szNombre, szApellidos, iTipoEmpleado, szEmail, szRFCEmpresa, szCURP, iObjeto, iClase) == true)
                    //    {
                    //        bProcesoExitoso = true;
                    //    }                        
                    //}

                    // Limpiar objetos
                    //GC.Collect();
                    //GC.WaitForPendingFinalizers();

                    //// Liberar objetos para matar el proceso Excel que esta corriendo por detras del sistema
                    //Marshal.ReleaseComObject(oRangeColumns);
                    //Marshal.ReleaseComObject(ws);

                    //// Cerrar y liberar
                    //wb.Close();
                    //Marshal.ReleaseComObject(wb);

                    //// Quitar y liberar
                    //oExcelFile.Quit();
                    //Marshal.ReleaseComObject(oExcelFile);

                    // Fin del proceso de carga masiva de empleado
                    if (bProcesoExitoso == true)
                    {
                        // Documento Procesado con exito
                        iEstadoProcesamiento = 1;
                    }
                    else
                    {
                        // No Procesado o Termino en Error
                        iEstadoProcesamiento = 2;
                    }

                    // Actualizacion de estado de procesamiento
                    var oLookup = new Lookup();
                    var oObjID = new ObjID();

                    oObjID.SetIDs
                    (
                        ObjType: (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument,
                        ID: env.ObjVer.ID
                    );

                    var oPropertyValue = new PropertyValue
                    {
                        PropertyDef = pd_EstadoProcesamiento
                    };

                    oLookup.Item = iEstadoProcesamiento;

                    oPropertyValue.TypedValue.SetValueToLookup(oLookup);

                    env.Vault.ObjectPropertyOperations.SetProperty
                    (
                        ObjVer: env.ObjVer,
                        PropertyValue: oPropertyValue
                    );
                }
            }
            catch (Exception ex)
            {
                // Cambiar el estado de procesamiento a No Procesado (Error)
                SysUtils.ReportErrorToEventLog("Ocurrio un error en el proceso de carga masiva de empleados, ", ex);
            }
            finally
            {
                // Always clean up the files (whether it works or not).
                foreach (var sFile in filesToDelete)
                {
                    File.Delete(sFile);
                }
            }            
        }

        #endregion

        #region Methods

        private void CreateOrganizacion(int iObjeto, int iClase, string sRFC, int iTipoEmpleado, string sClassName)
        {
            string sNombreOtitulo = "";
            var pd_RfcEmpresa = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");          

            // Antes de crear el proveedor o empresa interna
            if (Busqueda(iObjeto, iClase, sRFC, pd_RfcEmpresa) == false)
            {
                //if (iTipoEmpleado == 1)
                //    sNombreOtitulo = "Empresa Interna (Titulo temporal)";
                //else
                //    sNombreOtitulo = "Proveedor (Titulo temporal)";

                // Concatenar titulo o nombre de la clase
                sNombreOtitulo = sRFC + " - " + sClassName;

                // Instead of this, try "MFPropertyValuesBuilder":
                var builderCreate = new MFPropertyValuesBuilder(PermanentVault);
                builderCreate.SetClass(iClase);
                builderCreate.Add
                (
                    MFBuiltInPropertyDef.MFBuiltInPropertyDefNameOrTitle, 
                    MFDataType.MFDatatypeText,
                    sNombreOtitulo
                );
                builderCreate.Add(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRFC);

                // Created object type
                var objectTypeID = iObjeto;

                // Define the source files to add (none, in this case).
                var sourceFiles = new MFilesAPI.SourceObjectFiles();

                // Validate if the document is single-file or multi-file
                var isSingleFileDocument =
                    objectTypeID == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument &&
                    sourceFiles.Count == 1;

                // Create the new object and check it in.
                var objectVersion = PermanentVault.ObjectOperations.CreateNewObjectEx
                (
                    objectTypeID,
                    builderCreate.Values,
                    sourceFiles,
                    SFD: isSingleFileDocument,
                    CheckIn: true
                );
            }
        }

        private bool Busqueda(int iObjetoABuscar, int iClaseABuscar, string sRFCValue, int iPropertyDefRFC)
        {
            bool bExist = false;
           
            // Busqueda
            var searchBuilder = new MFSearchBuilder(PermanentVault);
            searchBuilder.Deleted(false);
            searchBuilder.ObjType(iObjetoABuscar);
            searchBuilder.Property(iPropertyDefRFC, MFDataType.MFDatatypeText, sRFCValue);
            searchBuilder.Property
            (
                (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                MFDataType.MFDatatypeLookup,
                iClaseABuscar
            );

            var searchResults = searchBuilder.Find();

            // Validate if objects were found in the search
            if (searchResults.Count > 0)
            {
                bExist = true;
            }

            return bExist;
        }

        private bool CreateEmpleadoInternoOExterno(
            string sNombreCompleto, 
            int iTipoEmpleado, 
            string sEmail, 
            string sRfcEmpresa, 
            string sCURP,
            int iObjetoEmpresa, 
            int iClaseEmpresa,
            string sClassName)
        {
            var wf_FlujoValidaciones = PermanentVault
                .WorkflowOperations
                .GetWorkflowIDByAlias("WF.FlujoValidaciones");

            var wfs_EstadoPorValidar = PermanentVault
                .WorkflowOperations
                .GetWorkflowStateIDByAlias("WFS.ValidacionDeCep.Validar");

            bool bCreado = false;
            int iObjetoEmpleado = 0;
            int iClaseEmpleado = 0;
            //int iClaseOrganizacion = 0;
            int iPropertyDef = 0;            

            var oLookupTE = new Lookup
            {
                Item = iTipoEmpleado
            };
           
            try
            {
                //string sNombreCompleto = sNombres + " " + sApellidoPaterno;

                var ot_Empleado = PermanentVault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.Employee");
                var ot_ContactoExterno = PermanentVault.ObjectTypeOperations.GetObjectTypeIDByAlias("MF.OT.ExternalContact");
                var cl_Empleado = PermanentVault.ClassOperations.GetObjectClassIDByAlias("MF.CL.Employee");
                var cl_ContactoExternoServEsp = PermanentVault.ClassOperations.GetObjectClassIDByAlias("CL.ContactoExternoServicioEspecializado");
                //var cl_Proveedor = PermanentVault.ClassOperations.GetObjectClassIDByAlias("CL.ProveedorServicioEspecializado");
                //var cl_EmpresaInterna = PermanentVault.ClassOperations.GetObjectClassIDByAlias("CL.EmpresaInterna");
                var pd_DisplayName = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.DisplayName");
                //var pd_Nombres = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.FirstName");
                //var pd_ApellidoPaterno = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.LastName.Paterno");
                //var pd_ApellidoMaterno = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.ApellidoMaterno");
                var pd_EmploymentType = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.EmploymentType");
                var pd_Email = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.EmailAddress");
                var pd_RfcEmpresa = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.RfcEmpresa");
                var pd_EmpresaInterna = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.EmpresaInterna");
                //var pd_ProveedorSE = PermanentVault.PropertyDefOperations.GetPropertyDefIDByGUID("{422B1B0F-6285-4713-94C6-DD4157325626}"); // 1730;
                var pd_Proveedor = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("MF.PD.Proveedor");
                var pd_CURP = PermanentVault.PropertyDefOperations.GetPropertyDefIDByAlias("PD.Curp");

                SysUtils.ReportInfoToEventLog("Tipo de empleado: " + iTipoEmpleado);

                // Busca y crea empresa interna / empresa externa (proveedor)
                CreateOrganizacion(iObjetoEmpresa, iClaseEmpresa, sRfcEmpresa, iTipoEmpleado, sClassName);

                if (iTipoEmpleado == 1)
                {
                    //iClaseOrganizacion = cl_EmpresaInterna;
                    iPropertyDef = pd_EmpresaInterna;
                    iObjetoEmpleado = ot_Empleado;
                    iClaseEmpleado = cl_Empleado;
                }
                else
                {
                    //iClaseOrganizacion = cl_Proveedor;
                    iPropertyDef = pd_Proveedor;
                    iObjetoEmpleado = ot_ContactoExterno;
                    iClaseEmpleado = cl_ContactoExternoServEsp;
                }                    

                // Buscar empresa para relacion con empleado
                var searchBuilderEmpresa = new MFSearchBuilder(PermanentVault);
                searchBuilderEmpresa.Deleted(false);
                searchBuilderEmpresa.Property(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcEmpresa);
                searchBuilderEmpresa.ObjType(iObjetoEmpresa);
                //searchBuilderEmpresa.Property
                //(
                //    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                //    MFDataType.MFDatatypeLookup,
                //    iClaseOrganizacion
                //);

                var searchResultsEmpresa = searchBuilderEmpresa.Find();

                // Validar si el empleado ya esta creado
                var searchBuilderEmpleado = new MFSearchBuilder(PermanentVault);
                searchBuilderEmpleado.Deleted(false);
                searchBuilderEmpleado.Property(pd_CURP, MFDataType.MFDatatypeText, sCURP);
                searchBuilderEmpleado.Property
                (
                    (int)MFBuiltInPropertyDef.MFBuiltInPropertyDefClass,
                    MFDataType.MFDatatypeLookup,
                    iClaseEmpleado
                );               

                var searchResultsEmpleado = searchBuilderEmpleado.Find();

                if (searchResultsEmpleado.Count <= 0) // Crear Employee o Contacto Externo SE
                {                    
                    var builderCrearEmpleado = new MFPropertyValuesBuilder(PermanentVault);
                    builderCrearEmpleado.SetClass(iClaseEmpleado);
                    //builderCrearEmpleado.Add(pd_Nombres, MFDataType.MFDatatypeText, sNombres);
                    //builderCrearEmpleado.Add(pd_ApellidoPaterno, MFDataType.MFDatatypeText, sApellidoPaterno);
                    //builderCrearEmpleado.Add(pd_ApellidoMaterno, MFDataType.MFDatatypeText, sApellidoMaterno);
                    builderCrearEmpleado.Add(pd_EmploymentType, MFDataType.MFDatatypeLookup, oLookupTE);
                    builderCrearEmpleado.Add(pd_Email, MFDataType.MFDatatypeText, sEmail);
                    builderCrearEmpleado.Add(pd_RfcEmpresa, MFDataType.MFDatatypeText, sRfcEmpresa);
                    builderCrearEmpleado.Add(pd_CURP, MFDataType.MFDatatypeText, sCURP);
                    builderCrearEmpleado.SetWorkflowState(wf_FlujoValidaciones, wfs_EstadoPorValidar);

                    if (iTipoEmpleado == 1)
                    {
                        builderCrearEmpleado.Add(pd_DisplayName, MFDataType.MFDatatypeText, sNombreCompleto);
                    }

                    if (searchResultsEmpresa.Count > 0)
                    {
                        var oLookups = new Lookups();
                        var oLookup = new Lookup();

                        foreach (ObjectVersion result in searchResultsEmpresa)
                        {
                            oLookup.Item = result.ObjVer.ID;
                            oLookups.Add(-1, oLookup);
                        }

                        builderCrearEmpleado.Add(iPropertyDef, MFDataType.MFDatatypeMultiSelectLookup, oLookups);

                        //if (iTipoEmpleado == 1) // Empleado Interno
                        //{
                            
                        //}
                        //else // Empleado Externo
                        //{
                        //    oLookup.Item = searchResultsEmpresa[1].ObjVer.ID;
                        //    builderCrearEmpleado.Add(iPropertyDef, MFDataType.MFDatatypeLookup, oLookup);
                        //}
                    }

                    // Created object type
                    var objectTypeID = iObjetoEmpleado;

                    // Define the source files to add (none, in this case).
                    var sourceFiles = new SourceObjectFiles();

                    // Validate if the document is single-file or multi-file
                    var isSingleFileDocument =
                        objectTypeID == (int)MFBuiltInObjectType.MFBuiltInObjectTypeDocument &&
                        sourceFiles.Count == 1;

                    // Create the new object and check it in.
                    var objectVersion = PermanentVault.ObjectOperations.CreateNewObjectEx
                    (
                        objectTypeID,
                        builderCrearEmpleado.Values,
                        sourceFiles,
                        SFD: isSingleFileDocument,
                        CheckIn: true
                    );                    
                }

                bCreado = true;
            }
            catch (Exception ex)
            {
                SysUtils.ReportErrorToEventLog("Error en creacion de empleado en la boveda, ", ex);
            }

            return bCreado;
        }

        #endregion
    }
}