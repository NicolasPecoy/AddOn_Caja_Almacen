using AddOn_Caja.Controlador;
using B1SLayer;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace AddOn_Caja.Clases
{
    public class SboClass
    {
        #region Global Definitions
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private String sPath;
        char c = Convert.ToChar(92);
        private string tipoConexionBaseDatos = "HANNA";
        private string usuarioLogueado = ""; // Usuario logueado
        public string usuarioLogueadoCode = String.Empty;
        private string monedaStrISO = "UYU"; private string monedaStrSimbolo = "$";
        SAPbouiCOM.Form oFormVisor;
        SAPbouiCOM.Form oFormLogin;
        clsConfiguracion configAddOn = new clsConfiguracion();
        clsCliente clienteSeleccionado = new clsCliente();
        String monedaSistema;
        string cambio = "";
        string monedaDocSeleccionado = "";
        bool esSuperUsuario = false;
        bool verDevoluciones = false;
        string proveedorPOS = "GEOCOM";
        bool produccion = true;
        string sucursalActiva = "";
        int decretoLey = 0;
        double dbCambio = 0;
        private string contrasena = "";

        #endregion

        #region Estructura SBO
        private void init()
        {
            try
            {
                try
                {
                    SetApplication();
                    getUsuarioLogueado();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR " + ex.Message);
                    System.Environment.Exit(1);
                }
                try
                {
                    SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(this.SBO_Application_MenuEvent);
                    SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                    SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);


                }
                catch
                { }
                try
                {
                    oCompany = SBO_Application.Company.GetDICompany();
                }
                catch (Exception ex)
                { }
                try
                {
                    // InitDeclareUdfs();
                    obtenerDatosConexion();
                    monedaSistema = ObtenerMonedaLocal(); // Obtengo las monedas locales
                    AddMenuItems();
                }
                catch (Exception ex)
                { }
            }
            catch (Exception ex)
            { }
        }

        public SboClass()
        {
            init();
            sucursalActiva = ObtenerSucActiva();
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi;
            String sConnectionString;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = Environment.GetCommandLineArgs().GetValue(1).ToString(); // 1 Estaba antes el uno pero no lo traia como param
            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication();
        }

        private int SetConnectionContext()
        {
            String sCookie;
            String sConnectionContext;
            oCompany = new SAPbobsCOM.Company();
            sCookie = oCompany.GetContextCookie();
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
            if (oCompany.Connected == true)
            {
                oCompany.Disconnect();
            }
            return oCompany.SetSboLoginContext(sConnectionContext);
        }

        public String LoadFromXML(String FileName)
        {
            System.Xml.XmlDocument oXmlDoc;
            String sPath;

            oXmlDoc = new System.Xml.XmlDocument();
            sPath = System.Windows.Forms.Application.StartupPath + c;

            oXmlDoc.Load(sPath + FileName);
            return (oXmlDoc.InnerXml);
        }

        public bool obtenerDatosConexion()
        {
            bool res = false;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select U_EMPRESA, U_FORMATO_FECHA, U_CAJAMN, U_CAJAME, U_TRANSFMN, U_TRANSFME, U_CHEQUEMN, U_CHEQUEME, U_TARJETAMN, U_TARJETAME, U_IMPRIME, U_TERMINAL,U_HASH,U_EMPTRANSACT from [@ADDONCAJADATOS]";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "Select \"U_EMPRESA\",\"U_FORMATO_FECHA\",\"U_CAJAMN\",\"U_CAJAME\",\"U_TRANSFMN\",\"U_TRANSFME\",\"U_CHEQUEMN\",\"U_CHEQUEME\",\"U_TARJETAMN\",\"U_TARJETAME\",\"U_IMPRIME\", \"U_TERMINAL\",\"U_HASH\",\"U_EMPTRANSACT\" from \"@ADDONCAJADATOS\"";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        configAddOn.GuardaLog = true;
                        configAddOn.Empresa = oRSMyTable.Fields.Item("U_EMPRESA").Value;
                        configAddOn.FormatoFecha = oRSMyTable.Fields.Item("U_FORMATO_FECHA").Value;
                        configAddOn.CajaME = oRSMyTable.Fields.Item("U_CAJAME").Value;
                        configAddOn.CajaMN = oRSMyTable.Fields.Item("U_CAJAMN").Value;
                        configAddOn.ChequeME = oRSMyTable.Fields.Item("U_CHEQUEME").Value;
                        configAddOn.ChequeMN = oRSMyTable.Fields.Item("U_CHEQUEMN").Value;
                        configAddOn.TarjetaME = oRSMyTable.Fields.Item("U_TARJETAME").Value;
                        configAddOn.TarjetaMN = oRSMyTable.Fields.Item("U_TARJETAMN").Value;
                        configAddOn.TransferenciaME = oRSMyTable.Fields.Item("U_TRANSFME").Value;
                        configAddOn.TransferenciaMN = oRSMyTable.Fields.Item("U_TRANSFMN").Value;
                        configAddOn.Imprime = Convert.ToBoolean(oRSMyTable.Fields.Item("U_IMPRIME").Value);
                        configAddOn.terminal = oRSMyTable.Fields.Item("U_TERMINAL").Value;
                        configAddOn.hash = oRSMyTable.Fields.Item("U_HASH").Value;
                        configAddOn.emprTransact = oRSMyTable.Fields.Item("U_EMPTRANSACT").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                res = true;

                return res;
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al obtenerDatosConexion", ex.Message.ToString());
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al leer los campos de la BD.");
                return res;
            }
        }
        #endregion

        #region Eventos

        //Menu Events
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == false)
                {
                    if (pVal.MenuUID.Equals("VisorCaja"))
                    {
                        //if (existeUsuarioRegistrado())
                        //{
                        CargarFormulario();
                        //}
                        //else
                        //{
                        //    CargarFormularioLogin();
                        //}

                    }
                    else if (pVal.MenuUID.Equals("vpagos"))
                    {
                        CargarFormularioPagos();
                    }
                    else if (pVal.MenuUID.Equals("VisorDev"))
                    {
                        CargarFormularioPagosDevolucion();
                    }
                    else if (pVal.MenuUID.Equals("Conf"))
                    {
                        CargarTerminales();
                    }
                    else if (pVal.MenuUID.Equals("VisorError"))
                    {
                        CargarFormularioPagosError();
                    }
                    else if (pVal.MenuUID.Equals("VisorG"))
                    {
                        cargarFormularioError();
                    }
                }
            }
            catch
            {
                BubbleEvent = false;
            }
            BubbleEvent = true;
        }

        private void SBO_Application_ItemEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            try
            {
                if (!pVal.BeforeAction)
                {
                    #region "InicializaAddOn"
                    try
                    {
                        // Entra aca cuando se da enter en la pantalla de Bloqueo
                        //if (pVal.FormTypeEx.Equals("821") && pVal.ItemUID.Equals("1") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                        //    //init();
                    }
                    catch (Exception ex)
                    { }
                    #endregion

                    if (pVal.ItemUID.Equals("3") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.ColUID.Equals("V_19") && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                            SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", pVal.Row); // DocEntry
                            int docEntry = Int32.Parse(ed.Value);
                            ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_4", pVal.Row); // DocEntry
                            if (ed.Value.ToString().Equals("Factura"))
                                SBO_Application.OpenForm(BoFormObjectEnum.fo_Invoice, "", docEntry.ToString()); // Otra opción para abrir un Form
                            else
                                SBO_Application.OpenForm(BoFormObjectEnum.fo_InvoiceCreditMemo, "", docEntry.ToString()); // Otra opción para abrir un Form
                        }
                        catch (Exception ex)
                        { }
                    }

                    #region "Crea Boton y Clic en el mismo"
                    // Creacion del botón Descargar Cotizaciones en Form 866
                    /*
                    if (pVal.FormTypeEx.ToString().Equals("133") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == false)
                    {
                        try
                        {
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("133", 0);
                            SAPbouiCOM.Button oButton;
                            SAPbouiCOM.Item oItem;
                            oForm.Items.Add("btnPagoPos", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem = oForm.Items.Item("btnPagoPos");
                            oItem.LinkTo = "2";
                            oItem.Left = oForm.Items.Item("2").Left + 68;
                            oItem.Top = oForm.Items.Item("2").Top;
                            oItem.Width = oForm.Items.Item("2").Width + 10;
                            oItem.Height = oForm.Items.Item("2").Height;
                            oButton = oForm.Items.Item("btnPagoPos").Specific;
                            oButton.Caption = "Pagar con POS";


                        }
                        catch (Exception ex)
                        {
                            //SBO_Application.MessageBox("Error al Cargar el botón: " + ex.Message);
                        }
                    }
                    */
                    // Clic en Botón BtnWsBCU
                    if (pVal.FormTypeEx.ToString().Equals("133") && pVal.ItemUID == "btnPagoPos" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.Before_Action == false)
                    {
                        try
                        {
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("133", 0);
                            SAPbouiCOM.EditText oEdit;
                            oEdit = oForm.Items.Item("33").Specific;
                            string estatus = oEdit.Value;
                            string rut = "";
                            int rutValidado = -1;
                            int digitoAvalidar = 0;

                            if (!String.IsNullOrEmpty(estatus))
                            {

                                SBO_Application.Menus.Item("VisorCaja").Activate();

                                oEdit = oForm.Items.Item("8").Specific; //Obtiene el DocEntry de la factura
                                int DocEntry = Convert.ToInt32(oEdit.Value);

                                if (configAddOn.Empresa.Equals("ARTICO"))
                                {
                                    SAPbobsCOM.Recordset oRSMyTable = null;
                                    String query = "";
                                    try
                                    {
                                        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                                        {
                                            query += "select DocEntry,LicTradNum from oinv where DocNum = '" + DocEntry + "'";
                                        }
                                        else
                                        {
                                            query += "select \"DocEntry\", \"LicTradNum\"  from \"OINV\" where \"DocNum\" = '" + DocEntry + "'";
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                    }

                                    oRSMyTable.DoQuery(query);

                                    DocEntry = Convert.ToInt32(oRSMyTable.Fields.Item("DocEntry").Value.ToString());

                                    rut = oRSMyTable.Fields.Item("DocEntry").Value.ToString();
                                    string rutAvalidar = rut.Substring(0, rut.Length - 1);
                                    digitoAvalidar = Convert.ToInt32(rut.Substring(rut.Length - 1, 1));


                                    rutValidado = validarRUT(rutAvalidar);


                                }

                                if (rutValidado == digitoAvalidar)
                                {
                                    SAPbouiCOM.ComboBox oComboConsumidor = oFormVisor.Items.Item("cmbCons").Specific;
                                    oComboConsumidor.Select("Empresa", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                    SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                                    oComboLey.Select("No aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }

                                oEdit = oForm.Items.Item("4").Specific; //Obtiene el CardCode de la factura
                                String CardCode = oEdit.Value.ToString();
                                oEdit = oForm.Items.Item("54").Specific; //Obtiene el CardName de la factura
                                String CardName = oEdit.Value.ToString();
                                oEdit = oForm.Items.Item("22").Specific; //Obtene el monto antes del impuesto
                                String[] monedaMonto = oEdit.Value.Split(' ');
                                double MontoGravado = Convert.ToDouble(monedaMonto[1]);
                                oEdit = oForm.Items.Item("27").Specific; //Obtiene el monto del impuesto
                                monedaMonto = oEdit.Value.Split(' ');
                                double impuesto = Convert.ToDouble(monedaMonto[1]);
                                oEdit = oForm.Items.Item("29").Specific; //Obtiene el monto total del documento
                                monedaMonto = oEdit.Value.Split(' ');
                                double montoTotal = Convert.ToDouble(monedaMonto[1]);
                                oEdit = oForm.Items.Item("10").Specific;
                                string fechaDoc = oEdit.Value.ToString();
                                oEdit = oForm.Items.Item("33").Specific;
                                monedaMonto = oEdit.Value.Split(' ');
                                double saldoV = Convert.ToDouble(monedaMonto[1]);
                                string moneda = monedaMonto[0];
                                oEdit = oForm.Items.Item("4").Specific; //Obtiene el CardCode de la factura




                                oEdit = oForm.Items.Item("211").Specific;
                                string cfe = oEdit.Value;

                                SAPbouiCOM.ComboBox oCombo = oFormVisor.Items.Item("cmbMone").Specific;
                                if (moneda.Equals("$") || moneda.Equals("CLP"))
                                {
                                    oCombo.Select("Pesos", BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    oCombo.Select("Dolares", BoSearchKey.psk_ByValue);
                                }

                                SAPbouiCOM.Matrix matriz = oFormVisor.Items.Item("3").Specific;
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").Rows.Clear();
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").Rows.Add(1);

                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColNumDoc", 0, DocEntry);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColFecha", 0, fechaDoc);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCliente", 0, CardName);
                                //  string comments = (string)ds.Fields.Item("Comentarios").Value;
                                //if (!String.IsNullOrEmpty(comments))
                                //  if (comments.Contains("\r"))
                                //    comments = comments.Substring(0, comments.IndexOf("\r"));

                                // oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColComentarios", 0, comments);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColVendedor", 0, CardName);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColDocEntry", 0, DocEntry);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColMonto", 0, montoTotal);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColSaldo", 0, saldoV);
                                //oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColTipo", 0, ds.Fields.Item("Tipo").Value);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColMoneda", 0, moneda);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCardCode", 0, CardCode);
                                oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCFE", 0, cfe);

                                matriz.Columns.Item("V_9").DataBind.Bind("DatosDoc", "ColComentarios");
                                matriz.Columns.Item("V_10").DataBind.Bind("DatosDoc", "ColCliente");
                                matriz.Columns.Item("V_1").DataBind.Bind("DatosDoc", "ColFecha");
                                matriz.Columns.Item("V_2").DataBind.Bind("DatosDoc", "ColNumDoc");
                                matriz.Columns.Item("V_3").DataBind.Bind("DatosDoc", "ColCFE");
                                matriz.Columns.Item("V_4").DataBind.Bind("DatosDoc", "ColTipo");
                                matriz.Columns.Item("V_19").DataBind.Bind("DatosDoc", "ColDocEntry");
                                matriz.Columns.Item("V_20").DataBind.Bind("DatosDoc", "ColSaldo");
                                matriz.Columns.Item("V_7").DataBind.Bind("DatosDoc", "ColCardCode");
                                matriz.Columns.Item("V_8").DataBind.Bind("DatosDoc", "ColMoneda");
                                matriz.Columns.Item("V_11").DataBind.Bind("DatosDoc", "ColMonto");
                                matriz.Columns.Item("V_12").DataBind.Bind("DatosDoc", "ColVendedor");

                                // Se comentan estas líneas porque se maneja desde el Event
                                SAPbouiCOM.LinkedButton oLink;
                                oLink = matriz.Columns.Item("V_19").ExtendedObject;
                                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;

                                matriz.Columns.Item("V_1").Visible = false;
                                matriz.Columns.Item("V_2").Visible = false;
                                matriz.Columns.Item("V_7").Visible = false;
                                matriz.Columns.Item("V_4").Visible = false;
                                matriz.Columns.Item("V_8").Visible = false;
                                matriz.Columns.Item("V_11").Visible = true;
                                matriz.Columns.Item("V_11").RightJustified = true;
                                matriz.Columns.Item("V_20").RightJustified = true;
                                matriz.LoadFromDataSource();
                                matriz.AutoResizeColumns();


                                oForm.Close();
                            }
                            else
                            {
                                SBO_Application.MessageBox("La factura no tiene saldo pendiente");
                            }

                        }

                        //////////////////////////////////////////////

                        catch (Exception ex)
                        { }
                    }

                    if (pVal.FormTypeEx.ToString().Equals("170") && pVal.ItemUID == "btnCanPos" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.Before_Action == false)
                    {
                        try
                        {
                            SBO_Application.Menus.Item("vpagos").Activate();
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("170", 0);
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion
                    if (pVal.FormUID.ToString().Equals("TuneUp2") && pVal.ItemUID == "btnLogin" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.Before_Action == false)
                    {
                        try
                        {
                            SAPbouiCOM.EditText ed = SBO_Application.Forms.Item(pVal.FormUID).Items.Item("txtUsu").Specific;
                            string usuario = ed.Value;
                            ed = SBO_Application.Forms.Item(pVal.FormUID).Items.Item("txtPass").Specific;
                            contrasena = ed.Value;
                            string claveEncriptada = encrypt(contrasena);
                            agregarAUDT(usuario, claveEncriptada);
                            SBO_Application.Forms.Item(pVal.FormUID).Close(); //Cierro el formulario de Login
                            CargarFormulario();
                        }
                        catch (Exception ex)
                        {

                        }
                        //SBO_Application.ActivateMenuItem("Cotizaciones"); // Abro el formulario de Descarga de BCU
                        /*SAPbouiCOM.Form fo = SBO_Application.Forms.ActiveForm;
                        SAPbouiCOM.EditText oStatic;
                        oStatic = oFormVisor.Items.Item("Nombre1").Specific; 
                        oStatic.String = "FUNCIONA"; */

                        //fo.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                    }
                    if (pVal.ItemUID.Equals("efMonto") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        //Multiplico o divido segun moneda
                        string moneda = "";
                        SAPbouiCOM.StaticText oLabel;
                        oLabel = oFormVisor.Items.Item("mndDoc").Specific; // moneda DOC
                        oLabel.Caption = monedaDocSeleccionado;
                        CultureInfo culture = new CultureInfo("en-US");
                        SAPbouiCOM.EditText oEdi = oFormVisor.Items.Item("efMonto").Specific;
                        SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                        int pRow = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                        string montTemp = oEdi.Value;
                        montTemp = montTemp.Replace(",", ".");
                        oLabel = oFormVisor.Items.Item("lblConv").Specific; // moneda DOC

                        //SAPbouiCOM.ComboBox OMoneda = oFormVisor.Items.Item("cmbMone").Specific;
                        //string comboMoneda = OMoneda.Selected.Value.ToString();

                        if (monedaDocSeleccionado.Equals("USD") || monedaDocSeleccionado.Equals("U$S"))
                        {
                            dbCambio = Math.Round((double.Parse(montTemp, culture) * double.Parse(cambio, culture)), 2);
                            oLabel.Caption = "$: " + dbCambio.ToString("N", new CultureInfo("en-US"));
                        }
                        else
                        {
                            dbCambio = Math.Round((double.Parse(montTemp, culture) / double.Parse(cambio, culture)), 2);
                            oLabel.Caption = "U$S: " + dbCambio.ToString("N", new CultureInfo("en-US"));
                        }

                        if (pRow != -1)
                        {
                            SAPbouiCOM.EditText saldo = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", pRow); // Saldo

                            if (double.Parse(montTemp, culture) > double.Parse(saldo.Value, culture))
                            {
                                SBO_Application.MessageBox("El monto ingresado no puede ser mayor del Saldo.");
                                oEdi.Value = double.Parse(saldo.Value, culture).ToString().Replace(",", ".");
                            }
                        }
                    }

                    if (pVal.ItemUID.Equals("btnBuscar") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorGeo"))
                    {
                        SAPbouiCOM.EditText oStatic;
                        oStatic = oFormVisor.Items.Item("fhcDesde").Specific;
                        string desde = oStatic.Value.ToString();
                        oStatic = oFormVisor.Items.Item("Item_4").Specific;
                        string hasta = oStatic.Value.ToString(); //fecha hasta
                        cargarGrilla(desde, hasta);

                    }

                    if (pVal.ItemUID.Equals("btnAnular") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorGeo"))
                    {
                        List<clsPago> listaDocumentos = new List<clsPago>();
                        SAPbouiCOM.Grid oGrid;
                        SAPbouiCOM.EditText ed;
                        oGrid = (SAPbouiCOM.Grid)oFormVisor.Items.Item("grillaLog").Specific;
                        if (oGrid.Rows.SelectedRows.Count > 0)
                        {

                            int selrow = oGrid.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder); //OBTIENE EL ROW SELECCIONADO
                            string docEntry = Convert.ToString(oGrid.DataTable.GetValue("DocEntry", selrow));

                            int opcion = SBO_Application.MessageBox("¿Está seguro de que quiere anular el documento?", 1, "Si", "No");
                            if (opcion == 1)
                            {
                                string DocNum = "";
                                if (oGrid.Rows.IsSelected(selrow))
                                {
                                    DevolucionPagoDocumentosErrorSAP(docEntry.ToString());
                                }
                            }
                            //DevolucionPagoPos();
                        }
                        else { SBO_Application.MessageBox("Debe seleccionar una fila!"); }
                    }

                    if (pVal.ItemUID.Equals("3") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                        int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                        int digitoAvalidar = 0;
                        string cuentaEfectivo = "";
                        if (row != -1)
                        {
                            try
                            {
                                SAPbouiCOM.EditText recibo = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_2", pVal.Row); // DocNum
                                SAPbouiCOM.EditText oEdiRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                oEdiRecibo.Value = recibo.Value.ToString();
                                //efMonto
                                SAPbouiCOM.StaticText oLabel;
                                SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", pVal.Row); // DocEntry
                                SAPbouiCOM.EditText edMon = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", pVal.Row); // Saldo

                                CultureInfo culture = new CultureInfo("en-US");
                                string temp = edMon.Value;
                                double monto = double.Parse(edMon.Value, culture);
                                //monto = monto.Substring(0,monto.Length-4);

                                //doc.Monto = getDouble(doc.Monto.ToString())

                                SAPbouiCOM.EditText oEdi = oFormVisor.Items.Item("efMonto").Specific;
                                string montTemp = monto.ToString();
                                oEdi.Value = montTemp.Replace(",", ".");
                                string rut = "";
                                string query = "";
                                int docEntry = Int32.Parse(ed.Value);

                                if (tipoConexionBaseDatos.Equals("SQL"))
                                    query = "select DocCur, LicTradNum from OINV where DocEntry = " + docEntry;
                                else
                                    query = "select \"DocCur\", \"LicTradNum\" from \"OINV\" where \"DocEntry\" = " + docEntry;

                                SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oRSMyTable.DoQuery(query);
                                monedaDocSeleccionado = oRSMyTable.Fields.Item("DocCur").Value.ToString();
                                rut = oRSMyTable.Fields.Item("LicTradNum").Value.ToString();

                                string rutAvalidar = "";
                                if (rut.Length == 12)
                                {
                                    rutAvalidar = rut.Substring(0, rut.Length - 1);
                                    digitoAvalidar = Convert.ToInt32(rut.Substring(rut.Length - 1, 1));
                                }
                                else
                                {
                                    digitoAvalidar = -1;
                                }

                                SAPbouiCOM.ComboBox oCombo = oFormVisor.Items.Item("cmbMone").Specific;
                                oLabel = oFormVisor.Items.Item("mndDoc").Specific; // moneda DOC
                                oLabel.Caption = monedaDocSeleccionado;

                                //Almacen Rural
                                //SAPbouiCOM.StaticText oLabel;
                                oLabel = oFormVisor.Items.Item("lblTC").Specific; // Tasa de cambio
                                cambio = ObtenerCambioAlmacen(monedaDocSeleccionado).Replace(",", ".");
                                oLabel.Caption = cambio;

                                //Multiplico o divido segun moneda
                                //lblConv
                                oLabel = oFormVisor.Items.Item("lblConv").Specific; // moneda DOC
                                if (monedaDocSeleccionado.Equals("USD"))
                                {
                                    dbCambio = Convert.ToDouble(montTemp) * Convert.ToDouble(cambio);
                                    oLabel.Caption = "$: " + dbCambio.ToString();
                                }
                                else
                                {
                                    dbCambio = Convert.ToDouble(montTemp) / Convert.ToDouble(cambio);
                                    oLabel.Caption = "U$S: " + dbCambio.ToString();
                                }

                                if (monedaDocSeleccionado.Equals("$") || monedaDocSeleccionado.Equals("CLP"))
                                    oCombo.Select("Pesos", BoSearchKey.psk_ByValue);
                                else
                                    oCombo.Select("Dolares", BoSearchKey.psk_ByValue);

                                string monedaTemp;
                                if (monedaDocSeleccionado.Equals("U$S"))
                                    monedaTemp = "USD";
                                else
                                    monedaTemp = "UYU";

                                //Cuentas efectivo ALMACEN RURAL
                                SAPbouiCOM.ComboBox oComboEfec = oFormVisor.Items.Item("cmbEfec").Specific;
                                int Count = oComboEfec.ValidValues.Count;

                                //Borrar datos de combobox
                                for (int i = 0; i < Count; i++)
                                {
                                    oComboEfec.ValidValues.Remove(oComboEfec.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                string q = "select \"U_Cuenta\" as Code, \"U_DescCuent\" as Name from \"@CUENTASPAGSEFECTIVO\" where \"U_Moneda\" = '" + monedaTemp + "' and \"U_Sucursal\" in " +
                                   "(SELECT T0.\"Branch\" FROM \"OUSR\" T0 , \"OUBR\" T1 WHERE T1.\"Code\" = t0.\"Branch\" and \"U_NAME\" = '" + usuarioLogueado + "')";
                                llenarCombo(oComboEfec, q, false, false, false, false);

                                SAPbouiCOM.ComboBox oComboCtaCheque = oFormVisor.Items.Item("chCta").Specific;
                                Count = oComboCtaCheque.ValidValues.Count;
                                //Borrar datos de combobox
                                for (int i = 0; i < Count; i++)
                                {
                                    oComboCtaCheque.ValidValues.Remove(oComboCtaCheque.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                                sucursalActiva = ObtenerSucActiva();

                                string queryCheques = "";
                                if (monedaDocSeleccionado == "U$S")
                                    queryCheques = "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaCheques\" = '1' and \"ActCurr\" = '" + monedaDocSeleccionado + "' and \"U_SucCheques\" ='" + sucursalActiva + "' order by Name";
                                else
                                    queryCheques = "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaCheques\" = '1' and (\"ActCurr\" = '" + monedaDocSeleccionado + "' or \"ActCurr\" = '##' ) and \"U_SucCheques\" ='" + sucursalActiva + "' order by Name";
                                llenarCombo(oComboCtaCheque, queryCheques, false, false, false, false);

                                //("ActCurr" = '$' or "ActCurr" = '##')
                                SAPbouiCOM.ComboBox oComboCtaTransf = oFormVisor.Items.Item("transfCta").Specific;
                                Count = oComboCtaTransf.ValidValues.Count;

                                //Borrar datos de combobox
                                for (int i = 0; i < Count; i++)
                                {
                                    oComboCtaTransf.ValidValues.Remove(oComboCtaTransf.ValidValues.Count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                if (monedaDocSeleccionado == "$")
                                    llenarCombo(oComboCtaTransf, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"Finanse\" = 'Y' and \"U_CtaTransf\" = '1' and (\"ActCurr\" = '" + monedaDocSeleccionado + "' or \"ActCurr\" = '##') order by Name", false, false, false, false);
                                else
                                    llenarCombo(oComboCtaTransf, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"Finanse\" = 'Y' and \"U_CtaTransf\" = '1' and \"ActCurr\" = '" + monedaDocSeleccionado + "' order by Name", false, false, false, false);

                                int rutValidado = validarRUT(rutAvalidar);

                                if (rutValidado == digitoAvalidar)
                                {
                                    SAPbouiCOM.ComboBox oComboConsumidor = oFormVisor.Items.Item("cmbCons").Specific;
                                    oComboConsumidor.Select("Empresa", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                    SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                                    oComboLey.Select("No aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    SAPbouiCOM.ComboBox oComboConsumidor = oFormVisor.Items.Item("cmbCons").Specific;
                                    oComboConsumidor.Select("Final", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                    SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                                    oComboLey.Select("Aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                            }
                            catch (Exception ex)
                            { }
                        }
                    }

                    #region "BuscarGrilla"
                    if (pVal.ItemUID.Equals("btnBuscar") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            SAPbouiCOM.Button oButton = oFormVisor.Items.Item("btnBuscar").Specific;
                            oButton.Item.Enabled = false;
                            CargarGrilla(); // Ejecuto metodo que carga la grilla
                            oButton.Item.Enabled = true;
                        }
                        catch (Exception ex)
                        { }
                    }
                    else if (pVal.ItemUID.Equals("btnBus2") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorPago"))
                    {
                        try
                        {
                            SAPbouiCOM.Button oButton = oFormVisor.Items.Item("btnBus2").Specific;
                            oButton.Item.Enabled = false;
                            CargarGrillaPagos(); // Ejecuto la funcion que carga la grilla

                            // se limpian cajas de texto de los filtros
                            SAPbouiCOM.EditText oEdit = oFormVisor.Items.Item("nTIcket").Specific;
                            oEdit.Value = "";
                            oEdit = oFormVisor.Items.Item("nFac").Specific;
                            oEdit.Value = "";

                            oButton.Item.Enabled = true;
                        }
                        catch (Exception ex)
                        { }
                    }
                    else if (pVal.ItemUID.Equals("btnBus2") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorDev"))
                    {
                        try
                        {
                            SAPbouiCOM.Button oButton = oFormVisor.Items.Item("btnBus2").Specific;
                            oButton.Item.Enabled = false;
                            CargarGrillaPagosDevoluciones(); // Ejecuto la funcion que carga la grilla

                            // se limpian cajas de texto de los filtros
                            SAPbouiCOM.EditText oEdit = oFormVisor.Items.Item("nTIcket").Specific;
                            oEdit.Value = "";
                            oEdit = oFormVisor.Items.Item("nFac").Specific;
                            oEdit.Value = "";

                            oButton.Item.Enabled = true;
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "BotonConfig"
                    if (pVal.ItemUID.Equals("btnOK") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("Conf"))
                    {
                        try
                        {
                            string empresa, formatofecha, efectivoMN, efectivoME, chequeMN, chequeME, trasfMe, transfMN, imprime, agregarTerminal, resComboTerminal, hash, empresatransact;
                            empresa = formatofecha = efectivoMN = efectivoME = chequeMN = chequeME = trasfMe = transfMN = imprime = agregarTerminal = resComboTerminal = hash = empresatransact = "";
                            int impresion = 0;
                            SAPbouiCOM.EditText oStatic;

                            oStatic = oFormVisor.Items.Item("txtTer").Specific;
                            agregarTerminal = oStatic.Value;

                            SAPbouiCOM.ComboBox OTerminal = oFormVisor.Items.Item("cmbTer").Specific;
                            if (OTerminal.Selected == null) resComboTerminal = ""; else resComboTerminal = OTerminal.Selected.Value.ToString();


                            SAPbouiCOM.ComboBox Oimprime = oFormVisor.Items.Item("cmbPDF").Specific;
                            if (Oimprime.Selected == null) impresion = 0;
                            else
                            {
                                imprime = Oimprime.Selected.Value.ToString();
                                if (imprime.Equals("SI")) impresion = 1;
                                ;
                            }

                            SAPbouiCOM.ComboBox Ofecha = oFormVisor.Items.Item("cmbFormatF").Specific;
                            if (Ofecha.Selected == null) formatofecha = ""; else formatofecha = Ofecha.Selected.Description.ToString();



                            oStatic = oFormVisor.Items.Item("txtEmp").Specific;
                            empresa = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtEmn").Specific;
                            efectivoMN = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtEme").Specific;
                            efectivoME = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtChmn").Specific;
                            chequeMN = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txttrmn").Specific;
                            transfMN = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtChme").Specific;
                            chequeME = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtTrme").Specific;
                            trasfMe = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtHash").Specific;
                            hash = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtrans").Specific;
                            empresatransact = oStatic.Value;

                            //se realiza update
                            bool res = modificarConfiguracion(empresa, formatofecha, efectivoMN, efectivoME, chequeMN, chequeME, trasfMe, transfMN, impresion, agregarTerminal, resComboTerminal, hash, empresatransact);


                            if (res)
                            {
                                SBO_Application.MessageBox("Configuración modificada correctamente");
                            }
                            else
                            {
                                SBO_Application.MessageBox("Error al realizar modificación");
                            }
                        }
                        catch (Exception ex)
                        { }
                    }

                    if (pVal.ItemUID.Equals("btninsert") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("Conf"))
                    {
                        try
                        {
                            string empresa, formatofecha, efectivoMN, efectivoME, chequeMN, chequeME, trasfMe, transfMN, imprime, agregarTerminal, resComboTerminal, hash, empresatransact;
                            empresa = formatofecha = efectivoMN = efectivoME = chequeMN = chequeME = trasfMe = transfMN = imprime = agregarTerminal = resComboTerminal = hash = empresatransact = " ";
                            int impresion = 0;
                            SAPbouiCOM.EditText oStatic;

                            oStatic = oFormVisor.Items.Item("txtTer").Specific;
                            agregarTerminal = oStatic.Value;

                            SAPbouiCOM.ComboBox OTerminal = oFormVisor.Items.Item("cmbTer").Specific;
                            if (OTerminal.Selected == null) resComboTerminal = ""; else resComboTerminal = OTerminal.Selected.Value.ToString();


                            SAPbouiCOM.ComboBox Oimprime = oFormVisor.Items.Item("cmbPDF").Specific;
                            if (Oimprime.Selected == null) impresion = 0;
                            else
                            {
                                imprime = Oimprime.Selected.Value.ToString();
                                if (imprime.Equals("SI")) impresion = 1;
                                ;
                            }

                            SAPbouiCOM.ComboBox Ofecha = oFormVisor.Items.Item("cmbFormatF").Specific;
                            if (Ofecha.Selected == null) formatofecha = ""; else formatofecha = Ofecha.Selected.Description.ToString();



                            oStatic = oFormVisor.Items.Item("txtEmp").Specific;
                            empresa = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtEmn").Specific;
                            efectivoMN = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtEme").Specific;
                            efectivoME = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtChmn").Specific;
                            chequeMN = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txttrmn").Specific;
                            transfMN = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtChme").Specific;
                            chequeME = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtTrme").Specific;
                            trasfMe = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtHash").Specific;
                            hash = oStatic.Value;

                            oStatic = oFormVisor.Items.Item("txtrans").Specific;
                            empresatransact = oStatic.Value;

                            if (!String.IsNullOrEmpty(agregarTerminal))
                            {


                                if (!String.IsNullOrEmpty(formatofecha) && !String.IsNullOrEmpty(empresa))
                                {
                                    // se realiza insert

                                    bool res = insertarConfiguracion(empresa, formatofecha, efectivoMN, efectivoME, chequeMN, chequeME, trasfMe, transfMN, impresion, agregarTerminal, hash, empresatransact);

                                    if (res)
                                    {
                                        SBO_Application.MessageBox("Configuración ingresada correctamente");
                                    }
                                    else
                                    {
                                        SBO_Application.MessageBox("Error al ingresar configuración");

                                    }
                                }
                                else
                                {
                                    SBO_Application.MessageBox("Para ingresar una nueva terminal son obligatorios los campos Empresa y Formato Fecha");

                                }

                            }
                            else
                            {
                                SBO_Application.MessageBox("Para ingresar una nueva terminal son obligatorios los campos Empresa y Formato Fecha");
                            }

                        }

                        catch (Exception ex)
                        { }
                    }

                    #endregion

                    #region "BotonesCobrar"
                    if (pVal.ItemUID.Equals("btEfectivo") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cobrar el documento seleccionado?", 1, "Aceptar", "Cancelar");
                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {
                                List<clsPago> listaDocumentos = new List<clsPago>();
                                CultureInfo culture = new CultureInfo("en-US");

                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;

                                string cardCodeAnt = "";

                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", row); // Saldo
                                        docSeleccionado.Monto = double.Parse(ed.Value, culture); // double.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_10", row); // CardName
                                        docSeleccionado.CardName = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;
                                        docSeleccionado.Fecha = DateTime.Now;
                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                // Monto en efectivo
                                SAPbouiCOM.EditText oMontoEf = oFormVisor.Items.Item("efMonto").Specific;
                                string montoCaja = oMontoEf.String;
                                double motnoPagoTemp = TextoaDecimal(montoCaja);

                                if (motnoPagoTemp <= listaDocumentos[0].Monto && motnoPagoTemp > 0)
                                {///////
                                    string numRecibo = "";
                                    if (configAddOn.Empresa.Equals("ALMACEN"))
                                    {

                                        SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                        numRecibo = oRecibo.String;

                                    }
                                    if (!String.IsNullOrEmpty(numRecibo) || !configAddOn.Empresa.Equals("ALMACEN"))
                                    {

                                        if (listaDocumentos.Count > 0) // Si tiene algún pago entonces realiza el Cobro
                                            if (ingresarPagoEfectivo(listaDocumentos) == true)
                                            {
                                                if (configAddOn.Imprime == true)
                                                {
                                                    foreach (clsPago doc in listaDocumentos)
                                                    {
                                                        //bool imprimioTEST = imprimirDocumentoCrystal(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);

                                                        bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                        for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                        { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                                    }
                                                }

                                                SBO_Application.MessageBox("El pago se ingresó correctamente!");

                                                // se limpian cajas de texto de los filtros
                                                SAPbouiCOM.EditText oEdit = oFormVisor.Items.Item("efMonto").Specific;
                                                oEdit.Value = "";
                                                SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                                oRecibo.Value = "";
                                                LimpiarDatos();


                                                CargarGrilla(); // Ejecuto la funcion que carga la grilla
                                            }
                                    }
                                    else
                                    {
                                        SBO_Application.MessageBox("El número de recibo es obliatorio!");
                                    }
                                    /////

                                }
                                else
                                {
                                    SBO_Application.MessageBox("El monto a cobrar debe ser menor al saldo de la factura");
                                }
                            }

                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago efectivo.", ex.Message.ToString()); }
                    }

                    // Se cambia combo de consumidor final
                    if (pVal.ItemUID.Equals("cmbCons") && pVal.FormUID.Equals("VisorFacturas") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {

                        SAPbouiCOM.ComboBox oConsu = oFormVisor.Items.Item("cmbCons").Specific;
                        string checkConsumidor = oConsu.Selected.Value.ToString();
                        if (checkConsumidor.Equals("Empresa"))
                        {
                            //cmbCons
                            SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                            oComboLey.Select("No aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);

                        }
                        else
                        {
                            //cmbCons
                            SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                            oComboLey.Select("Aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }


                    }

                    // Se cambia combo de cterminal en configuraciones
                    if (pVal.ItemUID.Equals("cmbTer") && pVal.FormUID.Equals("Conf") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {



                        SAPbouiCOM.ComboBox oTerm = oFormVisor.Items.Item("cmbTer").Specific;
                        string terminal = oTerm.Selected.Value.ToString();
                        CargarFormularioConfigAddOn(terminal);


                    }

                    if (pVal.ItemUID.Equals("transfCta") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        SAPbouiCOM.ComboBox oCuenta = oFormVisor.Items.Item("transfCta").Specific;
                        string selVal = oCuenta.Selected.Description.Trim();
                        if (selVal.Contains("Millas"))
                            oFormVisor.Items.Item("transfRef").Visible = true;
                        else
                            oFormVisor.Items.Item("transfRef").Visible = false;
                    }

                    if (pVal.ItemUID.Equals("btTransf") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cobrar el documento seleccionado?", 1, "Aceptar", "Cancelar");
                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {
                                List<clsPago> listaDocumentos = new List<clsPago>();
                                CultureInfo culture = new CultureInfo("en-US");

                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                                string cardCodeAnt = "";
                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", row); // Saldo
                                        docSeleccionado.Monto = double.Parse(ed.Value, culture); // double.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_10", row); // CardName
                                        docSeleccionado.CardName = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;
                                        docSeleccionado.Fecha = DateTime.Now;
                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                string numRecibo = "";
                                if (configAddOn.Empresa.Equals("ALMACEN"))
                                {
                                    SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                    numRecibo = oRecibo.String;
                                }

                                if (!String.IsNullOrEmpty(numRecibo) || !configAddOn.Empresa.Equals("ALMACEN"))
                                {
                                    if (listaDocumentos.Count > 0) // Si tiene algún pago entonces realiza el Cobro
                                        if (ingresarPagoTransferencia(listaDocumentos) == true)
                                        {
                                            if (configAddOn.Imprime == true)
                                            {
                                                foreach (clsPago doc in listaDocumentos)
                                                {
                                                    bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                    for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                    { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                                }
                                            }

                                            SBO_Application.MessageBox("El pago se ingresó correctamente!");
                                            SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                            oRecibo.Value = "";
                                            CargarGrilla(); // Ejecuto la funcion que carga la grilla
                                            LimpiarDatos();
                                        }
                                }
                                else
                                {
                                    SBO_Application.MessageBox("El número de recibo es obligatorio!");
                                }

                            }
                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago Transferencia.", ex.Message.ToString()); }
                    }

                    if (pVal.ItemUID.Equals("btCheque") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cobrar el documento seleccionado?", 1, "Aceptar", "Cancelar");
                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {
                                List<clsPago> listaDocumentos = new List<clsPago>();

                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                                string cardCodeAnt = "";
                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", row); // Saldo
                                        docSeleccionado.Monto = getDouble(ed.Value); // double.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_10", row); // CardName
                                        docSeleccionado.CardName = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;
                                        docSeleccionado.Fecha = DateTime.Now;
                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                string numRecibo = "";
                                if (configAddOn.Empresa.Equals("ALMACEN"))
                                {

                                    SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                    numRecibo = oRecibo.String;

                                }
                                if (!String.IsNullOrEmpty(numRecibo) || !configAddOn.Empresa.Equals("ALMACEN"))
                                {
                                    if (listaDocumentos.Count > 0) // Si tiene algún pago entonces realiza el Cobro
                                        if (ingresarPagoCheque(listaDocumentos) == true)
                                        {
                                            if (configAddOn.Imprime == true)
                                            {
                                                foreach (clsPago doc in listaDocumentos)
                                                {
                                                    bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                    for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                    { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                                }
                                            }

                                            SBO_Application.MessageBox("El pago se ingresó correctamente!");
                                            SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                            oRecibo.Value = "";
                                            CargarGrilla(); // Ejecuto la funcion que carga la grilla
                                            LimpiarDatos();
                                        }

                                }
                                else
                                {
                                    SBO_Application.MessageBox("El número de recibo es obligatorio!");
                                }

                            }
                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago cheque.", ex.Message.ToString()); }
                    }

                    if (pVal.ItemUID.Equals("btTarjeta") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cobrar el documento seleccionado?", 1, "Aceptar", "Cancelar");
                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {
                                List<clsPago> listaDocumentos = new List<clsPago>();

                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                                string cardCodeAnt = "";
                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", row); // Saldo
                                        docSeleccionado.Monto = getDouble(ed.Value); // double.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_10", row); // CardName
                                        docSeleccionado.CardName = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;
                                        docSeleccionado.Fecha = DateTime.Now;
                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                if (listaDocumentos.Count > 0) // Si tiene algún pago entonces realiza el Cobro
                                    if (ingresarPagoTarjeta(listaDocumentos) == true)
                                    {
                                        if (configAddOn.Imprime == true)
                                        {
                                            foreach (clsPago doc in listaDocumentos)
                                            {
                                                bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                            }
                                        }

                                        SBO_Application.MessageBox("El pago se ingresó correctamente!");
                                        CargarGrilla(); // Ejecuto la funcion que carga la grilla
                                    }
                            }
                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago tarjeta.", ex.Message.ToString()); }
                    }


                    if (pVal.ItemUID.Equals("btnPost") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cobrar el documento seleccionado?", 1, "Aceptar", "Cancelar");

                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {
                                List<clsPago> listaDocumentos = new List<clsPago>();
                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                                string cardCodeAnt = "";
                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    CultureInfo culture = new CultureInfo("en-US");
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", row); // Saldo
                                        docSeleccionado.Monto = double.Parse(ed.Value, culture); // double.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_10", row); // CardName
                                        docSeleccionado.CardName = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_11", row);
                                        docSeleccionado.TotalFactura = double.Parse(ed.Value.ToString(), culture); ;

                                        docSeleccionado.Fecha = DateTime.Now;
                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                string numRecibo = "";
                                if (configAddOn.Empresa.Equals("ALMACEN"))
                                {
                                    SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                    numRecibo = oRecibo.String;
                                }

                                if (!String.IsNullOrEmpty(numRecibo) || !configAddOn.Empresa.Equals("ALMACEN"))
                                {
                                    if (listaDocumentos.Count > 0)
                                    { // Si tiene algún pago entonces realiza el Cobro
                                        if (mandarPagoAPostAsync(listaDocumentos))
                                        {
                                            if (configAddOn.Imprime == true)
                                            {
                                                foreach (clsPago doc in listaDocumentos)
                                                {
                                                    bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                    for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                    { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                                }
                                            }

                                            SBO_Application.MessageBox("El pago se ingresó correctamente!");
                                            SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                                            oRecibo.Value = "";
                                            LimpiarDatos();
                                            CargarGrilla(); // Ejecuto la funcion que carga la grilla
                                        } //Manda los datos indicados de la factura
                                    }
                                }
                                else
                                {
                                    SBO_Application.MessageBox("El número de recibo es obliatorio!");
                                }
                            }
                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago tarjeta.", ex.Message.ToString()); }
                    }

                    #region botonCancelarPago

                    if (pVal.ItemUID.Equals("btnCanc") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorPago"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cancelar el pago del documento seleccionado?", 1, "Aceptar", "Cancelar");
                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {

                                List<clsPago> listaDocumentos = new List<clsPago>();
                                CultureInfo culture = new CultureInfo("en-US");

                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("2").Specific;
                                string cardCodeAnt = "";
                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_11", row); // Saldo
                                        docSeleccionado.Monto = double.Parse(ed.Value, culture); // double.Parse(ed.Value);

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_12", row); // CardName
                                        docSeleccionado.CardName = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_3", row);
                                        docSeleccionado.Ticket = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("colTran", row);
                                        docSeleccionado.TranId = ed.Value;

                                        try
                                        {
                                            ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("colFac", row);
                                            string facSinSerie = ed.Value;
                                            facSinSerie = facSinSerie.Substring(1, facSinSerie.Length - 1);
                                            docSeleccionado.Factura = Convert.ToInt32(facSinSerie);
                                        }
                                        catch (Exception)
                                        {

                                            docSeleccionado.Factura = 0;
                                        }


                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("TC", row); //  Tasa de cambio
                                        docSeleccionado.TasaCambio = (double.Parse(ed.Value, culture)).ToString();

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("Col_50", row);
                                        docSeleccionado.MonedaPago = ed.Value;

                                        docSeleccionado.Fecha = DateTime.Now;

                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                if (listaDocumentos.Count > 0)
                                { // Si tiene algún pago entonces realiza el Cobro
                                    if (cancelarPagoPos(listaDocumentos))
                                    {
                                        if (configAddOn.Imprime == true)
                                        {
                                            foreach (clsPago doc in listaDocumentos)
                                            {
                                                bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                            }
                                        }

                                        SBO_Application.MessageBox("El pago se canceló correctamente!");
                                        CargarGrillaPagos(); // Ejecuto la funcion que carga la grilla
                                    } //Manda los datos indicados de la factura
                                }
                                else
                                {
                                    SBO_Application.MessageBox("Error al cancelar el pago");
                                }


                            }
                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago tarjeta.", ex.Message.ToString()); }
                    }

                    if (pVal.ItemUID.Equals("btnDev") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorDev"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cancelar el pago del documento seleccionado?", 1, "Aceptar", "Cancelar");
                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {

                                List<clsPago> listaDocumentos = new List<clsPago>();
                                CultureInfo culture = new CultureInfo("en-US");

                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;
                                string cardCodeAnt = "";
                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_11", row); // Saldo
                                        docSeleccionado.Monto = double.Parse(ed.Value, culture); // double.Parse(ed.Value);

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_12", row); // CardName
                                        docSeleccionado.CardName = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_3", row);
                                        docSeleccionado.Ticket = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("colTran", row);
                                        docSeleccionado.TranId = ed.Value;

                                        try
                                        {
                                            ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("colFac", row);
                                            string facSinSerie = ed.Value;
                                            facSinSerie = facSinSerie.Substring(1, facSinSerie.Length - 1);

                                        }
                                        catch (Exception)
                                        {

                                            docSeleccionado.Factura = 0;
                                        }


                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("TC", row); //  Tasa de cambio
                                        docSeleccionado.TasaCambio = ed.Value;

                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("Col_50", row);
                                        docSeleccionado.MonedaPago = ed.Value;

                                        docSeleccionado.Fecha = DateTime.Now;

                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                if (listaDocumentos.Count > 0)
                                { // Si tiene algún pago entonces realiza el Cobro
                                    if (DevolucionPagoPos(listaDocumentos))
                                    {
                                        if (configAddOn.Imprime == true)
                                        {
                                            foreach (clsPago doc in listaDocumentos)
                                            {
                                                bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                            }
                                        }

                                        SBO_Application.MessageBox("El pago se canceló correctamente!");
                                        CargarGrillaPagosDevoluciones(); // Ejecuto la funcion que carga la grilla
                                    } //Manda los datos indicados de la factura
                                }
                                else
                                {
                                    SBO_Application.MessageBox("Error al cancelar el pago");
                                }


                            }
                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago tarjeta.", ex.Message.ToString()); }
                    }
                    #endregion
                    #endregion

                    #region "BotonPagoMultiple"
                    if (pVal.ItemUID.Equals("btPagoMult") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisorFacturas"))
                    {
                        try
                        {
                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cobrar el documento seleccionado?", 1, "Aceptar", "Cancelar");
                            if (respuestaMge == 1) // Quiere decir que dio clic en Aceptar
                            {
                                List<clsPago> listaDocumentos = new List<clsPago>();

                                SAPbouiCOM.Matrix oMatrix = oFormVisor.Items.Item("3").Specific;

                                string cardCodeAnt = "";

                                int row = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar un documento");
                                else
                                {
                                    while (row != -1)
                                    {
                                        clsPago docSeleccionado = new clsPago();
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_19", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_20", row); // Saldo
                                        docSeleccionado.Monto = getDouble(ed.Value); // double.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_7", row); // CardCode
                                        docSeleccionado.CardCode = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_10", row); // CardName
                                        docSeleccionado.CardName = ed.Value;
                                        ed = (SAPbouiCOM.EditText)oMatrix.GetCellSpecific("V_8", row); // Moneda
                                        docSeleccionado.Moneda = ed.Value;
                                        docSeleccionado.Fecha = DateTime.Now;
                                        if (String.IsNullOrEmpty(cardCodeAnt) || cardCodeAnt.Equals(docSeleccionado.CardCode)) // Si es el primer Doc o si es del mismo cliente
                                        {
                                            listaDocumentos.Add(docSeleccionado);
                                            cardCodeAnt = docSeleccionado.CardCode;
                                        }
                                        else
                                            SBO_Application.MessageBox("El documento Nro " + docSeleccionado.DocEntry + " por " + docSeleccionado.Moneda + " " + docSeleccionado.Monto + " no es del mismo Cliente");

                                        row = oMatrix.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                }

                                if (listaDocumentos.Count > 0) // Si tiene algún pago entonces realiza el Cobro
                                    if (ingresarPagoMultiple(listaDocumentos) == true)
                                    {
                                        if (configAddOn.Imprime == true)
                                        {
                                            foreach (clsPago doc in listaDocumentos)
                                            {
                                                bool imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), 0);
                                                for (int i = 1; i < 3 && imprimio == false; i++) // Si no imprimio de primera vuelve a intentarlo hasta 2 veces mas 
                                                { imprimio = imprimirDocumento(doc.DocEntry.ToString(), doc.DocNum.ToString(), i); }
                                            }
                                        }

                                        SBO_Application.MessageBox("El pago se ingresó correctamente!");
                                        CargarGrilla(); // Ejecuto la funcion que carga la grilla
                                    }
                            }
                        }
                        catch (Exception ex)
                        { guardaLogProceso("", "", "ItemEvent Pago efectivo.", ex.Message.ToString()); }
                    }
                    #endregion

                    #region "ChooseFromList"
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "chEmisor" && pVal.Before_Action == false)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                        string sCFL_ID = null;
                        sCFL_ID = oCFLEvento.ChooseFromListUID;
                        SAPbouiCOM.Form oForm = null;
                        oForm = SBO_Application.Forms.Item(FormUID);
                        SAPbouiCOM.ChooseFromList oCFL = null;
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                        if (oCFLEvento.BeforeAction == false && oCFLEvento.SelectedObjects != null)
                        {
                            SAPbouiCOM.DataTable oDataTable = null;
                            oDataTable = oCFLEvento.SelectedObjects;
                            string val = null; /*string rutCliente = "";*/

                            try
                            {
                                string cardCode = System.Convert.ToString(oDataTable.GetValue(0, 0));
                                string cardName = System.Convert.ToString(oDataTable.GetValue(1, 0)); //CardName
                                string licTradNum = System.Convert.ToString(oDataTable.GetValue(23, 0)); //LicTradNum

                                oForm.DataSources.UserDataSources.Item("CodCli").ValueEx = cardCode;

                                clienteSeleccionado.CardName = cardName;
                                clienteSeleccionado.LicTradNum = licTradNum;
                                clienteSeleccionado.CardCode = cardCode;
                                ////////val = System.Convert.ToString(oDataTable.GetValue(1, 0)); //CardName

                                ////////oForm.DataSources.UserDataSources.Item("CodCli").ValueEx = val; //  cardCode;

                                //rutCliente = System.Convert.ToString(oDataTable.GetValue(23, 0)); //LicTradNum

                                //SAPbouiCOM.EditText oStaticText;
                                //oStaticText = oFormCompras.Items.Item("5").Specific;
                                //oStaticText.DataBind.SetBound(true, "", "EditPP");
                            }
                            catch (Exception ex)
                            {
                                guardaLogProceso("et_CHOOSE_FROM_LIST", "", "Cliente - oDataTable.GetValue " + val, ex.Message);
                            }
                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            { }
            BubbleEvent = true;
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                        break;
                }
            }
            catch (Exception ex)
            { }
        }
        #endregion

        #region texto a decimal

        public double TextoaDecimal(string s)
        {
            var clone = (CultureInfo)CultureInfo.InvariantCulture.Clone();
            clone.NumberFormat.NumberDecimalSeparator = ".";
            clone.NumberFormat.NumberGroupSeparator = ",";
            // ejemplo string s = "1,14535765" o string s="1.141516";
            double d = double.Parse(s, clone);

            return d;
        }
        #endregion


        #region "Funciones"
        public bool ingresarPagoEfectivo(List<clsPago> pDocumentos)
        {
            bool res = false;
            string usuarioPago = string.Empty;
            string numRecibo = String.Empty;
            string cuentaEfectivo = String.Empty;
            CultureInfo culture = new CultureInfo("en-US");

            try
            {
                if (configAddOn.Empresa.Equals("ALMACEN"))
                {
                    SAPbouiCOM.ComboBox oComboUsu = oFormVisor.Items.Item("cmbUsr").Specific;
                    usuarioPago = oComboUsu.Selected.Value.ToString();

                    SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                    numRecibo = oRecibo.String;

                    SAPbouiCOM.ComboBox oComboCuenta = oFormVisor.Items.Item("cmbEfec").Specific;
                    try
                    {
                        cuentaEfectivo = oComboCuenta.Selected.Value.ToString();
                        if (String.IsNullOrEmpty(cuentaEfectivo))
                        {
                            SBO_Application.MessageBox("Seleccione cuenta de efectivo.");
                            return false;
                        }
                    }
                    catch (Exception)
                    {

                        SBO_Application.MessageBox("Seleccione cuenta de efectivo.");
                        return false;
                    }
                }

                if (pDocumentos.Count != 0)
                {
                    double montoTotalPago = 0;

                    SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                    int lRetCode;

                    oDoc.CardCode = pDocumentos[0].CardCode;
                    oDoc.CardName = pDocumentos[0].CardName;
                    oDoc.TransferDate = pDocumentos[0].Fecha;
                    oDoc.DocCurrency = pDocumentos[0].Moneda;
                    //oDoc.BPLID = ObtenerSucursal(pDocumentos[0].DocEntry.ToString());
                    oDoc.DocRate = Convert.ToDouble(ObtenerCambio());
                    // Monto en efectivo
                    SAPbouiCOM.EditText oMontoEf = oFormVisor.Items.Item("efMonto").Specific;
                    string montoCaja = oMontoEf.String;
                    montoTotalPago = TextoaDecimal(montoCaja);

                    foreach (clsPago doc in pDocumentos)
                    {
                        doc.Monto = double.Parse(doc.Monto.ToString(), culture); // Para corregir decimales

                        //modificacion maxi
                        //montoTotalPago += doc.Monto;

                        if (!String.IsNullOrEmpty(oMontoEf.String))
                        {
                            double numberDt;
                            bool isNumeric = double.TryParse(oMontoEf.String, out numberDt);
                            if (numberDt != 0 && montoTotalPago <= doc.Monto && Convert.ToDouble(oMontoEf.String) > 0)
                            {
                                //oDoc.CashSum = getDouble(oMontoEf.String);
                                montoTotalPago = double.Parse(oMontoEf.String, culture);
                            }
                            else
                            {
                                SBO_Application.MessageBox("El monto para Efectivo no es correcto.");
                                return false;
                            }
                        }

                        oDoc.Invoices.DocEntry = doc.DocEntry;

                        if (configAddOn.Empresa.Equals("ALMACEN"))
                        {
                            if (doc.Moneda.Equals("$"))
                            {
                                doc.Moneda = "UYU";
                            }
                            else
                            {
                                doc.Moneda = "USD";
                            }
                        }

                        if (doc.Moneda.ToString().Equals(monedaSistema))
                        {// Si el documento es en Moneda Local

                            // modificado 

                            oDoc.Invoices.SumApplied = montoTotalPago;
                        }
                        else
                        {
                            //modificado

                            // oDoc.Invoices.AppliedFC = doc.Monto;
                            oDoc.Invoices.AppliedFC = montoTotalPago;
                        }

                        if (doc.Monto >= 0)
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                        else
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                        oDoc.Invoices.Add();
                    }

                    // oDoc.CashSum = montoTotalPago;
                    oDoc.CashSum = montoTotalPago;

                    if (configAddOn.Empresa.Equals("ALMACEN"))
                    {
                        //Se cambia a una variable temporal, ya que la parametrización de Almacen es diferente al resto
                        string monedaDocTemp = "";
                        if (oDoc.DocCurrency.Equals("$")) monedaDocTemp = "UYU";
                        if (oDoc.DocCurrency.Equals("U$S")) monedaDocTemp = "USD";

                        if (ObtenerCuentaEfectivoSucursal(monedaDocTemp, cuentaEfectivo))
                        {
                            if (monedaDocTemp.Equals(monedaSistema))
                                oDoc.CashAccount = cuentaEfectivo;
                            else
                                oDoc.CashAccount = cuentaEfectivo;
                        }
                        else
                        {
                            SBO_Application.MessageBox("La moneda de la cuenta seleccionada no coincide con la moneda del documento.");
                            return false;

                        }


                        //Campos adicionales de Almacen Rural, por el momento estan harcodeados

                        // se agrega usuario logueado a documento de pago
                        oDoc.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;

                        //Numero de recibo harcodeado

                        Random rnd = new Random();
                        int randomTemp = rnd.Next(1000, 3000);

                        //oDoc.CounterReference = numRecibo; //ASPL - 2022.02.18, campo vacio.

                    }
                    else
                    {
                        if (oDoc.DocCurrency.ToString().Equals(monedaSistema))
                            oDoc.CashAccount = configAddOn.CajaMN;
                        else
                            oDoc.CashAccount = configAddOn.CajaME;

                    }

                    //    SAPbouiCOM.EditText oObservaciones = oFormVisor.Items.Item("txObser").Specific; // Se agrega campo Observaciones 28/12/18
                    //  if (!String.IsNullOrEmpty(oObservaciones.String))
                    //     oDoc.Remarks = oObservaciones.String; // Observaciones

                    if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                    {
                        lRetCode = oDoc.Add();

                        if (lRetCode != 0)
                            SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ". Codigo error:" + lRetCode.ToString());
                        else
                        {
                            res = true; // Pago ingresado correctamente 
                            clienteSeleccionado = new clsCliente();
                            //  oObservaciones.String = "";
                        }
                    }
                }
                return res;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al ingresarPagoEfectivo", ex.Message.ToString()); }
            return res;
        }

        public bool ingresarPagoTransferencia(List<clsPago> pDocumentos)
        {
            string usuarioPago = string.Empty;
            string numRecibo = String.Empty;
            CultureInfo culture = new CultureInfo("en-US");

            bool res = false;
            try
            {
                if (pDocumentos.Count != 0)
                {
                    double montoTotalPago = 0;

                    SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                    int lRetCode;

                    oDoc.CardCode = pDocumentos[0].CardCode;
                    oDoc.CardName = pDocumentos[0].CardName;
                    oDoc.TransferDate = pDocumentos[0].Fecha;
                    oDoc.DocCurrency = pDocumentos[0].Moneda;


                    //monto caja de texto efectivo
                    SAPbouiCOM.EditText oMontoEf = oFormVisor.Items.Item("efMonto").Specific;
                    string montoCaja = oMontoEf.String;
                    montoTotalPago = TextoaDecimal(montoCaja);

                    if (configAddOn.Empresa.Equals("ALMACEN"))
                    {
                        SAPbouiCOM.ComboBox oComboUsu = oFormVisor.Items.Item("cmbUsr").Specific;
                        usuarioPago = oComboUsu.Selected.Value.ToString();

                        SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                        numRecibo = oRecibo.String;
                    }

                    foreach (clsPago doc in pDocumentos)
                    {
                        if (!String.IsNullOrEmpty(oMontoEf.String))
                        {
                            double numberDt;
                            bool isNumeric = double.TryParse(oMontoEf.String, out numberDt);
                            if (numberDt != 0 && montoTotalPago <= doc.Monto && Convert.ToDouble(oMontoEf.String) > 0)
                            {
                                //oDoc.CashSum = getDouble(oMontoEf.String);
                                montoTotalPago = double.Parse(oMontoEf.String, culture);
                            }
                            else
                            {
                                SBO_Application.MessageBox("El monto para Efectivo no es correcto.");
                                return false;
                            }
                        }

                        //doc.Monto = getDouble(doc.Monto.ToString()); // Para corregir decimales
                        //montoTotalPago += doc.Monto;
                        //Modificacion 21/4/2020
                        doc.Monto = montoTotalPago;
                        oDoc.Invoices.DocEntry = doc.DocEntry;

                        if (doc.Moneda.ToString().Equals(monedaStrISO) || doc.Moneda.ToString().Equals(monedaStrSimbolo)) // Si el documento es en Moneda Local
                            oDoc.Invoices.SumApplied = doc.Monto;
                        else
                            oDoc.Invoices.AppliedFC = doc.Monto;

                        if (doc.Monto >= 0)
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                        else
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                        oDoc.Invoices.Add();
                    }

                    SAPbouiCOM.ComboBox oCuenta = oFormVisor.Items.Item("transfCta").Specific; // Cuenta seleccionada

                    if (!String.IsNullOrEmpty(oCuenta.Value.ToString()))
                        oDoc.TransferAccount = oCuenta.Selected.Value.ToString();
                    else
                    {
                        if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                            oDoc.TransferAccount = configAddOn.TransferenciaMN;
                        else
                            oDoc.TransferAccount = configAddOn.TransferenciaME;
                    }

                    //SAPbouiCOM.EditText oFecha = oFormVisor.Items.Item("transfDate").Specific; // Fecha
                    SAPbouiCOM.StaticText oLabel;
                    oLabel = oFormVisor.Items.Item("lblfecha").Specific; // moneda DOC
                    oLabel.Caption = DateTime.Now.ToString("dd/MM/yyyy");

                    SAPbouiCOM.EditText oFecha = oFormVisor.Items.Item("fcTran").Specific; // Fecha
                    if (!String.IsNullOrEmpty(oFecha.String))
                    {
                        string dueDateStr = oFecha.String.Substring(6, 4) + "-" + oFecha.String.Substring(3, 2) + "-" + oFecha.String.Substring(0, 2);
                        oDoc.TransferDate = Convert.ToDateTime(dueDateStr);
                        DateTime fecha = Convert.ToDateTime(dueDateStr);

                        if (fecha > DateTime.Now)
                        {
                            SBO_Application.MessageBox("La fecha de transferencia no puede ser mayor a hoy.");
                            return false;
                        }
                    }

                    SAPbouiCOM.EditText oReferencia = oFormVisor.Items.Item("transfRef").Specific; // Referencia
                    if (!String.IsNullOrEmpty(oReferencia.String))
                        oDoc.TransferReference = oReferencia.String;

                    SAPbouiCOM.EditText oObservaciones = oFormVisor.Items.Item("txObser").Specific; // Se agrega campo Observaciones 28/12/18
                    if (!String.IsNullOrEmpty(oObservaciones.String))
                        oDoc.Remarks = oObservaciones.String; // Observaciones

                    oDoc.TransferSum = montoTotalPago;

                    if (configAddOn.Empresa.Equals("ALMACEN"))
                    {
                        //Campos adicionales de Almacen Rural, por el momento estan harcodeados

                        // se agrega usuario logueado a documento de pago
                        oDoc.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;

                        //Numero de recibo harcodeado

                        //  Random rnd = new Random();
                        //  int randomTemp = rnd.Next(1000, 3000);

                        //oDoc.CounterReference = numRecibo; //ASPL - 2022.02.18, se cambia a campo vacio.
                    }

                    if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                    {
                        lRetCode = oDoc.Add();

                        if (lRetCode != 0)
                            SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ". Codigo error:" + lRetCode.ToString());
                        else
                        {
                            res = true; // Pago ingresado correctamente 
                            clienteSeleccionado = new clsCliente();
                            oLabel.Caption = ""; oReferencia.String = ""; oObservaciones.String = "";
                            if (!String.IsNullOrEmpty(configAddOn.TransferenciaMN))
                                oCuenta.Select(configAddOn.TransferenciaMN, BoSearchKey.psk_ByDescription); // configAddOn.TransferenciaMN;
                        }
                    }
                }
                return res;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al ingresarPagoTransferencia", ex.Message.ToString()); }
            return res;
        }

        public bool ingresarPagoTarjeta(List<clsPago> pDocumentos)
        {
            bool res = false;

            try
            {
                if (pDocumentos.Count != 0)
                {
                    double montoTotalPago = 0;

                    SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                    int lRetCode;

                    oDoc.CardCode = pDocumentos[0].CardCode;
                    oDoc.CardName = pDocumentos[0].CardName;
                    oDoc.TransferDate = pDocumentos[0].Fecha;
                    oDoc.DocCurrency = pDocumentos[0].Moneda;

                    foreach (clsPago doc in pDocumentos)
                    {
                        doc.Monto = getDouble(doc.Monto.ToString()); // Para corregir decimales

                        montoTotalPago += doc.Monto;

                        oDoc.Invoices.DocEntry = doc.DocEntry;

                        if (doc.Moneda.ToString().Equals(monedaStrISO) || doc.Moneda.ToString().Equals(monedaStrSimbolo)) // Si el documento es en Moneda Local
                            oDoc.Invoices.SumApplied = doc.Monto;
                        else
                            oDoc.Invoices.AppliedFC = doc.Monto;

                        if (doc.Monto >= 0)
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                        else
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                        oDoc.Invoices.Add();
                    }

                    SAPbouiCOM.ComboBox oCuenta = oFormVisor.Items.Item("tjaCta").Specific; // Cuenta seleccionada
                    if (!String.IsNullOrEmpty(oCuenta.Value.ToString()))
                        oDoc.CreditCards.CreditAcct = oCuenta.Selected.Value.ToString();
                    else
                    {
                        if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                            oDoc.CreditCards.CreditAcct = configAddOn.TarjetaMN;
                        else
                            oDoc.CreditCards.CreditAcct = configAddOn.TarjetaME;
                    }

                    SAPbouiCOM.ComboBox oMedioPago = oFormVisor.Items.Item("tjaDesc").Specific; // Tarjeta seleccionada
                    if (!String.IsNullOrEmpty(oMedioPago.Value.ToString()))
                    {
                        oDoc.CreditCards.CreditCard = Convert.ToInt32(oMedioPago.Selected.Value.ToString());
                        string medPago = oMedioPago.Value.ToString();

                        // corto string para reutilizar metodo ObtenerCodCC

                        string[] palabrasSeparadas = medPago.Split(' ');
                        string[] datoscc = ObtenerCodCC(palabrasSeparadas[0], palabrasSeparadas[1], oDoc.DocCurrency);

                        string cuentaPrueba = datoscc[1]; // Cuenta seleccionada

                    }

                    SAPbouiCOM.EditText oNumeroTja = oFormVisor.Items.Item("tjaNumero").Specific; // Numero de tarjeta
                    if (!String.IsNullOrEmpty(oNumeroTja.String))
                    {
                        oDoc.CreditCards.CreditCardNumber = oNumeroTja.String;
                        if (oDoc.CreditCards.CreditCardNumber.Length > 4)
                            oDoc.CreditCards.CreditCardNumber = oDoc.CreditCards.CreditCardNumber.Substring(0, 4);
                    }

                    SAPbouiCOM.EditText oFecha = oFormVisor.Items.Item("tjaFecha").Specific; // Fecha
                    if (!String.IsNullOrEmpty(oFecha.String))
                    {
                        string dueDateStr = oFecha.String.Substring(6, 4) + "-" + oFecha.String.Substring(3, 2) + "-" + oFecha.String.Substring(0, 2);
                        oDoc.CreditCards.CardValidUntil = Convert.ToDateTime(dueDateStr);
                    }

                    SAPbouiCOM.EditText oCantCuotas = oFormVisor.Items.Item("tjaCantCuo").Specific; // Cant de cuotas
                    if (!String.IsNullOrEmpty(oCantCuotas.String))
                        oDoc.CreditCards.NumOfPayments = Convert.ToInt32(oCantCuotas.String);

                    SAPbouiCOM.EditText oVoucherNro = oFormVisor.Items.Item("tjaNroCert").Specific; // Nro Certificado
                    if (!String.IsNullOrEmpty(oVoucherNro.String))
                        oDoc.CreditCards.VoucherNum = oVoucherNro.String;

                    SAPbouiCOM.EditText oOwnerId = oFormVisor.Items.Item("tjaID").Specific; // Id Tja
                    if (!String.IsNullOrEmpty(oOwnerId.String))
                        oDoc.CreditCards.OwnerIdNum = oOwnerId.String;

                    SAPbouiCOM.EditText oObservaciones = oFormVisor.Items.Item("txObser").Specific; // Se agrega campo Observaciones 28/12/18
                    if (!String.IsNullOrEmpty(oObservaciones.String))
                        oDoc.Remarks = oObservaciones.String; // Observaciones

                    oDoc.CreditCards.CreditSum = montoTotalPago;
                    oDoc.CreditCards.NumOfCreditPayments = oDoc.CreditCards.NumOfPayments;
                    oDoc.CreditCards.PaymentMethodCode = oDoc.CreditCards.NumOfPayments; // oDoc.CreditCards.CreditCard;

                    if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                    {
                        lRetCode = oDoc.Add();

                        if (lRetCode != 0)
                            SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ". Codigo error:" + lRetCode.ToString());
                        else
                        {
                            res = true; // Pago ingresado correctamente
                            clienteSeleccionado = new clsCliente();
                            oFecha.String = ""; oCantCuotas.String = ""; oVoucherNro.String = ""; oOwnerId.String = ""; oNumeroTja.String = ""; oObservaciones.String = "";
                            oCuenta.Select("", BoSearchKey.psk_ByDescription);
                            oMedioPago.Select("", BoSearchKey.psk_ByDescription);
                            /*if (!String.IsNullOrEmpty(configAddOn.TarjetaMN))
                                oCuenta.Select(configAddOn.TarjetaMN, BoSearchKey.psk_ByDescription); // configAddOn.TransferenciaMN;*/
                        }
                    }
                }
                return res;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al ingresar", ex.Message.ToString()); }
            return res;
        }

        public bool ingresarPagoCheque(List<clsPago> pDocumentos)
        {
            bool res = false;
            string usuarioPago = string.Empty;
            string numRecibo = String.Empty;
            CultureInfo culture = new CultureInfo("en-US");

            try
            {
                if (pDocumentos.Count != 0)
                {
                    double montoTotalPago = 0;

                    SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                    int lRetCode;

                    oDoc.CardCode = pDocumentos[0].CardCode;
                    oDoc.CardName = pDocumentos[0].CardName;
                    oDoc.DocType = BoRcptTypes.rCustomer;
                    oDoc.DocDate = pDocumentos[0].Fecha; // Se agrego 08/08/18
                    oDoc.DueDate = pDocumentos[0].Fecha; // Se agrego 08/08/18
                    oDoc.DocCurrency = pDocumentos[0].Moneda;


                    //monto caja de texto efectivo
                    SAPbouiCOM.EditText oMontoEf = oFormVisor.Items.Item("efMonto").Specific;
                    string montoCaja = oMontoEf.String;
                    montoTotalPago = TextoaDecimal(montoCaja);

                    if (configAddOn.Empresa.Equals("ALMACEN"))
                    {
                        SAPbouiCOM.ComboBox oComboUsu = oFormVisor.Items.Item("cmbUsr").Specific;
                        usuarioPago = oComboUsu.Selected.Value.ToString();

                        SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                        numRecibo = oRecibo.String;
                    }

                    foreach (clsPago doc in pDocumentos)
                    {
                        if (!String.IsNullOrEmpty(oMontoEf.String))
                        {
                            double numberDt;
                            bool isNumeric = double.TryParse(oMontoEf.String, out numberDt);
                            if (numberDt != 0 && montoTotalPago <= doc.Monto && Convert.ToDouble(oMontoEf.String) > 0)
                            {
                                //oDoc.CashSum = getDouble(oMontoEf.String);
                                montoTotalPago = double.Parse(oMontoEf.String, culture);
                            }
                            else
                            {
                                SBO_Application.MessageBox("El monto para Efectivo no es correcto.");
                                return false;
                            }
                        }

                        //doc.Monto = getDouble(doc.Monto.ToString()); // Para corregir decimales
                        //montoTotalPago += doc.Monto;
                        doc.Monto = montoTotalPago;

                        oDoc.Invoices.DocEntry = doc.DocEntry;

                        if (doc.Moneda.ToString().Equals(monedaStrISO) || doc.Moneda.ToString().Equals(monedaStrSimbolo)) // Si el documento es en Moneda Local
                            oDoc.Invoices.SumApplied = doc.Monto;
                        else
                            oDoc.Invoices.AppliedFC = doc.Monto;

                        if (doc.Monto >= 0)
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                        else
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;


                        if (configAddOn.Empresa.Equals("ALMACEN"))
                        {
                            //Campos adicionales de Almacen Rural, por el momento estan harcodeados

                            // se agrega usuario logueado a documento de pago
                            oDoc.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;

                            //Numero de recibo harcodeado

                            /* Random rnd = new Random();
                             int randomTemp = rnd.Next(1000, 3000);*/

                            //oDoc.CounterReference = numRecibo; //ASPL - 2022.02.18, Campo vacio.
                        }


                        oDoc.Invoices.Add();
                    }

                    SAPbouiCOM.ComboBox oCuenta = oFormVisor.Items.Item("chCta").Specific; // Cuenta seleccionada
                    if (!String.IsNullOrEmpty(oCuenta.Value.ToString()))
                        oDoc.Checks.CheckAccount = oCuenta.Selected.Value.ToString();
                    else
                    {
                        SBO_Application.MessageBox("Debe seleccionar cuenta.");
                        return false;
                        /*if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                            oDoc.Checks.CheckAccount = configAddOn.ChequeMN;
                        else
                            oDoc.Checks.CheckAccount = configAddOn.ChequeME;*/
                    }

                    SAPbouiCOM.ComboBox oBanco = oFormVisor.Items.Item("chBanco").Specific; // Banco seleccionado
                    if (!String.IsNullOrEmpty(oBanco.Value.ToString()))
                        oDoc.Checks.BankCode = oBanco.Selected.Value.ToString();
                    else
                    {
                        SBO_Application.MessageBox("Debe seleccionar banco.");
                        return false;
                    }

                    SAPbouiCOM.EditText oNumero = oFormVisor.Items.Item("chNumero").Specific; // Numero de cheque
                    if (!String.IsNullOrEmpty(oNumero.String))
                    {
                        int number;
                        bool isNumeric = int.TryParse(oNumero.String, out number);
                        if (number != 0)
                            oDoc.Checks.CheckNumber = Convert.ToInt32(oNumero.String);
                        else
                        {
                            SBO_Application.MessageBox("Debe indicar el número del Cheque");
                            return false;
                        }
                    }

                    oDoc.CheckAccount = oDoc.Checks.CheckAccount;

                    SAPbouiCOM.EditText oFecha = oFormVisor.Items.Item("chVto").Specific; // Fecha
                    if (!String.IsNullOrEmpty(oFecha.String))
                    {
                        try
                        {
                            guardaLogProceso("", "", "Datos Pre Save", "Value chVto: " + oFecha.Value + " String chVto: " + oFecha.String);
                            //oDoc.Checks.DueDate = Convert.ToDateTime(oFecha.String);
                            string dueDateStr = oFecha.String.Substring(6, 4) + "-" + oFecha.String.Substring(3, 2) + "-" + oFecha.String.Substring(0, 2);
                            DateTime fecha = Convert.ToDateTime(dueDateStr);
                            if (fecha < DateTime.Now.Date)
                            {
                                SBO_Application.MessageBox("La fecha de vencimiento no puedo ser menor a hoy.");
                                return false;
                            }
                            oDoc.Checks.DueDate = Convert.ToDateTime(dueDateStr);
                            guardaLogProceso("", "", "Datos Pre Save", "DueDate: " + oDoc.Checks.DueDate + " dueDateStr: " + dueDateStr.ToString());

                            dueDateStr = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();
                            oDoc.DueDate = Convert.ToDateTime(dueDateStr);
                            oDoc.DocDate = Convert.ToDateTime(dueDateStr);
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso("", "", "ERROR al ingresarPagoCheque con DueDates", ex.Message.ToString());
                        }
                    }

                    SAPbouiCOM.EditText oEmisor = oFormVisor.Items.Item("chEmisor").Specific; // Emisor
                    if (!String.IsNullOrEmpty(oEmisor.String))
                    {
                        //clsCliente clienteCheque = ObtenerCliente(oEmisor.String, "C");
                        if (clienteSeleccionado != null)
                        {
                            oDoc.Checks.OriginallyIssuedBy = clienteSeleccionado.CardName;
                            oDoc.Checks.FiscalID = clienteSeleccionado.LicTradNum;
                        }
                    }

                    SAPbouiCOM.EditText oObservaciones = oFormVisor.Items.Item("txObser").Specific; // Se agrega campo Observaciones 28/12/18
                    if (!String.IsNullOrEmpty(oObservaciones.String))
                        oDoc.Remarks = oObservaciones.String; // Observaciones

                    oDoc.Checks.CheckSum = montoTotalPago;
                    oDoc.Checks.CountryCode = "UY";

                    guardaLogProceso("", "", "Datos Pre Save", "Cta: " + oDoc.Checks.CheckAccount + ". CardCode: " + oDoc.CardCode + ". CardName: " + oDoc.CardName + ". Currency: " + oDoc.DocCurrency); // + ". TransferDate: " + oDoc.TransferDate.ToString()
                    guardaLogProceso("", "", "Datos Pre Save", "Bank: " + oDoc.Checks.BankCode + ". CheckNumber: " + oDoc.Checks.CheckNumber + ". FiscalID: " + oDoc.Checks.FiscalID + ". OriginallyIssuedBy: " + oDoc.Checks.OriginallyIssuedBy + ". CheckSum: " + oDoc.Checks.CheckSum + ". DueDate: " + oDoc.DueDate.ToString());

                    if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                    {
                        lRetCode = oDoc.Add();

                        if (lRetCode != 0)
                            SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ". Codigo error:" + lRetCode.ToString());
                        else
                        {
                            res = true; // Pago ingresado correctamente
                            clienteSeleccionado = new clsCliente();
                            oFecha.String = ""; oNumero.String = ""; oEmisor.String = ""; oObservaciones.String = "";
                            if (!String.IsNullOrEmpty(configAddOn.ChequeMN))
                                oCuenta.Select(configAddOn.ChequeMN, BoSearchKey.psk_ByDescription); // configAddOn.TransferenciaMN;
                        }
                    }
                }
                return res;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al ingresarPagoCheque", ex.Message.ToString()); }
            return res;
        }

        public bool ingresarPagoMultiple(List<clsPago> pDocumentos)
        {
            bool res = false;
            try
            {

                if (pDocumentos.Count != 0)
                {
                    double montoTotalPago = 0;

                    SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                    int lRetCode;

                    oDoc.CardCode = pDocumentos[0].CardCode;
                    oDoc.CardName = pDocumentos[0].CardName;
                    oDoc.TransferDate = pDocumentos[0].Fecha;
                    oDoc.DocCurrency = pDocumentos[0].Moneda;

                    foreach (clsPago doc in pDocumentos)
                    {
                        doc.Monto = getDouble(doc.Monto.ToString()); // Para corregir decimales

                        montoTotalPago += doc.Monto;

                        oDoc.Invoices.DocEntry = doc.DocEntry;

                        if (doc.Moneda.ToString().Equals(monedaStrISO) || doc.Moneda.ToString().Equals(monedaStrSimbolo)) // Si el documento es en Moneda Local
                            oDoc.Invoices.SumApplied = doc.Monto;
                        else
                            oDoc.Invoices.AppliedFC = doc.Monto;

                        if (doc.Monto >= 0)
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                        else
                            oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                        oDoc.Invoices.Add();
                    }

                    SAPbouiCOM.EditText oObservaciones = oFormVisor.Items.Item("txObser").Specific; // Se agrega campo Observaciones 28/12/18
                    if (!String.IsNullOrEmpty(oObservaciones.String))
                        oDoc.Remarks = oObservaciones.String; // Observaciones

                    double montoTotalPagoMedios = 0; // Suma de monto con los medios de pagos

                    // Monto en efectivo
                    SAPbouiCOM.EditText oMontoEf = oFormVisor.Items.Item("efMonto").Specific;


                    if (!String.IsNullOrEmpty(oMontoEf.String))
                    {
                        double numberDt;
                        bool isNumeric = double.TryParse(oMontoEf.String, out numberDt);
                        if (numberDt != 0)
                        {
                            oDoc.CashSum = getDouble(oMontoEf.String); montoTotalPagoMedios += oDoc.CashSum;

                            if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                oDoc.CashAccount = configAddOn.CajaMN;
                            else
                                oDoc.CashAccount = configAddOn.CajaME;
                        }
                        else
                            SBO_Application.MessageBox("El monto para Efectivo no es correcto.");
                    }

                    SAPbouiCOM.ComboBox oTjCuenta = oFormVisor.Items.Item("tjaCta").Specific; // Cuenta seleccionada
                    SAPbouiCOM.ComboBox oTjMedioPago = oFormVisor.Items.Item("tjaDesc").Specific; // Tarjeta seleccionada
                    SAPbouiCOM.EditText oTjNumeroTja = oFormVisor.Items.Item("tjaNumero").Specific; // Numero de tarjeta
                    SAPbouiCOM.EditText oTjFecha = oFormVisor.Items.Item("tjaFecha").Specific; // Fecha
                    SAPbouiCOM.EditText oTjCantCuotas = oFormVisor.Items.Item("tjaCantCuo").Specific; // Cant de cuotas
                    SAPbouiCOM.EditText oTjVoucherNro = oFormVisor.Items.Item("tjaNroCert").Specific; // Nro Certificado
                    SAPbouiCOM.EditText oTjOwnerId = oFormVisor.Items.Item("tjaID").Specific; // Id Tja
                    #region "Tarjeta"
                    // Monto en tarjeta
                    SAPbouiCOM.EditText oMontoTja = oFormVisor.Items.Item("tjMonto").Specific;
                    if (!String.IsNullOrEmpty(oMontoTja.String))
                    {
                        try
                        {
                            double numberDt;
                            bool isNumeric = double.TryParse(oMontoTja.String, out numberDt);
                            if (numberDt != 0)
                            {
                                oDoc.CreditCards.CreditSum = getDouble(oMontoTja.String); montoTotalPagoMedios += oDoc.CreditCards.CreditSum;

                                if (!String.IsNullOrEmpty(oTjCuenta.Value.ToString()))
                                    oDoc.CreditCards.CreditAcct = oTjCuenta.Selected.Value.ToString();
                                else
                                {
                                    if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                        oDoc.CreditCards.CreditAcct = configAddOn.TarjetaMN;
                                    else
                                        oDoc.CreditCards.CreditAcct = configAddOn.TarjetaME;
                                }

                                if (!String.IsNullOrEmpty(oTjMedioPago.Value.ToString()))
                                    oDoc.CreditCards.CreditCard = Convert.ToInt32(oTjMedioPago.Selected.Value.ToString());

                                if (!String.IsNullOrEmpty(oTjNumeroTja.String))
                                {
                                    oDoc.CreditCards.CreditCardNumber = oTjNumeroTja.String;
                                    if (oDoc.CreditCards.CreditCardNumber.Length > 4)
                                        oDoc.CreditCards.CreditCardNumber = oDoc.CreditCards.CreditCardNumber.Substring(0, 4);
                                }

                                if (!String.IsNullOrEmpty(oTjFecha.String))
                                {
                                    string dueDateStr = oTjFecha.String.Substring(6, 4) + "-" + oTjFecha.String.Substring(3, 2) + "-" + oTjFecha.String.Substring(0, 2);
                                    oDoc.CreditCards.CardValidUntil = Convert.ToDateTime(dueDateStr);
                                }


                                if (!String.IsNullOrEmpty(oTjCantCuotas.String))
                                    oDoc.CreditCards.NumOfPayments = Convert.ToInt32(oTjCantCuotas.String);

                                if (!String.IsNullOrEmpty(oTjVoucherNro.String))
                                    oDoc.CreditCards.VoucherNum = oTjVoucherNro.String;

                                if (!String.IsNullOrEmpty(oTjOwnerId.String))
                                    oDoc.CreditCards.OwnerIdNum = oTjOwnerId.String;

                                oDoc.CreditCards.NumOfCreditPayments = oDoc.CreditCards.NumOfPayments;
                                oDoc.CreditCards.PaymentMethodCode = oDoc.CreditCards.NumOfPayments; // oDoc.CreditCards.CreditCard;
                            }
                            else
                                SBO_Application.MessageBox("El monto para Tarjeta no es correcto.");
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    SAPbouiCOM.ComboBox oChCuenta = oFormVisor.Items.Item("chCta").Specific; // Cuenta seleccionada
                    SAPbouiCOM.ComboBox oChBanco = oFormVisor.Items.Item("chBanco").Specific; // Banco seleccionado
                    SAPbouiCOM.EditText oChNumero = oFormVisor.Items.Item("chNumero").Specific; // Numero de cheque
                    SAPbouiCOM.EditText oChFecha = oFormVisor.Items.Item("chVto").Specific; // Fecha
                    SAPbouiCOM.EditText oChEmisor = oFormVisor.Items.Item("chEmisor").Specific; // Emisor
                    #region "Cheque"
                    SAPbouiCOM.EditText oMontoCh = oFormVisor.Items.Item("chMonto").Specific; // Monto en cheque
                    if (!String.IsNullOrEmpty(oMontoCh.String))
                    {
                        try
                        {
                            double numberDt;
                            bool isNumeric = double.TryParse(oMontoCh.String, out numberDt);
                            if (numberDt != 0)
                            {
                                oDoc.Checks.CheckSum = getDouble(oMontoCh.String); montoTotalPagoMedios += oDoc.Checks.CheckSum;

                                if (!String.IsNullOrEmpty(oChCuenta.Value.ToString()))
                                    oDoc.Checks.CheckAccount = oChCuenta.Selected.Value.ToString();
                                else
                                {
                                    if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                        oDoc.Checks.CheckAccount = configAddOn.ChequeMN;
                                    else
                                        oDoc.Checks.CheckAccount = configAddOn.ChequeME;
                                }

                                if (!String.IsNullOrEmpty(oChBanco.Value.ToString()))
                                    oDoc.Checks.BankCode = oChBanco.Selected.Value.ToString();

                                if (!String.IsNullOrEmpty(oChNumero.String))
                                {
                                    int number;
                                    bool isNumericN = int.TryParse(oChNumero.String, out number);
                                    if (number != 0)
                                        oDoc.Checks.CheckNumber = Convert.ToInt32(oChNumero.String);
                                    else
                                    {
                                        SBO_Application.MessageBox("Debe indicar el número del Cheque");
                                        return false;
                                    }
                                }

                                oDoc.CheckAccount = oDoc.Checks.CheckAccount;

                                if (!String.IsNullOrEmpty(oChFecha.String))
                                {
                                    try
                                    {
                                        string dueDateStr = oChFecha.String.Substring(6, 4) + "-" + oChFecha.String.Substring(3, 2) + "-" + oChFecha.String.Substring(0, 2);
                                        oDoc.Checks.DueDate = Convert.ToDateTime(dueDateStr);

                                        dueDateStr = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();
                                        oDoc.DueDate = Convert.ToDateTime(dueDateStr);
                                        oDoc.DocDate = Convert.ToDateTime(dueDateStr);
                                    }
                                    catch (Exception ex)
                                    {
                                        guardaLogProceso("", "", "ERROR al ingresarPagoCheque con DueDates", ex.Message.ToString());
                                    }
                                }

                                if (!String.IsNullOrEmpty(oChEmisor.String))
                                {
                                    //clsCliente clienteCheque = ObtenerCliente(oEmisor.String, "C");
                                    if (clienteSeleccionado != null)
                                    {
                                        oDoc.Checks.OriginallyIssuedBy = clienteSeleccionado.CardName;
                                        oDoc.Checks.FiscalID = clienteSeleccionado.LicTradNum;
                                    }
                                }

                                oDoc.Checks.CountryCode = "UY";
                            }
                            else
                                SBO_Application.MessageBox("El monto para Cheque no es correcto.");
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    SAPbouiCOM.ComboBox oTrCuenta = oFormVisor.Items.Item("transfCta").Specific; // Cuenta seleccionada
                    SAPbouiCOM.EditText oTrFecha = oFormVisor.Items.Item("transfDate").Specific; // Fecha
                    SAPbouiCOM.EditText oTrReferencia = oFormVisor.Items.Item("transfRef").Specific; // Referencia
                    #region "Transferencia"
                    // Monto en transferencia
                    SAPbouiCOM.EditText oMontoTr = oFormVisor.Items.Item("trMonto").Specific;
                    if (!String.IsNullOrEmpty(oMontoTr.String))
                    {
                        try
                        {
                            double numberDt;
                            bool isNumeric = double.TryParse(oMontoTr.String, out numberDt);
                            if (numberDt != 0)
                            {
                                oDoc.TransferSum = getDouble(oMontoTr.String); montoTotalPagoMedios += oDoc.TransferSum;

                                if (!String.IsNullOrEmpty(oTrCuenta.Value.ToString()))
                                    oDoc.TransferAccount = oTrCuenta.Selected.Value.ToString();
                                else
                                {
                                    if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                        oDoc.TransferAccount = configAddOn.TransferenciaMN;
                                    else
                                        oDoc.TransferAccount = configAddOn.TransferenciaME;
                                }

                                if (!String.IsNullOrEmpty(oTrFecha.String))
                                {
                                    string dueDateStr = oTrFecha.String.Substring(6, 4) + "-" + oTrFecha.String.Substring(3, 2) + "-" + oTrFecha.String.Substring(0, 2);
                                    oDoc.TransferDate = Convert.ToDateTime(dueDateStr);
                                }

                                if (!String.IsNullOrEmpty(oTrReferencia.String))
                                    oDoc.TransferReference = oTrReferencia.String;
                            }
                            else
                                SBO_Application.MessageBox("El monto para Transferencia no es correcto.");
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    if (montoTotalPago == montoTotalPagoMedios)
                    {
                        if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                        {
                            lRetCode = oDoc.Add();

                            if (lRetCode != 0)
                                SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ". Codigo error:" + lRetCode.ToString());
                            else
                            {
                                res = true; // Pago ingresado correctamente 
                                clienteSeleccionado = new clsCliente();
                                oObservaciones.String = "";
                                oMontoEf.String = ""; oMontoTja.String = ""; oMontoCh.String = ""; oMontoTr.String = "";

                                // Limpio los campos de Tarjeta
                                oTjFecha.String = ""; oTjCantCuotas.String = ""; oTjVoucherNro.String = ""; oTjOwnerId.String = ""; oTjNumeroTja.String = "";

                                // Limpio los campos de Cheque
                                oChFecha.String = ""; oChNumero.String = ""; oChEmisor.String = "";

                                // Limpio los campos de Transferencia
                                oTrFecha.String = ""; oTrReferencia.String = "";

                                oTjCuenta.Select("", BoSearchKey.psk_ByDescription);
                                oTjMedioPago.Select("", BoSearchKey.psk_ByDescription);
                                if (!String.IsNullOrEmpty(configAddOn.ChequeMN))
                                    oChCuenta.Select(configAddOn.ChequeMN, BoSearchKey.psk_ByDescription); // configAddOn.TransferenciaMN;
                                if (!String.IsNullOrEmpty(configAddOn.TransferenciaMN))
                                    oTrCuenta.Select(configAddOn.TransferenciaMN, BoSearchKey.psk_ByDescription); // configAddOn.TransferenciaMN;
                            }
                        }
                    }
                    else
                        SBO_Application.MessageBox("El monto del documento seleccionado no coincide con el monto de los medios de pagos");
                }
                return res;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al ingresarPagoTransferencia", ex.Message.ToString()); }
            return res;
        }

        public bool imprimirDocumento(string pDocEntry, string pDocNum, int pIntento) // Imprimir documento por Crystal Reports
        {
            bool res = false;
            try
            {
                Thread.Sleep(500);

                SAPbouiCOM.Form fo = SBO_Application.OpenForm(BoFormObjectEnum.fo_Invoice, "", pDocEntry.ToString()); // Otra opción para abrir un Form

                SBO_Application.Menus.Item("520").Activate(); // Open printing dialog
                fo.Items.Item("1").Click(); // Cierro el formulario de la factura

                res = true;
                return res;
            }
            catch (Exception ex)
            {
                if (pIntento != 0)
                    guardaLogProceso(pDocEntry.ToString(), pDocNum.ToString(), "ERROR al imprimirDocumento. Usuario: " + usuarioLogueado.ToString() + ". Intento: " + pIntento, ex.Message.ToString()); // Guarda log del Proceso
            }
            return res;
        }


        public bool imprimirDocumentoPago(string pDocEntry, string pDocNum, int pIntento) // Imprimir documento por Crystal Reports
        {
            bool res = false;
            try
            {
                Thread.Sleep(500);

                SAPbouiCOM.Form fo = SBO_Application.OpenForm(BoFormObjectEnum.fo_ContractTemplete, "24", "95"); // Otra opción para abrir un Form

                SBO_Application.Menus.Item("520").Activate(); // Open printing dialog
                fo.Items.Item("1").Click(); // Cierro el formulario de la factura

                res = true;
                return res;
            }
            catch (Exception ex)
            {
                if (pIntento != 0)
                    guardaLogProceso(pDocEntry.ToString(), pDocNum.ToString(), "ERROR al imprimirDocumento. Usuario: " + usuarioLogueado.ToString() + ". Intento: " + pIntento, ex.Message.ToString()); // Guarda log del Proceso
            }
            return res;
        }


        public string ObtenerMonedaLocal()
        {
            string res = "";
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "Select MainCurncy from OADM";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"MainCurncy\" from OADM";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = oRSMyTable.Fields.Item("MainCurncy").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                if (res.ToString().Equals("$") || res.ToString().Equals("UYU"))
                {
                    res = "UYU";
                    monedaStrISO = "UYU"; monedaStrSimbolo = "$";
                }
                else if (res.ToString().Equals("U$S") || res.ToString().Equals("USD"))
                {
                    res = "USD";
                    monedaStrISO = "USD"; monedaStrSimbolo = "U$S";
                }
                else if (res.ToString().Equals("CLP") || res.ToString().Equals("PCH"))
                {
                    monedaStrISO = "CLP"; monedaStrSimbolo = "CLP";
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("ERROR AL OBTENER MONEDA LOCAL" + ex.Message);
                return res;
            }
        }

        #region "DecimalDouble"

        public decimal getDecimal(string pNumero)
        {
            decimal res = 0;
            try
            {
                var ci = CultureInfo.InvariantCulture.Clone() as CultureInfo;
                ci.NumberFormat.NumberDecimalSeparator = System.Globalization.CultureInfo.InvariantCulture.NumberFormat.CurrencyDecimalSeparator;
                //res = decimal.Parse(pNumero, ci);

                string separadorDec = obtenerSeparadorDecimal();
                string separadorMil = obtenerSeparadorMiles();

                string cotizacionStr = String.Format("{0:0" + separadorMil + "0" + separadorDec + "######}", pNumero);

                if (cotizacionStr.Contains(separadorDec))
                    res = decimal.Parse(cotizacionStr);
                else
                    res = decimal.Parse(pNumero, ci);

                return decimal.Round(res, 2);
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR getDecimal", ex.Message.ToString());
                return res;
            }
        }

        public double getDouble(string pNumero)
        {
            double res = 0;
            try
            {
                var ci = CultureInfo.InvariantCulture.Clone() as CultureInfo;
                ci.NumberFormat.NumberDecimalSeparator = System.Globalization.CultureInfo.InvariantCulture.NumberFormat.CurrencyDecimalSeparator;
                //res = decimal.Parse(pNumero, ci);

                string separadorDec = obtenerSeparadorDecimal();
                string separadorMil = obtenerSeparadorMiles();

                string cotizacionStr = String.Format("{0:0" + separadorMil + "0" + separadorDec + "######}", pNumero);

                if (cotizacionStr.Contains(separadorDec))
                    res = double.Parse(cotizacionStr);
                else
                    res = double.Parse(pNumero, ci);

                res = Math.Round(res, 2);
                return res;
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR getDouble", ex.Message.ToString());
                return res;
            }
        }

        public string obtenerSeparadorDecimal()
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            string res = "";
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "SELECT DecSep FROM OADM";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"DecSep\" from OADM";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = (string)oRSMyTable.Fields.Item("DecSep").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
        }  // Esta funcion devuelve el Separador Decimal configurado en SAP

        public string obtenerSeparadorMiles()
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            string res = "";
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "SELECT ThousSep FROM OADM";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"ThousSep\" from OADM";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = (string)oRSMyTable.Fields.Item("ThousSep").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Obtener separador de miles", ex.Message.ToString());
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
        } // Esta funcion devuelve el Separador Miles configurado en SAP
        #endregion

        public clsCliente ObtenerCliente(string pCardCode, string pCardType)
        {
            clsCliente cliente = null;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "select CardName, LicTradNum from OCRD where CardCode = '" + pCardCode + "' and CardType = '" + pCardType + "'";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"CardName\", \"LicTradNum\" from OCRD where \"CardCode\" = \'" + pCardCode + "\' and \"CardType\" = \'" + pCardType + "\'";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        cliente.CardName = oRSMyTable.Fields.Item("CardName").Value;
                        cliente.LicTradNum = oRSMyTable.Fields.Item("LicTradNum").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return cliente;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("ERROR AL OBTENER LOS DATOS DEL CLIENTE" + ex.Message);
                return null;
            }
        }

        public bool guardaLogProceso(string pFormFactura, string pCodigoFactura, string pAccion, string pXML)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;

            try
            {
                if (configAddOn.GuardaLog == true)
                {
                    long docEntry = obtenerDocEntryLogProceso();
                    DateTime fechaHoy = DateTime.Now;
                    oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = "INSERT INTO [@ADDONLOGS] (Code, Name, U_PANTALLA, U_CODIGO,U_ACCION,U_LOGXML, U_FECHA, U_CREATE_DATE) VALUES (" + docEntry + ",'" + docEntry + "','" + pFormFactura + "','" + pCodigoFactura + "','" + pAccion + "','" + pXML.ToString() + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "')";

                    if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                        query = "INSERT INTO \"@ADDONLOGS\" (\"Code\", \"Name\", \"U_PANTALLA\", \"U_CODIGO\",\"U_ACCION\",\"U_LOGXML\", \"U_FECHA\", \"U_CREATE_DATE\") VALUES (" + docEntry + ",'" + docEntry + "','" + pFormFactura + "','" + pCodigoFactura + "','" + pAccion + "','" + pXML.ToString().Replace("'", "") + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "')";

                    oRSMyTable.DoQuery(query);
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }
        }

        private void AddChooseFromList(string pObjType, string pUniqueID, string pAlias, string pCondVal)
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                SAPbouiCOM.Conditions oCons = null;
                SAPbouiCOM.Condition oCon = null;
                oCFLs = oFormVisor.ChooseFromLists;
                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                // Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = pObjType;
                oCFLCreationParams.UniqueID = pUniqueID;
                oCFL = oCFLs.Add(oCFLCreationParams);

                if (!String.IsNullOrEmpty(pAlias) && !String.IsNullOrEmpty(pCondVal)) // Si tiene una condición para aplicar
                {
                    // Adding Conditions to CFL1
                    oCons = oCFL.GetConditions();
                    oCon = oCons.Add();
                    oCon.Alias = pAlias;
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = pCondVal;
                    oCFL.SetConditions(oCons);
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + " " + ex.ToString());
            }
        }

        public long obtenerDocEntryLogProceso() //Obtengo el último DocEntry de la tabla LOGPROCESO
        {
            long res = 1;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = "select case when MAX(CAST(Code AS bigint)) is null then 1 else MAX(CAST(Code AS bigint)) + 1 end as Prox from [@ADDONLOGS]";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select case when MAX(CAST(\"Code\" AS bigint)) is null then 1 else MAX(CAST(\"Code\" AS bigint)) + 1 end as Prox from \"@ADDONLOGS\"";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = Convert.ToInt64(oRSMyTable.Fields.Item("Prox").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo DocEntry de LogProceso.");
                return res;
            }
        }

        public bool getUsuarioLogueado()
        {
            bool res = false;
            try
            {
                SAPbouiCOM.StaticText oStatic;
                SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("169", 0);
                oStatic = (SAPbouiCOM.StaticText)oForm.Items.Item("8").Specific;
                usuarioLogueado = oStatic.Caption;
                res = true;

                return res;
            }
            catch (Exception ex)
            {
                guardaLogProceso("169", "", "ERROR al buscar usuario logueado", ex.Message.ToString());// Guarda log del Proceso
            }
            return res;
        } // Obtiene el 


        public bool getSuperUser()
        {
            bool res = false;
            SAPbobsCOM.Recordset oRSMyTable = null;

            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "SUPERUSER from OUSR where U_NAME = '" + usuarioLogueado + "'";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"SUPERUSER\" from \"OUSR\" where \"U_NAME\" = \'" + usuarioLogueado + "\'";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {

                        string esSuperU = oRSMyTable.Fields.Item("SUPERUSER").Value;
                        if (esSuperU.ToString().Equals("Y"))
                            esSuperUsuario = true;
                        else
                            esSuperUsuario = false;
                        oRSMyTable.MoveNext();
                    }
                }

                //if (idSucursalUsuario < 0) // Si la sucursal es menor a 0 entonces se usa la 1 por defecto
                //    idSucursalUsuario = 1;

                getVerDevoluciones();
                res = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }

            return res;
        }


        public bool getUsuarioLogueadoCod()
        {
            bool res = false;
            SAPbobsCOM.Recordset oRSMyTable = null;

            try
            {

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "SUPERUSER from OUSR where U_NAME = '" + usuarioLogueado + "'";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"USER_CODE\" from \"OUSR\" where \"U_NAME\" = \'" + usuarioLogueado + "\'";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        usuarioLogueadoCode = oRSMyTable.Fields.Item("USER_CODE").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                res = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }

            return res;
        }


        public bool getVerDevoluciones()
        {
            bool res = false;
            SAPbobsCOM.Recordset oRSMyTable = null;

            try
            {

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"U_Usuario\" from \"@AUTORIZACIONCAJA\" where \"U_Nombre\" = \'" + usuarioLogueado + "\'";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    string usuarioAutorizado = oRSMyTable.Fields.Item("U_Usuario").Value;

                    if (!String.IsNullOrEmpty(usuarioAutorizado))
                    {
                        verDevoluciones = true;
                        res = true;
                        return verDevoluciones;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }

            return res;
        }
        #endregion

        #region "Form"
        private void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus;
            SAPbouiCOM.MenuItem oMenuItem;
            oMenus = SBO_Application.Menus;

            //configuracion Addon
            getSuperUser();
            getUsuarioLogueadoCod();

            if (!oMenus.Exists("CAJA"))
            {
                SAPbouiCOM.MenuCreationParams oCreationPackage;
                oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuItem = SBO_Application.Menus.Item("43520");

                sPath = System.Windows.Forms.Application.StartupPath;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "CAJA";
                oCreationPackage.String = "Caja";
                oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;
                oCreationPackage.Image = sPath + "\\submenu1.BMP";

                oMenus = oMenuItem.SubMenus;

                try
                {
                    //usuario logueado sucursal activa
                    sucursalActiva = ObtenerSucActiva();

                    oMenus.AddEx(oCreationPackage);
                    oMenuItem = SBO_Application.Menus.Item("CAJA");


                    oMenus = oMenuItem.SubMenus;
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "VisorCaja";
                    oCreationPackage.String = "Visualizar Facturas";
                    oMenus.AddEx(oCreationPackage);

                    if (verDevoluciones)
                    {
                        //prueba
                        oMenus = oMenuItem.SubMenus;
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "vpagos";
                        oCreationPackage.String = "Visualizar Pagos Tarjetas";
                        oMenus.AddEx(oCreationPackage);

                        if (usuarioLogueadoCode.Equals("manager"))
                        {
                            //Devolución Pagos Geocom
                            oMenus = oMenuItem.SubMenus;
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "VisorDev";
                            oCreationPackage.String = "Devolución Lote Cerrado";
                            oMenus.AddEx(oCreationPackage);


                            oMenus = oMenuItem.SubMenus;
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "VisorG";
                            oCreationPackage.String = "Transacciones Error SAP";
                            oMenus.AddEx(oCreationPackage);
                        }
                    }


                    /*
                    //Visor pagos error
                    oMenus = oMenuItem.SubMenus;
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "VisorError";
                    oCreationPackage.String = "Visualizar Pagos Error";
                    oMenus.AddEx(oCreationPackage);

                    */



                    // if (esSuperUsuario)
                    //  {

                    /* if (usuarioLogueado.Equals("manager"))
                     {
                         oMenus = oMenuItem.SubMenus;
                         oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                         oCreationPackage.UniqueID = "Conf";
                         oCreationPackage.String = "Configuración de addOn";
                         oMenus.AddEx(oCreationPackage);
                     }*/

                    //  }

                }
                catch (Exception er)
                {
                    String msg = "";
                    if (er.Message.Equals("Menu - Already exists"))
                        msg = "Menú ya fue creado.";
                    else
                        msg = er.Message;
                }
            }
        }

        private void CargarFormulario()
        {

            try
            {

                oFormVisor = SBO_Application.Forms.Item("VisorFacturas");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisorFacturas";
                fcp.UniqueID = "VisorFacturas";
                try
                {
                    fcp.XmlData = LoadFromXML("VisorFacturas.srf");
                    oFormVisor = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {

                sPath = System.Windows.Forms.Application.StartupPath;
                string imagen = sPath + "\\almacen.jpg";

                //Logo
                SAPbouiCOM.PictureBox oImagen = oFormVisor.Items.Item("Item_6").Specific;
                oImagen.Picture = imagen;

                //Calendar
                oImagen = oFormVisor.Items.Item("calendar").Specific;
                //imagen = sPath + "\\calendar.jpg";
                imagen = sPath + "\\bmp.bmp";
                oImagen.Picture = imagen;

                //Cash
                oImagen = oFormVisor.Items.Item("imgCash").Specific;
                imagen = sPath + "\\cash.jpg";
                oImagen.Picture = imagen;

                //Credito
                oImagen = oFormVisor.Items.Item("imgcc").Specific;
                imagen = sPath + "\\credito.jpg";
                oImagen.Picture = imagen;

                //Invenzis
                //imgInv
                oImagen = oFormVisor.Items.Item("imgInv").Specific;
                imagen = sPath + "\\Invenzis_logo.jpg";
                oImagen.Picture = imagen;

                SAPbouiCOM.EditText oStatic;
                SAPbouiCOM.StaticText oLabel;
                //oFormVisor.DataSources.UserDataSources.Add("FecDes", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date3", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date4", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date5", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date6", SAPbouiCOM.BoDataType.dt_DATE, 10);

                try
                {
                    sucursalActiva = ObtenerSucActiva();
                    //Almacen Rural
                    DateTime fechaDesde = DateTime.Now;
                    DateTime fechaDesdeTemp = new DateTime(fechaDesde.Year, fechaDesde.Month, 1);
                    fechaDesde = Convert.ToDateTime(fechaDesdeTemp.ToString("dd/MM/yyyy"));

                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    oStatic.DataBind.SetBound(true, "", "FecDes");
                    oStatic.String = fechaDesde.ToString("dd/MM/yyyy");

                    // oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    oStatic.DataBind.SetBound(true, "", "Date2");

                    // oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                    // oStatic = oFormVisor.Items.Item("transfDate").Specific; // Transferencia Fecha
                    oLabel = oFormVisor.Items.Item("lblfecha").Specific; // moneda DOC
                    oLabel.Caption = DateTime.Now.ToString("dd/MM/yyyy");

                    oStatic.DataBind.SetBound(true, "", "Date3");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");

                    oStatic = oFormVisor.Items.Item("tjaFecha").Specific; // Vencimiento Tarjeta
                    oStatic.DataBind.SetBound(true, "", "Date4");

                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");

                    oStatic = oFormVisor.Items.Item("chVto").Specific; // Vencimiento Cheque
                    oStatic.DataBind.SetBound(true, "", "Date5");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");

                    oStatic = oFormVisor.Items.Item("fcTran").Specific; // Vencimiento Cheque
                    oStatic.DataBind.SetBound(true, "", "Date6");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                }
                catch (Exception ex)
                { }

                //SAPbouiCOM.StaticText oLabel;
                oLabel = oFormVisor.Items.Item("lblTC").Specific; // Tasa de cambio
                cambio = ObtenerCambio();
                oLabel.Caption = cambio;

                //oFormVisor.DataSources.UserDataSources.Add("Num1", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 2);
                //oFormVisor.DataSources.UserDataSources.Add("Num2", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 2);
                oStatic = oFormVisor.Items.Item("tjaNumero").Specific; // Tja Número
                oStatic.DataBind.SetBound(true, "", "Num1");
                oStatic = oFormVisor.Items.Item("tjaCantCuo").Specific; // Tja Cant cuotas
                oStatic.DataBind.SetBound(true, "", "Num2");

                AddChooseFromList("2", "CFL1", "CardType", "C");
                //oFormVisor.DataSources.UserDataSources.Add("CodCli", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oStatic = oFormVisor.Items.Item("chEmisor").Specific;
                oStatic.DataBind.SetBound(true, "", "CodCli");
                oStatic.ChooseFromListUID = "CFL1";
                oStatic.ChooseFromListAlias = "CardName"; // CardCode

                SAPbouiCOM.ComboBox oComboTipoFac = oFormVisor.Items.Item("TFac").Specific;
                oComboTipoFac.Select("Contado", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboCuotas = oFormVisor.Items.Item("cmbCT").Specific;
                oComboCuotas.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                oComboLey.ExpandType = BoExpandType.et_DescriptionOnly;
                oComboLey.ValidValues.Add("No aplicar devolución", "No aplicar devolución");
                oComboLey.ValidValues.Add("Aplicar devolución", "Aplicar devolución");
                oComboLey.Select("Aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboMoneda = oFormVisor.Items.Item("cmbMone").Specific;
                oComboMoneda.Select("Pesos", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboConsumidor = oFormVisor.Items.Item("cmbCons").Specific;
                oComboConsumidor.Select("Final", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboOperacion = oFormVisor.Items.Item("cmbOperac").Specific;
                oComboOperacion.Select("Venta", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboTerminal = oFormVisor.Items.Item("cmbTerm").Specific;

                SAPbouiCOM.ComboBox oComboSuc = oFormVisor.Items.Item("cmbSuc").Specific;
                oComboSuc.Select(sucursalActiva, SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboUsu = oFormVisor.Items.Item("cmbUsr").Specific;
                oComboUsu.Select(usuarioLogueado, SAPbouiCOM.BoSearchKey.psk_ByValue);

                //varias sucursales llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM [@ADDONCAJA] where U_SUCURSAL = '" + sucursalActiva + "'", false, false, false);

                try
                {
                    if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    {
                        llenarCombo(oComboTerminal, "SELECT \"U_TERMINAL\"  FROM \"@ADDONCAJADATOS\"  where \"U_SUCURSALCOD\" = '" + sucursalActiva + "' AND \"U_CODUSUARIO\" = '" + usuarioLogueadoCode + "'", false, false, false, true);
                    }
                    else
                    {
                        //solo Teyma
                        if (!usuarioLogueado.Equals("manager"))
                        {
                            llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM \"@ADDONCAJADATOS\" where \"U_SUCURSALCOD\" = '" + sucursalActiva + "'", false, false, false, true);
                        }
                        else
                        {
                            llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM \"@ADDONCAJADATOS\" ", false, false, false, true);
                        }

                    }
                }
                catch (Exception)
                {

                }

                oComboTerminal.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                CargarGrilla();

                // SAPbouiCOM.ComboBox oComboTipo = oFormVisor.Items.Item("cmbTarjeta").Specific;
                //oComboTipo.Select("Credito", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oFormVisor.Items.Item("transfRef").Visible = false; // Referencia

                SAPbouiCOM.ComboBox oComboCtaTransf = oFormVisor.Items.Item("transfCta").Specific;
                llenarCombo(oComboCtaTransf, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"Finanse\" = 'Y' and \"U_CtaTransf\" = '1' order by Name ", false, false, false, false);
                SAPbouiCOM.ComboBox oComboCtaTarjeta = oFormVisor.Items.Item("tjaCta").Specific;
                llenarCombo(oComboCtaTarjeta, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" order by Name ", true, false, false, false);
                SAPbouiCOM.ComboBox oComboDescTarjeta = oFormVisor.Items.Item("tjaDesc").Specific;
                llenarCombo(oComboDescTarjeta, "select  \"CreditCard\" as Code, \"CardName\" as Name from \"OCRC\" order by Name ", true, false, false, false);
                SAPbouiCOM.ComboBox oComboCtaCheque = oFormVisor.Items.Item("chCta").Specific;
                llenarCombo(oComboCtaCheque, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaCheques\" = '1' order by Name ", false, false, false, false);
                SAPbouiCOM.ComboBox oComboBancoCheque = oFormVisor.Items.Item("chBanco").Specific;
                llenarCombo(oComboBancoCheque, "select \"BankCode\" as Code, \"BankName\" as Name from \"ODSC\" order by Name ", false, false, false, false);
                SAPbouiCOM.ComboBox oComboUsr = oFormVisor.Items.Item("cmbUsr").Specific;
                llenarCombo(oComboUsr, "select \"U_NAME\" as Code , \"U_NAME\" as Name from \"OUSR\" WHERE \"U_NAME\" <> ''", false, false, false, false);
                oComboUsr.Select(usuarioLogueado, SAPbouiCOM.BoSearchKey.psk_ByValue);

                //Cuentas efectivo
                SAPbouiCOM.ComboBox oComboEfec = oFormVisor.Items.Item("cmbEfec").Specific;
                string q = "select \"U_Cuenta\" as Code, \"U_DescCuent\" as Name from \"@CUENTASPAGSEFECTIVO\" where \"U_Sucursal\" in " +
                   "(SELECT T0.\"Branch\" FROM \"OUSR\" T0 , \"OUBR\" T1 WHERE T1.\"Code\" = t0.\"Branch\" and \"U_NAME\" = '" + usuarioLogueado + "')";
                llenarCombo(oComboEfec, q, false, false, false, false);
                oComboEfec.Item.DisplayDesc = true;
                //  oComboUsr.Select(usuarioLogueado, SAPbouiCOM.BoSearchKey.psk_ByValue);
                //Seleccion por defecto

                SAPbouiCOM.ComboBox oComboTransf = oFormVisor.Items.Item("transfCta").Specific;
                oComboTransf.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboProveedor = oFormVisor.Items.Item("cmbProv").Specific;
                oComboProveedor.ExpandType = BoExpandType.et_DescriptionOnly;
                oComboProveedor.Item.DisplayDesc = true;
                oComboProveedor.ValidValues.Add("0", "");
                oComboProveedor.ValidValues.Add("ALMACEN RURAL", "ALMACEN RURAL");
                oComboProveedor.ValidValues.Add("BIAH", "BIAH");
                oComboProveedor.ValidValues.Add("EPICENTRO", "EPICENTRO");
                oComboProveedor.ValidValues.Add("OROFINO", "OROFINO");
                oComboProveedor.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                SAPbouiCOM.ComboBox oComboSello = oFormVisor.Items.Item("cmbSell").Specific;
                oComboSello.ExpandType = BoExpandType.et_DescriptionOnly;
                oComboSello.Item.DisplayDesc = true;
                oComboSello.ValidValues.Add("0", "");
                oComboSello.ValidValues.Add("2", "AMEX CRÉDITO");
                oComboSello.ValidValues.Add("5", "CABAL CRÉDITO");
                oComboSello.ValidValues.Add("55", "CREDITEL CREDITO");
                oComboSello.ValidValues.Add("52", "MASTERCARD CRÉDITO");
                oComboSello.ValidValues.Add("15", "MAESTRO DÉBITO");
                oComboSello.ValidValues.Add("21", "OCA CRÉDITO");
                oComboSello.ValidValues.Add("24", "VISA CRÉDITO");
                oComboSello.ValidValues.Add("34", "VISA DÉBITO");
                oComboSello.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                SAPbouiCOM.ComboBox oComboEfec2 = oFormVisor.Items.Item("cmbEfec").Specific;
                oComboEfec2.ExpandType = BoExpandType.et_DescriptionOnly;
                oComboEfec2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboBank = oFormVisor.Items.Item("chBanco").Specific;
                oComboBank.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboCta = oFormVisor.Items.Item("chCta").Specific;
                oComboCta.ExpandType = BoExpandType.et_DescriptionOnly;
                oComboCta.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                try
                {
                    oComboCtaTransf.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboCtaTarjeta.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboDescTarjeta.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboCtaCheque.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboBancoCheque.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboUsr.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                }
                catch (Exception ex)
                { }

                oFormVisor.Visible = true;
                if (!String.IsNullOrEmpty(configAddOn.TransferenciaMN))
                    oComboCtaTransf.Select(configAddOn.TransferenciaMN, BoSearchKey.psk_ByDescription);
                if (!String.IsNullOrEmpty(configAddOn.ChequeMN))
                    oComboCtaCheque.Select(configAddOn.ChequeMN, BoSearchKey.psk_ByDescription);

            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarFormulario", ex.Message.ToString()); }
        }

        private void CargarGrilla()
        {
            SAPbouiCOM.Matrix matriz = null;
            string tipoFac = "Contado";
            string Sucursal = "0";
            try
            {
                if (oFormVisor != null)
                {
                    matriz = oFormVisor.Items.Item("3").Specific;
                }
                else
                {
                    oFormVisor = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormVisor.Items.Item("3").Specific;
                }

                SAPbouiCOM.EditText oStatic;
                DateTime fechaDesde = Convert.ToDateTime(DateTime.Now);
                DateTime fechaHasta = Convert.ToDateTime(DateTime.Now);
                try
                {

                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaDesde.ToString("dd/MM/yyyy");

                    fechaDesde = Convert.ToDateTime(oStatic.String);

                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaHasta.ToString("dd/MM/yyyy");

                    fechaHasta = Convert.ToDateTime(oStatic.String);

                    SAPbouiCOM.ComboBox oComboTipoFac = oFormVisor.Items.Item("TFac").Specific;
                    tipoFac = oComboTipoFac.Selected.Value.ToString();

                    SAPbouiCOM.ComboBox oComboSuc = oFormVisor.Items.Item("cmbSuc").Specific;
                    Sucursal = oComboSuc.Selected.Value.ToString();



                }
                catch (Exception ex)
                { }

                SAPbobsCOM.Recordset ds = obtenerFacturasPendientes(fechaDesde, fechaHasta, tipoFac, Sucursal);

                SAPbouiCOM.ComboBox oComboCuotas = oFormVisor.Items.Item("cmbCT").Specific;
                oComboCuotas.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oFormVisor.DataSources.DataTables.Item("DatosDoc").Rows.Clear();
                oFormVisor.DataSources.DataTables.Item("DatosDoc").Rows.Add(ds.RecordCount);
                int cont = 0;

                while (!ds.EoF)
                {
                    try
                    {
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColNumDoc", cont, ds.Fields.Item("DocNum").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColFecha", cont, ds.Fields.Item("DocDate").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCardCode", cont, ds.Fields.Item("CardCode").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCliente", cont, ds.Fields.Item("CardName").Value);
                        string comments = (string)ds.Fields.Item("Comentarios").Value;
                        if (!String.IsNullOrEmpty(comments))
                            if (comments.Contains("\r"))
                                comments = comments.Substring(0, comments.IndexOf("\r"));

                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColComentarios", cont, comments);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColVendedor", cont, ds.Fields.Item("Vendedor").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColDocEntry", cont, ds.Fields.Item("DocEntry").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColMonto", cont, ds.Fields.Item("DocTotal").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColSaldo", cont, ds.Fields.Item("Saldo").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColTipo", cont, ds.Fields.Item("Tipo").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColMoneda", cont, ds.Fields.Item("Moneda").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCFE", cont, ds.Fields.Item("CFE").Value);
                        cont++;
                    }
                    catch (Exception ex)
                    { guardaLogProceso("", "", "ERROR al CargarGrilla 02", ex.Message.ToString()); }
                    ds.MoveNext();
                }

                matriz.Columns.Item("V_9").DataBind.Bind("DatosDoc", "ColComentarios");
                matriz.Columns.Item("V_7").DataBind.Bind("DatosDoc", "ColCardCode");
                matriz.Columns.Item("V_10").DataBind.Bind("DatosDoc", "ColCliente");
                matriz.Columns.Item("V_1").DataBind.Bind("DatosDoc", "ColFecha");
                matriz.Columns.Item("V_2").DataBind.Bind("DatosDoc", "ColNumDoc");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosDoc", "ColCFE");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosDoc", "ColTipo");
                matriz.Columns.Item("V_19").DataBind.Bind("DatosDoc", "ColDocEntry");
                matriz.Columns.Item("V_20").DataBind.Bind("DatosDoc", "ColSaldo");
                matriz.Columns.Item("V_8").DataBind.Bind("DatosDoc", "ColMoneda");
                matriz.Columns.Item("V_11").DataBind.Bind("DatosDoc", "ColMonto");
                matriz.Columns.Item("V_12").DataBind.Bind("DatosDoc", "ColVendedor");

                // Se comentan estas líneas porque se maneja desde el Event
                //SAPbouiCOM.LinkedButton oLink;
                //oLink = matriz.Columns.Item("V_19").ExtendedObject;
                //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;

                matriz.Columns.Item("V_1").Visible = false;
                matriz.Columns.Item("V_2").Visible = true;
                matriz.Columns.Item("V_7").Visible = true;
                matriz.Columns.Item("V_4").Visible = false;
                matriz.Columns.Item("V_8").Visible = true;
                matriz.Columns.Item("V_11").Visible = true;
                matriz.Columns.Item("V_11").RightJustified = true;
                matriz.Columns.Item("V_20").RightJustified = true;
                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarGrilla 03", ex.Message.ToString()); }
        }

        public SAPbobsCOM.Recordset obtenerFacturasPendientes(DateTime pDesde, DateTime pHasta, string tipoFac, string sucursal)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;

            //Solo Almacén Rural
            string condicion = "";
            string CondicionSuc = "";
            if (tipoFac.Equals("Contado"))
            {
                condicion = " and (T1.\"GroupNum\" = -1 or T1.\"GroupNum\" = 52 ) ";
            }
            else
            {
                condicion = " and T1.\"GroupNum\" <> -1";
            }
            if (sucursal.Equals("0"))
            {
                CondicionSuc = "";
            }
            else
            {
                CondicionSuc = " and temp.\"SUCURSAL\" = " + sucursal;
            }

            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = "select T1.DocEntry,T1.DocNum, T1.DocDate, T1.CardCode, T1.CardName, CONCAT(FolioPref,FolioNum) as CFE,  " +
                "case when T1.DocCur = '" + monedaStrSimbolo.ToString() + "' or T1.DocCur = '" + monedaStrISO.ToString() + "' then T1.DocTotal else T1.DocTotalFC end as DocTotal, " +
                "case when T1.DocCur = '" + monedaStrSimbolo.ToString() + "' or T1.DocCur = '" + monedaStrISO.ToString() + "' then (T1.DocTotal - T1.PaidToDate) else (T1.DocTotalFC - T1.PaidFC) end as Saldo, " +
                "case when T1.DocSubType <> 'DN' then 'Factura' else 'Nota Debito' end as Tipo, T1.DocCur as Moneda, T1.Comments as Comentarios, T3.SlpName as Vendedor from OINV as T1 " +
                "inner join OCTG as T2 ON T2.GroupNum = T1.GroupNum " +
                "left join OSLP as T3 ON T3.SlpCode = T1.SlpCode " +
               // "where T1.CANCELED = 'N' and UPPER(T2.PymntGroup) LIKE '%CONTADO%' AND T1.DocStatus = 'O' ";
               "where T1.CANCELED = 'N'  AND T1.DocStatus = 'O'";
                //"and t1.BPLId = '" + sucursalActiva + "'";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                {
                    //Almacen Rural
                    query = "select T1.\"CardCode\",T1.\"DocEntry\",T1.\"DocNum\", T1.\"DocDate\", T1.\"CardCode\", T1.\"CardName\",CONCAT (\"FolioPref\", CONCAT (\'-\' , \"FolioNum\")) as CFE, " +
                    "case when T1.\"DocCur\" = \'" + monedaStrSimbolo.ToString() + "\' or T1.\"DocCur\" = \'" + monedaStrISO.ToString() + "\' then T1.\"DocTotal\" else T1.\"DocTotalFC\" end as DocTotal, " +
                    "case when T1.\"DocCur\" = \'" + monedaStrSimbolo.ToString() + "\' or T1.\"DocCur\" = \'" + monedaStrISO.ToString() + "\' then (T1.\"DocTotal\" - T1.\"PaidToDate\") else (T1.\"DocTotalFC\" - T1.\"PaidFC\") end as Saldo, " +
                    "case when T1.\"DocSubType\" <> \'DN\' then 'Factura' else 'Nota Debito' end as Tipo, T1.\"DocCur\" as Moneda, T1.\"Comments\" as Comentarios, T3.\"SlpName\" as Vendedor, (select T8.\"Branch\" from \"OUSR\" as T8 where T8.\"USERID\" = T1.\"UserSign\") AS Sucursal from \"OINV\" as T1 " +
                    "inner join \"OCTG\" as T2 ON T2.\"GroupNum\" = T1.\"GroupNum\" " +
                    "left join \"OSLP\" as T3 ON T3.\"SlpCode\" = T1.\"SlpCode\" " +
                    //"where T1.\"CANCELED\" = \'N\' and UPPER(T2.\"PymntGroup\") LIKE '%CONTADO%' AND T1.\"DocStatus\" = \'O\' ";
                    "where T1.\"CANCELED\" = \'N\' AND T1.\"DocStatus\" = \'O\'" + condicion + " AND \"FolioPref\" IS NOT NULL AND \"FolioNum\" IS NOT NULL ";
                }

                if (!String.IsNullOrEmpty(pDesde.ToString()) && !String.IsNullOrEmpty(pHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                {
                    if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                        query += " and T1.DocDate >='" + pDesde.ToString(configAddOn.FormatoFecha) + "' and T1.DocDate <='" + pHasta.ToString(configAddOn.FormatoFecha) + "'";
                    else if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                        query += " and T1.\"DocDate\" >=\'" + pDesde.ToString(configAddOn.FormatoFecha) + "\' and T1.\"DocDate\" <=\'" + pHasta.ToString(configAddOn.FormatoFecha) + "\'";
                }

                query += " UNION ";

                if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                {
                    query += "select T1.DocEntry,T1.DocNum, T1.DocDate, T1.CardCode, T1.CardName, CONCAT(FolioPref,FolioNum) as CFE,  " +
                    "case when T1.DocCur = '" + monedaStrSimbolo.ToString() + "' or T1.DocCur = '" + monedaStrISO.ToString() + "' then (T1.DocTotal * -1) else (T1.DocTotalFC * -1) end as DocTotal, " +
                    "case when T1.DocCur = '" + monedaStrSimbolo.ToString() + "' or T1.DocCur = '" + monedaStrISO.ToString() + "' then ((T1.DocTotal - T1.PaidToDate) * -1) else ((T1.DocTotalFC - T1.PaidFC) * -1) end as Saldo, " +
                    "'Nota Crédito' as Tipo, T1.DocCur as Moneda, T1.Comments as Comentarios, T3.SlpName as Vendedor from ORIN as T1 " +
                    "inner join OCTG as T2 ON T2.GroupNum = T1.GroupNum " +
                    "left join OSLP as T3 ON T3.SlpCode = T1.SlpCode " +
                   // "where T1.CANCELED = 'N' and UPPER(T2.PymntGroup) LIKE '%CONTADO%' AND T1.DocStatus = 'O' ";
                   "where T1.CANCELED = 'N' AND T1.DocStatus = 'O'";
                    //"and t1.BPLId = '" + sucursalActiva + "'";
                }
                else
                {
                    query += "select T1.\"CardCode\",T1.\"DocEntry\",T1.\"DocNum\", T1.\"DocDate\", T1.\"CardCode\", T1.\"CardName\",CONCAT (\"FolioPref\", CONCAT (\'-\' , \"FolioNum\")) as CFE, " +
                    "case when T1.\"DocCur\" = \'" + monedaStrSimbolo.ToString() + "\' or T1.\"DocCur\" = \'" + monedaStrISO.ToString() + "\' then (T1.\"DocTotal\" * -1) else (T1.\"DocTotalFC\" * -1) end as DocTotal, " +
                    "case when T1.\"DocCur\" = \'" + monedaStrSimbolo.ToString() + "\' or T1.\"DocCur\" = \'" + monedaStrISO.ToString() + "\' then ((T1.\"DocTotal\" - T1.\"PaidToDate\") * -1) else ((T1.\"DocTotalFC\" - T1.\"PaidFC\") * -1) end as Saldo, " +
                    "'Nota Crédito' as Tipo, T1.\"DocCur\" as Moneda, T1.\"Comments\" as Comentarios, T3.\"SlpName\" as Vendedor , (select T8.\"Branch\" from \"OUSR\" as T8 where T8.\"USERID\" = T1.\"UserSign\") AS Sucursal from \"ORIN\" as T1 " +
                    "inner join \"OCTG\" as T2 ON T2.\"GroupNum\" = T1.\"GroupNum\" " +
                    "left join \"OSLP\" as T3 ON T3.\"SlpCode\" = T1.\"SlpCode\" " +
                    // "where T1.\"CANCELED\" = \'N\' and UPPER(T2.\"PymntGroup\") LIKE '%CONTADO%' AND T1.\"DocStatus\" = \'O\' ";
                    "where T1.\"CANCELED\" = \'N\' AND T1.\"DocStatus\" = \'O\' AND \"FolioPref\" IS NOT NULL AND \"FolioNum\" IS NOT NULL ";
                }

                if (!String.IsNullOrEmpty(pDesde.ToString()) && !String.IsNullOrEmpty(pHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                {
                    if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                        query += " and T1.DocDate >='" + pDesde.ToString(configAddOn.FormatoFecha) + "' and T1.DocDate <='" + pHasta.ToString(configAddOn.FormatoFecha) + "'";
                    else if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                        query += " and T1.\"DocDate\" >=\'" + pDesde.ToString(configAddOn.FormatoFecha) + "\' and T1.\"DocDate\" <=\'" + pHasta.ToString(configAddOn.FormatoFecha) + "\'";
                }

                if (tipoConexionBaseDatos.ToString().Equals("SQL") && !configAddOn.Empresa.Equals("ETAREY"))
                {
                    query += " order by T1.DocDate,T1.DocNum";
                }

                else if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query += " order by T1.\"DocDate\",T1.\"DocNum\"";

                string query2 = "";
                if (configAddOn.Empresa.Equals("ALMACEN"))
                {
                    query2 = "select * FROM (" + query + ") temp WHERE temp.\"DOCTOTAL\" > 0" + CondicionSuc;
                }

                if (!configAddOn.Empresa.Equals("ALMACEN"))
                    oRSMyTable.DoQuery(query);
                else
                    oRSMyTable.DoQuery(query2);

                return oRSMyTable;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al obtenerFacturasPendientes", ex.Message.ToString()); }
            return oRSMyTable;
        }

        private void llenarCombo(SAPbouiCOM.ComboBox pCombo, String pQuery, bool pSinRegistro, bool pBorrarRegistros, bool pTodosNinguno, bool tablaSistema)
        {
            try
            {
                SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRSMyTable.DoQuery(pQuery);

                if (pBorrarRegistros == true)
                {
                    try
                    {
                        int cant = pCombo.ValidValues.Count;

                        for (int i = cant; i > 0; i--) // Elimino los datos que tenga para cargarlos nuevamente
                            pCombo.ValidValues.Remove(i - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    catch (Exception ex)
                    { }
                }
                pCombo.ValidValues.Add("", "");
                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        if (tablaSistema)
                        {
                            pCombo.ValidValues.Add(oRSMyTable.Fields.Item("U_TERMINAL").Value, oRSMyTable.Fields.Item("U_TERMINAL").Value);
                        }
                        else
                        {
                            pCombo.ValidValues.Add(oRSMyTable.Fields.Item("Code").Value, oRSMyTable.Fields.Item("Name").Value);
                        }

                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }

                if (pSinRegistro == true)
                    pCombo.ValidValues.Add("-", "");

                if (pTodosNinguno == true)
                {
                    pCombo.ValidValues.Add("T", "Todos");
                    pCombo.ValidValues.Add("N", "Ning.");
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al llenarCombo", ex.Message.ToString()); }
        }

        private void llenarComboColumna(SAPbouiCOM.Column pColumna, String pQuery, bool pBorrarRegistros)
        {
            try
            {
                SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRSMyTable.DoQuery(pQuery);

                if (pBorrarRegistros == true)
                {
                    while (pColumna.ValidValues.Count > 0)
                        pColumna.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                while (!oRSMyTable.EoF)
                {
                    pColumna.ValidValues.Add(oRSMyTable.Fields.Item("Name").Value, oRSMyTable.Fields.Item("Code").Value);
                    oRSMyTable.MoveNext();
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al llenarComboColumna", ex.Message.ToString()); }
        }

        public string[] ObtenerCodCC(string ccemisor, string tarjetaTipo, string monedaSeleccionada)
        {
            string[] datosCC = new string[2];
            string query = "";
            string codigo = "";
            string cuenta = "";
            string ley = "";
            try
            {
                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                {
                    query += "select * from OCRC where CardName = '" + ccemisor + " " + tarjetaTipo + " " + monedaSeleccionada + "'";
                }
                else
                {
                    query += "select * from \"OCRC\" where \"CardName\" = '" + ccemisor + " " + tarjetaTipo + " " + monedaSeleccionada + "'";
                }

                oRSMyTable.DoQuery(query);

                codigo = oRSMyTable.Fields.Item("CreditCard").Value.ToString();
                cuenta = oRSMyTable.Fields.Item("AcctCode").Value.ToString();

                datosCC[0] = codigo;
                datosCC[1] = cuenta;


            }
            catch
            {
            }

            return datosCC;

        }

        public string ObtenerCambio()
        {
            string cambioMod = "";
            string query = "";
            string codigo = "";
            string cambio = "";
            string ley = "";
            try
            {

                DateTime fechaHoy = DateTime.Now;
                string fechaConsulta;
                fechaConsulta = fechaHoy.ToString(configAddOn.FormatoFecha);

                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                {
                    if (monedaSistema.Equals("$") || monedaStrISO.Equals("UYU"))
                    {
                        query += "select * from ORTT where Currency = 'USD' and RateDate = '" + fechaConsulta + "'";
                    }
                    else
                    {
                        query += "select * from ORTT where Currency = 'UYU' and RateDate = '" + fechaConsulta + "'";
                    }

                }
                else
                {
                    if (monedaSistema.Equals("$") || monedaStrISO.Equals("UYU"))
                    {
                        query += "select * from \"ORTT\" where \"Currency\" = 'USD' and \"RateDate\" = '" + fechaConsulta + "'";

                        if (configAddOn.Empresa.Equals("ALMACEN"))
                        {
                            query = "select * from \"ORTT\" where \"Currency\" = 'U$S' and \"RateDate\" = '" + fechaConsulta + "'";
                        }

                    }
                    else
                    {
                        query += "select * from \"ORTT\" where \"Currency\" = 'UYU' and \"RateDate\" = '" + fechaConsulta + "'";
                    }


                }
                oRSMyTable.DoQuery(query);

                cambio = oRSMyTable.Fields.Item("Rate").Value.ToString();
                double camb = Convert.ToDouble(cambio);
                //double cambioModificado = camb * 1.02;

                cambioMod = camb.ToString();
            }
            catch
            {
            }

            return cambioMod;

        }

        public string ObtenerCambioAlmacen(string moneda)
        {
            string cambioMod = "";
            string query = "";
            string codigo = "";
            string cambio = "";
            string ley = "";
            try
            {

                DateTime fechaHoy = DateTime.Now;
                string fechaConsulta;
                fechaConsulta = fechaHoy.ToString(configAddOn.FormatoFecha);

                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                if (moneda.Equals("$"))
                {

                    query = "select * from \"ORTT\" where \"Currency\" = 'U$C' and \"RateDate\" = '" + fechaConsulta + "'";


                }
                else
                {
                    query = "select * from \"ORTT\" where \"Currency\" = 'U$V' and \"RateDate\" = '" + fechaConsulta + "'";
                }


                oRSMyTable.DoQuery(query);

                cambio = oRSMyTable.Fields.Item("Rate").Value.ToString();
                double camb = Convert.ToDouble(cambio);
                //double cambioModificado = camb * 1.02;

                cambioMod = camb.ToString();
            }
            catch
            {
            }

            return cambioMod;

        }

        public string ObtenerCambioAlmacenFecha(string moneda, string fecha)
        {
            string cambioMod = "";
            string query = "";
            string codigo = "";
            string cambio = "";
            string ley = "";
            try
            {

                DateTime fechaHoy = DateTime.Now;
                string fechaConsulta;
                fechaConsulta = fechaHoy.ToString(configAddOn.FormatoFecha);

                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                if (moneda.Equals("$"))
                {

                    query = "select * from \"ORTT\" where \"Currency\" = 'U$C' and \"RateDate\" = '" + fecha + "'";


                }
                else
                {
                    query = "select * from \"ORTT\" where \"Currency\" = 'U$V' and \"RateDate\" = '" + fecha + "'";
                }


                oRSMyTable.DoQuery(query);

                cambio = oRSMyTable.Fields.Item("Rate").Value.ToString();
                // double camb = Convert.ToDouble(cambio);
                //double cambioModificado = camb * 1.02;

                cambioMod = cambio.ToString();
            }
            catch
            {
            }

            return cambioMod;

        }

        public string ObtenerPlan(string proveedor, string selloCod)
        {
            string plan = String.Empty;
            try
            {

                string comercio = String.Empty;

                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "SELECT TOP 1 * FROM \"@CESIONPAGOS\" where \"U_NOMEMPRESA\" = '" + proveedor + "' AND \"U_CODIGOFIN\" = '" + selloCod + "'";

                oRSMyTable.DoQuery(query);

                plan = oRSMyTable.Fields.Item("U_PLANID").Value.ToString();
                comercio = oRSMyTable.Fields.Item("U_CODCOMERCIO").Value.ToString();



            }
            catch (Exception)
            {

                return String.Empty;
            }

            return plan;
        }

        public string ObtenerComercio(string proveedor, string selloCod)
        {
            string comercio = String.Empty;

            try
            {
                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "SELECT TOP 1 * FROM \"@CESIONPAGOS\" where \"U_NOMEMPRESA\" = '" + proveedor + "' AND \"U_CODIGOFIN\" = '" + selloCod + "'";

                oRSMyTable.DoQuery(query);

                comercio = oRSMyTable.Fields.Item("U_CODCOMERCIO").Value.ToString();
            }
            catch (Exception)
            {
                return String.Empty;
            }

            return comercio;
        }


        private bool mandarPagoAPostAsync(List<clsPago> pDocumentos)
        {
            Boolean res = false;
            int docEntry = 0;
            Boolean tipoTarjeta = true;
            string terminal = "";
            string cardCode = "";
            string cardName = "";
            string monedaDoc = "";
            string operacion = "";
            DateTime fecha;
            double montoGravado = 0;
            string impuesto = "";
            string ley = "";
            double montoFactura = 0;
            double montoTotalFacturaNeto = 0;
            string descuentosFactura = "";
            string descuentosFacturasLineas = "";
            string tasaDeCambio = "";
            double montoTasaCambio = 0;
            string comboMoneda = "";
            string mensajeRespuestaTransact = "";
            string rut = "";
            int digitoVerificadorRespuesta = 0;
            bool rutValidado = false;
            bool consumidorFinal = true;
            bool continuarOperacion = true;
            string checkConsumidor = "Final";
            string monedaParaPago = "";
            double montoTotalFactura = 0;
            string numFolio = "";
            //Almacen Rural
            string usuarioPago = String.Empty;
            string numRecibo = String.Empty;
            string tipoFac = String.Empty;
            string proveedor = String.Empty;
            string sello = String.Empty;
            string Merchant = String.Empty;
            string Plan = String.Empty;

            CultureInfo culture = new CultureInfo("en-US");
            ControladorGeocom cg = new ControladorGeocom(this);

            try
            {
                // Monto caja aplicaccion
                SAPbouiCOM.EditText oMontoEf = oFormVisor.Items.Item("efMonto").Specific;
                double pagoVerificar = double.Parse(oMontoEf.String, culture);

                if (configAddOn.Empresa.Equals("ALMACEN"))
                {
                    SAPbouiCOM.ComboBox oComboUsu = oFormVisor.Items.Item("cmbUsr").Specific;
                    usuarioPago = oComboUsu.Selected.Value.ToString();
                    if (usuarioPago.Equals("manager"))
                    {
                        SBO_Application.MessageBox("Debe seleccionar un usuario diferente a manager.");
                        //return false;
                    }

                    SAPbouiCOM.EditText oRecibo = oFormVisor.Items.Item("txtRec").Specific;
                    numRecibo = oRecibo.String;

                    SAPbouiCOM.ComboBox oComboTipoFac = oFormVisor.Items.Item("TFac").Specific;
                    tipoFac = oComboTipoFac.Selected.Value.ToString();

                    try
                    {
                        SAPbouiCOM.ComboBox oComboProveedor = oFormVisor.Items.Item("cmbProv").Specific;
                        proveedor = oComboProveedor.Selected.Value.ToString();

                        SAPbouiCOM.ComboBox oComboSello = oFormVisor.Items.Item("cmbSell").Specific;
                        sello = oComboSello.Selected.Value.ToString();

                        if (String.IsNullOrEmpty(proveedor) || String.IsNullOrEmpty(sello))
                        {
                            SBO_Application.MessageBox("Debe seleccionar sello y proveedor.");
                            return false;
                        }
                    }
                    catch (Exception)
                    {

                        SBO_Application.MessageBox("Debe seleccionar sello y proveedor.");
                        return false;
                    }

                    if (!String.IsNullOrEmpty(proveedor) && !String.IsNullOrEmpty(sello))
                    {
                        Merchant = ObtenerComercio(proveedor, sello);
                        Plan = ObtenerPlan(proveedor, sello);

                        /*if (!Plan.Equals("0"))
                            //Merchant = "";
                        else
                            Plan = "0";*/

                        if (Plan.Equals("0") && Merchant.Equals("0"))
                        {
                            Merchant = "";
                            Plan = "0";
                        }
                        if (String.IsNullOrEmpty(Merchant) && String.IsNullOrEmpty(Plan))
                        {
                            SBO_Application.MessageBox("Error al obtener Merchant y Plan.");
                            return false;
                        }
                    }
                    else
                    { }
                }

                if (pagoVerificar <= pDocumentos[0].Monto && pagoVerificar > 0)
                {
                    SAPbouiCOM.StaticText oCambio = oFormVisor.Items.Item("lblTC").Specific;
                    tasaDeCambio = oCambio.Caption.ToString();
                    montoTasaCambio = double.Parse(tasaDeCambio, culture);
                    SAPbouiCOM.ComboBox OMoneda = oFormVisor.Items.Item("cmbMone").Specific;
                    comboMoneda = OMoneda.Selected.Value.ToString();

                    //Tomo valores del check box para validar si es consumidor final   
                    SAPbouiCOM.ComboBox oConsu = oFormVisor.Items.Item("cmbCons").Specific;
                    checkConsumidor = oConsu.Selected.Value.ToString();

                    try
                    {
                        docEntry = pDocumentos[0].DocEntry;
                        SAPbobsCOM.Recordset oRSMyTable = null;
                        String query = "";
                        try
                        {
                            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                                query += "select  LicTradNum, DocTotal, VatSumSy, FolioPref, FolioNum from OINV where DocEntry = '" + docEntry + "'";
                            else
                                query += "select \"LicTradNum\",\"DocTotal\",\"VatSumSy\",\"FolioPref\",\"FolioNum\" from \"OINV\" where \"DocEntry\" = '" + docEntry + "'";
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(ex.Message);
                        }

                        oRSMyTable.DoQuery(query);

                        rut = oRSMyTable.Fields.Item("LicTradNum").Value.ToString();
                        montoTotalFacturaNeto = Convert.ToDouble(oRSMyTable.Fields.Item("DocTotal").Value.ToString(), culture) - Convert.ToDouble(oRSMyTable.Fields.Item("VatSumSy").Value.ToString(), culture);
                        string rutAvalidar = "";
                        int digitoAvalidar = -1;
                        numFolio = oRSMyTable.Fields.Item("FolioPref").Value.ToString() + oRSMyTable.Fields.Item("FolioNum").Value.ToString();
                        numFolio = numFolio.Substring(1, numFolio.Length - 1);

                        if (rut.Length == 12)
                        {
                            rutAvalidar = rut.Substring(0, rut.Length - 1);
                            digitoAvalidar = Convert.ToInt32(rut.Substring(rut.Length - 1, 1));
                        }

                        digitoVerificadorRespuesta = validarRUT(rutAvalidar);

                        if (digitoAvalidar == digitoVerificadorRespuesta)
                            rutValidado = true;
                    }
                    catch
                    {
                        rutValidado = true;
                    }

                    if (checkConsumidor.Equals("Empresa") && !rutValidado)
                        continuarOperacion = false;
                    else if (checkConsumidor.Equals("Empresa") && rutValidado)
                        consumidorFinal = false;
                    else
                        consumidorFinal = true;

                    if (produccion && numFolio.Length > 1)
                    {
                        if (continuarOperacion)
                        {
                            if (tasaDeCambio.Length > 1)
                            {
                                if (pDocumentos.Count != 0)
                                {
                                    double montoTotalPago = 0;

                                    int lRetCode;

                                    cardCode = pDocumentos[0].CardCode;
                                    cardName = pDocumentos[0].CardName;
                                    fecha = pDocumentos[0].Fecha;
                                    monedaDoc = pDocumentos[0].Moneda;

                                    if (monedaDoc.Equals("$"))
                                        monedaDoc = "UYU";
                                    else if (monedaDoc.Equals("U$S"))
                                        monedaDoc = "USD";

                                    docEntry = pDocumentos[0].DocEntry;
                                    montoTotalFactura = pDocumentos[0].TotalFactura;
                                    SAPbobsCOM.Recordset oRSMyTable = null;
                                    String query = "";

                                    try
                                    {
                                        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                                        {
                                            if (monedaDoc.Equals(monedaSistema))
                                                query += "select VatSum as Impuesto, DiscSum as Descuento, LicTradNum from OINV where DocEntry = '" + docEntry + "'";
                                            else
                                                query += "select VatSumFC as Impuesto, DiscSumFC as Descuento, LicTradNum from OINV where DocEntry = '" + docEntry + "'";
                                        }
                                        else
                                        {
                                            if (monedaDoc.Equals(monedaSistema))
                                                query += "select \"VatSum\" as Impuesto, \"DiscSumSy\" as Descuento ,\"LicTradNum\"  from \"OINV\" where \"DocEntry\" = '" + docEntry + "'";
                                            else
                                                query += "select \"VatSumFC\" as Impuesto,\"DiscSumFC\" as Descuento ,\"LicTradNum\"  from \"OINV\" where \"DocEntry\" = '" + docEntry + "'";
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception(ex.Message);
                                    }

                                    oRSMyTable.DoQuery(query);

                                    impuesto = oRSMyTable.Fields.Item("Impuesto").Value.ToString();

                                    descuentosFactura = oRSMyTable.Fields.Item("Descuento").Value.ToString();
                                    //   descuentosFacturasLineas = obtenerDescuentosLineas(docEntry.ToString());

                                    // hay que modificar multiples pagos
                                    foreach (clsPago doc in pDocumentos)
                                    {
                                        //doc.Monto = double.Parse(doc.Monto.ToString()); // Para corregir decimales
                                        montoFactura += doc.Monto;
                                    }

                                    SAPbouiCOM.ComboBox oTerminal = oFormVisor.Items.Item("cmbTerm").Specific; //
                                    terminal = oTerminal.Selected.Description.ToString();
                                    //SAPbouiCOM.ComboBox oOperacion = oFormVisor.Items.Item("cmbOperac").Specific; // Operacion devolucion o venta

                                    //  if (!String.IsNullOrEmpty(oOperacion.Value.ToString()))
                                    //     operacion = oOperacion.Selected.Value.ToString();

                                    // prueba decreto de ley
                                    SAPbouiCOM.ComboBox oTarjeta = oFormVisor.Items.Item("cmdLey").Specific; //Si es debito o credito

                                    ley = oTarjeta.Selected.Value;

                                    if (proveedorPOS.Equals("TRANSACT"))
                                    {
                                        if (!String.IsNullOrEmpty(oTarjeta.Value.ToString()))
                                        {
                                            if (ley.Equals("No aplicar devolución"))
                                                decretoLey = 0;
                                            else if (ley.Equals("Aplicar devolución"))
                                                decretoLey = 1;
                                            else if (ley.Equals("17934 (Restaurant)"))
                                                decretoLey = 2;
                                            else if (ley.Equals("18083(IMESI)"))
                                                decretoLey = 3;
                                            else if (ley.Equals("18910 (AFAM)"))
                                                decretoLey = 4;
                                            else if (ley.Equals("18999 (Extranjeros)"))
                                                decretoLey = 5;
                                        }
                                        else
                                            decretoLey = 0;
                                    }
                                    else
                                    {
                                        if (!String.IsNullOrEmpty(oTarjeta.Value.ToString()))
                                        {
                                            if (ley.Equals("No aplicar devolución"))
                                                decretoLey = 0;
                                            else if (ley.Equals("Aplicar devolución"))
                                                decretoLey = 1;
                                            else if (ley.Equals("Devolución IMESI"))
                                                decretoLey = 2;
                                            else if (ley.Equals("Devolución AFAM"))
                                                decretoLey = 3;
                                            else if (ley.Equals("Devolución IVA"))
                                                decretoLey = 4;
                                            else if (ley.Equals("Reintegro IRPF"))
                                                decretoLey = 5;
                                        }
                                        else
                                            decretoLey = 0;
                                    }

                                    // combo cantidad cuotas
                                    SAPbouiCOM.ComboBox combocuotas = oFormVisor.Items.Item("cmbCT").Specific; //Si es debito o credito
                                    int cuotasCombo = 0;
                                    cuotasCombo = Convert.ToInt32(combocuotas.Selected.Value);

                                    // pago multiple
                                    if (!String.IsNullOrEmpty(oMontoEf.String) && Convert.ToDouble(oMontoEf.String) != montoTotalPago)
                                    {
                                        double numberDt;
                                        double neto;
                                        double porcent;
                                        double impuestoParcial;

                                        bool isNumeric = double.TryParse(oMontoEf.String, out numberDt);
                                        //Almacen Rural

                                        //if (tipoFac.Equals("Contado"))
                                        //{
                                        //    /* if (pagoVerificar != montoTotalFactura)
                                        //     {
                                        //         SBO_Application.MessageBox("El monto del pago debe ser por la totalidad de la factura.");
                                        //         return false;
                                        //     }*/
                                        //}
                                        //***********************************************//

                                        if (numberDt != 0 && Convert.ToDouble(oMontoEf.String) > 0)
                                        {
                                            //oDoc.CashSum = getDouble(oMontoEf.String);
                                            montoTotalPago = Math.Round(double.Parse(oMontoEf.String, culture), 2);
                                            if (montoTotalPago == montoTotalFactura)
                                            {
                                                //se calcula el monto gravado
                                                if (Convert.ToDouble(impuesto) != 0)
                                                    montoGravado = Math.Round((montoTotalFactura - (Convert.ToDouble(impuesto) + Convert.ToDouble(descuentosFactura))), 2);
                                                else
                                                    montoGravado = 0;
                                            }
                                            else
                                            {
                                                //se calcula el monto gravado
                                                if (Convert.ToDouble(impuesto) != 0)
                                                {
                                                    // montoGravado = Math.Round((montoTotalFactura - (Convert.ToDouble(impuesto) + Convert.ToDouble(descuentosFactura))), 2);
                                                    neto = Math.Round(montoTotalFactura - (Convert.ToDouble(impuesto) + Convert.ToDouble(descuentosFactura)), 2);
                                                    porcent = Math.Round((Convert.ToDouble(impuesto) * 100) / neto);
                                                    impuestoParcial = (montoTotalPago * porcent) / 100;
                                                    montoGravado = montoTotalPago - impuestoParcial;
                                                    impuesto = impuestoParcial.ToString();
                                                }
                                                else
                                                    montoGravado = 0;
                                            }
                                        }
                                        else
                                        {
                                            SBO_Application.MessageBox("El monto no es correcto.");
                                            return false;
                                        }
                                    }

                                    ///////////////////////////////////////////////////////////////////////////////////
                                    // Validar si el pago se realizara en Dolares o Pesos
                                    ///////////////////////////////////////////////////////////////////////////////////
                                    monedaParaPago = monedaDoc;

                                    if (comboMoneda.Equals("Dolares"))
                                    {
                                        if (monedaDoc.Equals("UYU"))
                                        {
                                            double tempImpuestos;
                                            montoTotalPago = Math.Round(montoTotalPago / montoTasaCambio, 2);
                                            //montoTotalPago = Math.Round(montoTotalPago, 3);
                                            montoGravado = Math.Round(montoGravado / montoTasaCambio, 2);
                                            //montoGravado = Math.Round(montoGravado, 3);
                                            // tempImpuestos = Math.Round(Convert.ToDouble(impuesto) / montoTasaCambio, 2);
                                            tempImpuestos = Convert.ToDouble(impuesto) * montoTasaCambio;
                                            //tempImpuestos = Math.Round(tempImpuestos, 3);
                                            impuesto = tempImpuestos.ToString();
                                            montoTotalFactura = Math.Round(montoTotalFactura / montoTasaCambio, 2);
                                            //cambio tipo de moneda para el pago
                                            monedaParaPago = "USD";
                                        }
                                    }
                                    else if (comboMoneda.Equals("Pesos"))
                                    {
                                        if (monedaDoc.Equals("USD"))
                                        {
                                            double tempImpuestos;
                                            montoTotalPago = montoTotalPago * montoTasaCambio;
                                            //montoTotalPago = Math.Round(montoTotalPago, 3);
                                            montoGravado = montoGravado * montoTasaCambio;
                                            //montoGravado = Math.Round(montoGravado, 3);
                                            tempImpuestos = Convert.ToDouble(impuesto) * montoTasaCambio;
                                            //tempImpuestos = Math.Round(tempImpuestos, 3);
                                            impuesto = tempImpuestos.ToString();
                                            montoTotalFactura = Math.Round(montoTotalFactura * montoTasaCambio, 2);

                                            //cambio tipo de moneda para el pago
                                            monedaParaPago = "UYU";
                                        }
                                    }

                                    SBO_Application.MessageBox("Ingrese Tarjeta.");

                                    //Redondeo de decimales para no tener diferencias
                                    /* if ((monedaDoc.Equals("USD") && comboMoneda.Equals("Pesos")) || (monedaDoc.Equals("UYU") && comboMoneda.Equals("Dolares")))
                                     {
                                         montoTotalPago += 0.010;
                                         montoTotalFactura += 0.010;
                                     }*/

                                    //SBO_Application.MessageBox("Monto a Pagar: " + montoTotalPago + "\n" + "IVA: " + impuesto + "\n Descuentos: " + descuentosFactura + "\n Monto Gravado: " + montoGravado);

                                    //Llamada a Geocom
                                    if (proveedorPOS.Equals("GEOCOM"))
                                    {
                                        LogGeocom objetoLog = new LogGeocom();
                                        //datos de test
                                        string PosID = terminal;
                                        string systemId = configAddOn.hash;
                                        string Branch = "Almacen";
                                        string clientAppId = "1";
                                        string userId = "1";
                                        double taxAmountTemp = 0;

                                        //Validar si el monto de IVA es 0
                                        if (impuesto.Equals("0")) decretoLey = 0;

                                        GeocomWSProductivo.PurchaseQueryResponse respuesta = cg.enviarVentaPosGeocom(montoTotalPago, PosID, monedaParaPago, montoGravado, Convert.ToDouble(impuesto), numFolio, cuotasCombo, decretoLey, montoTotalFactura, systemId, Branch, clientAppId, userId, Merchant, Plan, sello);

                                        //buscar en las tablas el codigo de respuesta
                                        if (!respuesta.PosResponseCode.Equals("-1"))
                                        {
                                            CodRespuestaPOSGeocom codPosRespuesta = CodigoPOSGeocom(respuesta.PosResponseCode);

                                            if (codPosRespuesta.estado.Equals("OK"))
                                            {
                                                try
                                                {
                                                    if (String.IsNullOrEmpty(respuesta.OriginCardType)) objetoLog.cardtype = "-"; else objetoLog.cardtype = respuesta.OriginCardType;
                                                    if (String.IsNullOrEmpty(respuesta.Ci)) objetoLog.ci = "-"; else objetoLog.ci = respuesta.Ci;
                                                    if (String.IsNullOrEmpty(respuesta.AuthorizationCode)) objetoLog.codigoAutorizacion = "-"; else objetoLog.codigoAutorizacion = respuesta.AuthorizationCode;
                                                    if (String.IsNullOrEmpty(codPosRespuesta.codigo)) objetoLog.codigoRespuestaPos = "-"; else objetoLog.codigoRespuestaPos = codPosRespuesta.codigo;
                                                    if (String.IsNullOrEmpty(codPosRespuesta.estado)) objetoLog.codigoRespuestaPosDescripcion = "-"; else objetoLog.codigoRespuestaPosDescripcion = codPosRespuesta.estado;
                                                    if (String.IsNullOrEmpty(respuesta.Quota)) objetoLog.cuotas = "-"; else objetoLog.cuotas = respuesta.Quota;
                                                    objetoLog.TransactionDateTime = respuesta.TransactionDate;
                                                    objetoLog.fechaTransaccion = DateTime.Now;
                                                    objetoLog.transactionDateTime = respuesta.TransactionHour;
                                                    if (String.IsNullOrEmpty(respuesta.TaxRefund)) objetoLog.impuestocodigo = "-"; else objetoLog.impuestocodigo = respuesta.TaxRefund;
                                                    if (String.IsNullOrEmpty(respuesta.Batch)) objetoLog.lote = "-"; else objetoLog.lote = respuesta.Batch;
                                                    if (String.IsNullOrEmpty(respuesta.Currency)) objetoLog.monedaTransaccionCod = "-"; else objetoLog.monedaTransaccionCod = respuesta.Currency;
                                                    //Proveedor Almacen Rural
                                                    objetoLog.Merchant = proveedor;

                                                    try
                                                    {
                                                        if (!string.IsNullOrEmpty(respuesta.TaxAmount))
                                                            taxAmountTemp = Convert.ToDouble(respuesta.TaxAmount) / 100;
                                                        double InvoiceAmount = Convert.ToDouble(respuesta.TotalAmount) / 100;
                                                        objetoLog.TaxableAmount = taxAmountTemp.ToString();
                                                        objetoLog.TaxRefund = respuesta.TaxRefund;
                                                        objetoLog.InvoiceAmount = InvoiceAmount.ToString();
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                    }

                                                    if (objetoLog.monedaTransaccionCod.Equals("858"))
                                                        objetoLog.monedaTransaccionDescrip = "UYU";
                                                    else if (objetoLog.monedaTransaccionCod.Equals("840"))
                                                        //moneda = "0840"; //dolares
                                                        objetoLog.monedaTransaccionDescrip = "USD";
                                                    else
                                                        objetoLog.monedaTransaccionDescrip = "-";

                                                    //En esta seccion se buscara en la tabla de tarjetas y se traera el Issuer y su cuenta
                                                    IssuerGeocom issuerTemp = ObtenerDatosIssuerGeocom(respuesta.Issuer.ToString(), monedaDoc, proveedor);
                                                    objetoLog.issuerCode = respuesta.Issuer.ToString();
                                                    if (String.IsNullOrEmpty(objetoLog.issuerCode)) objetoLog.issuerCode = "-";
                                                    if (String.IsNullOrEmpty(issuerTemp.nombreTarjeta)) objetoLog.issuerCodeDescripcion = "-"; else objetoLog.issuerCodeDescripcion = issuerTemp.nombreTarjeta;
                                                    // objetoLog.issuerCodeDescripcion = issuerTemp.nombreTarjeta;//ir a buscar en tabla nativa tarjetas
                                                    if (String.IsNullOrEmpty(issuerTemp.cuentaContable)) objetoLog.cuentaTarjeta = "-"; else objetoLog.cuentaTarjeta = issuerTemp.cuentaContable;
                                                    //objetoLog.cuentaTarjeta = issuerTemp.cuentaContable;
                                                    if (String.IsNullOrEmpty(issuerTemp.codigoTarjetaSAP)) objetoLog.codigoTarjetaSAP = "-"; else objetoLog.codigoTarjetaSAP = issuerTemp.codigoTarjetaSAP;
                                                    // objetoLog.codigoTarjetaSAP = issuerTemp.codigoTarjetaSAP;

                                                    //--------------------------------------------------------------------------------------------------//

                                                    double montoTemp = Convert.ToDouble(respuesta.TotalAmount) / 100;
                                                    objetoLog.monto = montoTemp.ToString();

                                                    if (String.IsNullOrEmpty(respuesta.CardOwnerName)) objetoLog.nombre = "-"; else objetoLog.nombre = respuesta.CardOwnerName;
                                                    if (String.IsNullOrEmpty(respuesta.EmvApplicationName)) objetoLog.nombreTarjeta = "-"; else objetoLog.nombreTarjeta = respuesta.EmvApplicationName;
                                                    if (String.IsNullOrEmpty(respuesta.CardNumber)) objetoLog.numerotarjeta = "-"; else objetoLog.numerotarjeta = respuesta.CardNumber;
                                                    if (String.IsNullOrEmpty(respuesta.Plan)) objetoLog.plan = "-"; else objetoLog.plan = respuesta.Plan;
                                                    if (String.IsNullOrEmpty(respuesta.PosID)) objetoLog.posId = "-"; else objetoLog.posId = respuesta.PosID;
                                                    if (String.IsNullOrEmpty(respuesta.AcquirerTerminal)) objetoLog.terminal = "-"; else objetoLog.terminal = respuesta.AcquirerTerminal;
                                                    if (String.IsNullOrEmpty(respuesta.Ticket)) objetoLog.ticket = ""; else objetoLog.ticket = respuesta.Ticket;

                                                    // ir a buscar en tabla SellosGeocom
                                                    objetoLog.selloCod = respuesta.Acquirer.ToString();
                                                    objetoLog.selloDescripcion = ObtenerSelloGeocom(objetoLog.selloCod);
                                                    if (String.IsNullOrEmpty(objetoLog.selloDescripcion)) objetoLog.selloDescripcion = "-";

                                                    //se va a buscar a la tabla TipoTranGeocom
                                                    if (String.IsNullOrEmpty(respuesta.TransactionType)) objetoLog.transaccionType = "-"; else objetoLog.transaccionType = respuesta.TransactionType;
                                                    objetoLog.transaccionTypeDescripcion = TipoTransaccionGeocom(objetoLog.transaccionType);
                                                    if (String.IsNullOrEmpty(objetoLog.transaccionTypeDescripcion)) objetoLog.transaccionTypeDescripcion = "-";

                                                    objetoLog.EstatusGeocomTransaccion = "OK";

                                                    if (decretoLey == 0)
                                                    {
                                                        SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                                                        oDoc.CardCode = pDocumentos[0].CardCode;
                                                        oDoc.CardName = pDocumentos[0].CardName;
                                                        oDoc.TransferDate = pDocumentos[0].Fecha;
                                                        oDoc.Series = 73;

                                                        // oDoc.DocCurrency = pDocumentos[0].Moneda;
                                                        oDoc.DocCurrency = monedaDoc;

                                                        montoTotalPago = 0;

                                                        foreach (clsPago doc in pDocumentos)
                                                        {
                                                            if (monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU"))
                                                                doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100) / montoTasaCambio; // Para corregir decimales
                                                            //doc.Monto = (TextoaDecimal(respuesta.DatosTransaccion.Monto.ToString()) / 100) * montoTasaCambio; // Para corregir decimales
                                                            else if (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))
                                                                doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100) * montoTasaCambio;
                                                            else
                                                                doc.Monto = TextoaDecimal(respuesta.TotalAmount.ToString()) / 100;

                                                            montoTotalPago += doc.Monto;

                                                            oDoc.Invoices.DocEntry = doc.DocEntry;

                                                            if (configAddOn.Empresa.Equals("ALMACEN"))
                                                            {
                                                                if (monedaDoc.Equals("USD"))
                                                                    oDoc.DocCurrency = "U$S";
                                                                else
                                                                    oDoc.DocCurrency = "$";
                                                            }

                                                            if (monedaDoc.Equals(monedaSistema)) // Si el documento es en Moneda Local
                                                                oDoc.Invoices.SumApplied = doc.Monto;
                                                            else
                                                                oDoc.Invoices.AppliedFC = doc.Monto;

                                                            if (doc.Monto >= 0)
                                                                oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                                                            else
                                                                oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                                                            oDoc.Invoices.Add();
                                                        }

                                                        /*SAPbouiCOM.ComboBox oMedioPago = oFormVisor.Items.Item("tjaDesc").Specific; // Tarjeta seleccionada
                                                        if (!String.IsNullOrEmpty(oMedioPago.Value.ToString()))
                                                            oDoc.CreditCards.CreditCard = 1;//Convert.ToInt32(oMedioPago.Selected.Value.ToString());*/
                                                        if (configAddOn.Empresa.Equals("ALMACEN"))
                                                        {
                                                            //Se cambia a una variable temporal, ya que la parametrización de Almacen es diferente al resto
                                                            string monedaDocTemp = "";
                                                            if (oDoc.DocCurrency.Equals("$")) monedaDocTemp = "UYU";
                                                            if (oDoc.DocCurrency.Equals("U$S")) monedaDocTemp = "USD";

                                                            if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))) //ASPL - 2020.08.18
                                                            {
                                                                oDoc.CashSum = montoTotalPago;

                                                                if (monedaDocTemp.Equals(monedaSistema))
                                                                    oDoc.CashAccount = "1113000001";
                                                                else
                                                                    oDoc.CashAccount = "1114000002";
                                                            }

                                                            // obtener id de tarjeta de la tabla OCRC
                                                            if (!String.IsNullOrEmpty(objetoLog.codigoTarjetaSAP))
                                                                oDoc.CreditCards.CreditCard = Convert.ToInt32(objetoLog.codigoTarjetaSAP);

                                                            if (!String.IsNullOrEmpty(objetoLog.cuentaTarjeta.ToString()))
                                                                oDoc.CreditCards.CreditAcct = objetoLog.cuentaTarjeta.ToString();
                                                            //else
                                                            //{
                                                            //    if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                                            //        oDoc.CreditCards.CreditAcct = configAddOn.TarjetaMN;
                                                            //    else
                                                            //        oDoc.CreditCards.CreditAcct = configAddOn.TarjetaME;
                                                            //}

                                                            String oNumeroTja = objetoLog.numerotarjeta; // Numero de tarjeta
                                                            oNumeroTja = oNumeroTja.Substring(oNumeroTja.Length - 4, 4);

                                                            if (!String.IsNullOrEmpty(oNumeroTja))
                                                            {
                                                                oDoc.CreditCards.CreditCardNumber = oNumeroTja;
                                                                if (oDoc.CreditCards.CreditCardNumber.Length > 4)
                                                                    oDoc.CreditCards.CreditCardNumber = oDoc.CreditCards.CreditCardNumber.Substring(0, 4);
                                                            }

                                                            string oFecha = objetoLog.fechaTransaccion.ToString(); // Fecha
                                                            if (!String.IsNullOrEmpty(oFecha))
                                                            {
                                                                string dueDateStr = "";
                                                                DateTime fechatransaccion = Convert.ToDateTime(oFecha);
                                                                dueDateStr = String.Format("{0}-{1}-{2}", fechatransaccion.Year, fechatransaccion.Month, fechatransaccion.Day);
                                                                oDoc.CreditCards.CardValidUntil = Convert.ToDateTime(dueDateStr);
                                                            }

                                                            String oCantCuotas = objetoLog.cuotas; // Cant de cuotas
                                                            if (!String.IsNullOrEmpty(oCantCuotas))
                                                                oDoc.CreditCards.NumOfPayments = Convert.ToInt32(oCantCuotas);

                                                            String oVoucherNro = objetoLog.ticket.ToString(); // Nro Certificado
                                                            if (!String.IsNullOrEmpty(oVoucherNro))
                                                                oDoc.CreditCards.VoucherNum = oVoucherNro;

                                                            String oOwnerId = objetoLog.selloCod; // Id Tja
                                                            if (!String.IsNullOrEmpty(oOwnerId))
                                                                oDoc.CreditCards.OwnerIdNum = oOwnerId;

                                                            // se comentan observaciones para pago por tarjeta
                                                            //SAPbouiCOM.EditText oObservaciones = oFormVisor.Items.Item("txObser").Specific; // Se agrega campo Observaciones 28/12/18
                                                            //if (!String.IsNullOrEmpty(oObservaciones.String))
                                                            //    oDoc.Remarks = oObservaciones.String; // Observaciones

                                                            //crear metodo para recibir el codigo
                                                            oDoc.CreditCards.CreditSum = montoTotalPago;
                                                            oDoc.CreditCards.NumOfCreditPayments = oDoc.CreditCards.NumOfPayments;

                                                            //Este codigo sale del metodo de pago creado en SAP, Tabla OCRP
                                                            oDoc.CreditCards.PaymentMethodCode = 3;//oDoc.CreditCards.NumOfPayments;
                                                            oDoc.CreditCards.ConfirmationNum = objetoLog.numerotarjeta;
                                                            //Se guarda en comentarios la moneda en la cual se realizo la transaccion
                                                            oDoc.Remarks = monedaParaPago;
                                                        }
                                                        // se agrega usuario logueado a documento de pago
                                                        oDoc.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;
                                                        //Numero de recibo
                                                        //oDoc.CounterReference = numRecibo; //ASPL - 2022.02.18, Campo vacio
                                                        //Almacen Rural Referencia Pago 2% IVA
                                                        oDoc.UserFields.Fields.Item("U_PagoRef").Value = "No Aplica";


                                                        Random r = new Random();
                                                        oDoc.CounterReference = oDoc.Invoices.DocEntry.ToString() + "CJ" + r.Next(9999);
                                                        //else
                                                        //{
                                                        //    if (oDoc.DocCurrency.ToString().Equals(monedaSistema))
                                                        //        oDoc.CashAccount = "1113000001";
                                                        //    else
                                                        //        oDoc.CashAccount = "1114000002";
                                                        //}

                                                        if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                                                        {
                                                            lRetCode = oDoc.Add();
                                                            string ultimoDoc = "";
                                                            oCompany.GetNewObjectCode(out ultimoDoc);

                                                            if (lRetCode != 0)
                                                            {
                                                                objetoLog.EstatusSAPTransaccion = "Error";
                                                                guardaLogGeocom(objetoLog, true);
                                                                SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ". Codigo error:" + lRetCode.ToString());
                                                                res = false;
                                                            }
                                                            else
                                                            {
                                                                if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))) //ASPL - 2021.07.15 - Nuevo proceso por pago en diferente monedas.
                                                                {
                                                                    res = CrearAsiento(monedaDoc, monedaParaPago, dbCambio, montoTotalPago);

                                                                    if (res)
                                                                    {
                                                                        /*var serviceLayer = new SLConnection("https://192.168.101.21:50000/b1s/v1", "DB_PRODUCCION", usuarioLogueado, contrasena); //Cambio Service Layer Nicolas Pecoy


                                                                        serviceLayer.AfterCall(async call =>
                                                                        {
                                                                            Console.WriteLine($"Request: {call.HttpRequestMessage.Method} {call.HttpRequestMessage.RequestUri}");
                                                                            Console.WriteLine($"Body sent: {call.RequestBody}");
                                                                            Console.WriteLine($"Response: {call.HttpResponseMessage?.StatusCode}");
                                                                            Console.WriteLine(await call.HttpResponseMessage?.Content?.ReadAsStringAsync());
                                                                            Console.WriteLine($"Call duration: {call.Duration.Value.TotalSeconds} seconds");
                                                                        });

                                                                        string monedaP = "";
                                                                        string cuentaPago = "";
                                                                        if (monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) //ASPL - 2021.07.19 - Nuevo proceso por pago en diferente monedas.
                                                                        {
                                                                            monedaP = "$";
                                                                            cuentaPago = "1113000001";
                                                                        }
                                                                        else if (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))
                                                                        {
                                                                            monedaP = "U$S";
                                                                            cuentaPago = "1114000002";
                                                                        }

                                                                        montoTotalPago = 0;

                                                                        foreach (clsPago doc in pDocumentos)
                                                                        {
                                                                            if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD")))
                                                                                doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100);

                                                                            montoTotalPago += doc.Monto;
                                                                        }

                                                                        IssuerGeocom issuerAux = ObtenerDatosIssuerGeocom(respuesta.Issuer.ToString(), monedaParaPago, proveedor);
                                                                        string CreditAcct = "";
                                                                        if (!String.IsNullOrEmpty(issuerAux.cuentaContable.ToString()))
                                                                            CreditAcct = issuerAux.cuentaContable.ToString();
                                                                        else
                                                                        {
                                                                            if (monedaP.ToString().Equals(monedaStrISO) || monedaP.ToString().Equals(monedaStrSimbolo))
                                                                                CreditAcct = configAddOn.TarjetaMN;
                                                                            else
                                                                                CreditAcct = configAddOn.TarjetaME;
                                                                        }
                                                                        int tarjeta = 0;
                                                                        if (!String.IsNullOrEmpty(objetoLog.codigoTarjetaSAP))
                                                                            tarjeta = Convert.ToInt32(objetoLog.codigoTarjetaSAP);

                                                                        string oNumeroTja = objetoLog.numerotarjeta; // Numero de tarjeta
                                                                        oNumeroTja = oNumeroTja.Substring(oNumeroTja.Length - 4, 4);

                                                                        if (!String.IsNullOrEmpty(oNumeroTja))
                                                                        {
                                                                            oNumeroTja = oNumeroTja;

                                                                            if (oNumeroTja.Length > 4)
                                                                                oNumeroTja = oNumeroTja.Substring(0, 4);
                                                                        }

                                                                        string oFecha = objetoLog.fechaTransaccion.ToString(); // Fecha
                                                                        DateTime fechaP = DateTime.Now;
                                                                        if (!String.IsNullOrEmpty(oFecha))
                                                                        {
                                                                            string dueDateStr = "";
                                                                            DateTime fechatransaccion = Convert.ToDateTime(oFecha);
                                                                            dueDateStr = String.Format("{0}-{1}-{2}", fechatransaccion.Year, fechatransaccion.Month, fechatransaccion.Day);
                                                                            fechaP = Convert.ToDateTime(dueDateStr);
                                                                        }

                                                                        int oCantCuotas = 0; // Cant de cuotas

                                                                        if (!String.IsNullOrEmpty(objetoLog.cuotas))
                                                                            oCantCuotas = Convert.ToInt32(objetoLog.cuotas);

                                                                        string oVoucherNro = ""; // Nro Certificado
                                                                        if (!String.IsNullOrEmpty(objetoLog.ticket.ToString()))
                                                                            oVoucherNro = objetoLog.ticket.ToString();

                                                                        string oOwnerId = ""; // Id Tja
                                                                        if (!String.IsNullOrEmpty(objetoLog.selloCod))
                                                                            oOwnerId = objetoLog.selloCod;

                                                                        Pago pg = new Pago
                                                                        {
                                                                            DocType = "rAccount",
                                                                            DocTypte = "rAccount",
                                                                            DocDate = DateTime.Now.ToString("yyyy-MM-dd"),
                                                                            TransferDate = pDocumentos[0].Fecha,
                                                                            DocObjectCode = "bopot_IncomingPayments",
                                                                            DocCurrency = monedaP,
                                                                            Remarks = monedaParaPago,
                                                                            U_PagoRef = ultimoDoc,
                                                                            U_Usuario = usuarioPago,
                                                                            CounterReference = numRecibo,
                                                                            PaymentAccounts = new Paymentaccount[] {
                                                                                new Paymentaccount {
                                                                                    AccountCode = cuentaPago,
                                                                                    SumPaid = Convert.ToSingle(montoTotalPago),
                                                                                    Decription = "Pago Tarjeta POS - " + pDocumentos[0].CardName
                                                                                }

                                                                            },
                                                                            PaymentCreditCards = new Paymentcreditcard[] {
                                                                                new Paymentcreditcard {
                                                                                    CreditAcct = CreditAcct,
                                                                                    CreditCard = tarjeta,
                                                                                    CreditCardNumber = oNumeroTja,
                                                                                    CardValidUntil = fechaP.ToString("yyyy-MM-dd"),
                                                                                    NumOfPayments = oCantCuotas,
                                                                                    VoucherNum = oVoucherNro,
                                                                                    OwnerIdNum = oOwnerId,
                                                                                    CreditSum = Convert.ToSingle(montoTotalPago - taxAmountTemp),
                                                                                    NumOfCreditPayments = 1,
                                                                                    PaymentMethodCode = 2,
                                                                                    ConfirmationNum = objetoLog.numerotarjeta,
                                                                                }
                                                                            }
                                                                        };

                                                                        agregarPagoServiceLayerAsync(serviceLayer, pg).RunSynchronously();*/

                                                                        //var pagoCreado = await serviceLayer.Request("IncomingPayments").PostAsync<Pago>(pg);



                                                                        int res2 = 0;
                                                                        //************ CREACIÓN SEGUNDO DOCUMENTO PAGO A CUENTA ************************//

                                                                        SAPbobsCOM.Payments pagoCuenta = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                                                                        pagoCuenta.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                                                                        pagoCuenta.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
                                                                        pagoCuenta.DocDate = DateTime.Now;
                                                                        pagoCuenta.TransferDate = pDocumentos[0].Fecha;

                                                                        pagoCuenta.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
                                                                        pagoCuenta.DocCurrency = monedaParaPago;

                                                                        if (monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) //ASPL - 2021.07.19 - Nuevo proceso por pago en diferente monedas.
                                                                        {
                                                                            pagoCuenta.DocCurrency = "$";
                                                                            pagoCuenta.AccountPayments.AccountCode = "1113000001";
                                                                        }
                                                                        else if (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))
                                                                        {
                                                                            pagoCuenta.DocCurrency = "U$S";
                                                                            pagoCuenta.AccountPayments.AccountCode = "1114000002";
                                                                        }

                                                                        montoTotalPago = 0;
                                                                        IssuerGeocom issuerAux = ObtenerDatosIssuerGeocom(respuesta.Issuer.ToString(), monedaParaPago, proveedor);
                                                                        string CreditAcct = "";
                                                                        if (!String.IsNullOrEmpty(issuerAux.cuentaContable.ToString()))
                                                                            CreditAcct = issuerAux.cuentaContable.ToString();
                                                                        else
                                                                        {
                                                                            if (monedaParaPago.ToString().Equals(monedaStrISO) || monedaParaPago.ToString().Equals(monedaStrSimbolo))
                                                                                CreditAcct = configAddOn.TarjetaMN;
                                                                            else
                                                                                CreditAcct = configAddOn.TarjetaME;
                                                                        }
                                                                        foreach (clsPago doc in pDocumentos)
                                                                        {
                                                                            if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD")))
                                                                                doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100);

                                                                            montoTotalPago += doc.Monto;
                                                                        }

                                                                        //Almacen Rural Canje
                                                                        pagoCuenta.AccountPayments.SumPaid = montoTotalPago;
                                                                        //  pagoCuenta.AccountPayments.SumPaid = 10;
                                                                        pagoCuenta.AccountPayments.Decription = "Pago Tarjeta POS - " + pDocumentos[0].CardName;
                                                                        pagoCuenta.AccountPayments.Add();

                                                                        if (!String.IsNullOrEmpty(issuerAux.cuentaContable.ToString()))
                                                                            pagoCuenta.CreditCards.CreditAcct = issuerAux.cuentaContable.ToString();
                                                                        else
                                                                        {
                                                                            if (pagoCuenta.DocCurrency.ToString().Equals(monedaStrISO) || pagoCuenta.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                                                                pagoCuenta.CreditCards.CreditAcct = configAddOn.TarjetaMN;
                                                                            else
                                                                                pagoCuenta.CreditCards.CreditAcct = configAddOn.TarjetaME;
                                                                        }

                                                                        // obtener id de tarjeta de la tabla OCRC
                                                                        if (!String.IsNullOrEmpty(objetoLog.codigoTarjetaSAP))
                                                                            pagoCuenta.CreditCards.CreditCard = Convert.ToInt32(objetoLog.codigoTarjetaSAP);

                                                                        string oNumeroTja = objetoLog.numerotarjeta; // Numero de tarjeta
                                                                        oNumeroTja = oNumeroTja.Substring(oNumeroTja.Length - 4, 4);

                                                                        if (!String.IsNullOrEmpty(oNumeroTja))
                                                                        {
                                                                            pagoCuenta.CreditCards.CreditCardNumber = oNumeroTja;

                                                                            if (pagoCuenta.CreditCards.CreditCardNumber.Length > 4)
                                                                                pagoCuenta.CreditCards.CreditCardNumber = pagoCuenta.CreditCards.CreditCardNumber.Substring(0, 4);
                                                                        }

                                                                        string oFecha = objetoLog.fechaTransaccion.ToString(); // Fecha

                                                                        if (!String.IsNullOrEmpty(oFecha))
                                                                        {
                                                                            string dueDateStr = "";
                                                                            DateTime fechatransaccion = Convert.ToDateTime(oFecha);
                                                                            dueDateStr = String.Format("{0}-{1}-{2}", fechatransaccion.Year, fechatransaccion.Month, fechatransaccion.Day);
                                                                            pagoCuenta.CreditCards.CardValidUntil = Convert.ToDateTime(dueDateStr);
                                                                        }

                                                                        string oCantCuotas = objetoLog.cuotas; // Cant de cuotas

                                                                        if (!String.IsNullOrEmpty(oCantCuotas))
                                                                            pagoCuenta.CreditCards.NumOfPayments = Convert.ToInt32(oCantCuotas);

                                                                        string oVoucherNro = objetoLog.ticket.ToString(); // Nro Certificado
                                                                        if (!String.IsNullOrEmpty(oVoucherNro))
                                                                            pagoCuenta.CreditCards.VoucherNum = oVoucherNro;

                                                                        string oOwnerId = objetoLog.selloCod; // Id Tja
                                                                        if (!String.IsNullOrEmpty(oOwnerId))
                                                                            pagoCuenta.CreditCards.OwnerIdNum = oOwnerId;

                                                                        //crear metodo para recibir el codigo
                                                                        pagoCuenta.CreditCards.CreditSum = montoTotalPago - taxAmountTemp;
                                                                        //pagoCuenta.CreditCards.CreditSum = 5;
                                                                        pagoCuenta.CreditCards.NumOfCreditPayments = pagoCuenta.CreditCards.NumOfPayments;

                                                                        //Este codigo sale del metodo de pago creado en SAP, Tabla OCRP
                                                                        pagoCuenta.CreditCards.PaymentMethodCode = 3;//oDoc.CreditCards.NumOfPayments;
                                                                        pagoCuenta.CreditCards.ConfirmationNum = objetoLog.numerotarjeta;
                                                                        //Se guarda en comentarios la moneda en la cual se realizo la transaccion
                                                                        //En almacen Rural en este campo guardamos referencia del documento anterior
                                                                        pagoCuenta.Remarks = monedaParaPago;

                                                                        //Almacen Rural Referencia Pago 2% IVA
                                                                        pagoCuenta.UserFields.Fields.Item("U_PagoRef").Value = ultimoDoc;
                                                                        // se agrega usuario logueado a documento de pago
                                                                        pagoCuenta.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;
                                                                        //Numero de recibo
                                                                        pagoCuenta.CounterReference = numRecibo; //ASPL - 2022.02.18, se cambia a campo vacio.

                                                                        //Random r = new Random();
                                                                        //pagoCuenta.CounterReference = r.Next(99) + "CJ" + r.Next(999);

                                                                        res2 = pagoCuenta.Add();

                                                                        if (res2 != 0)
                                                                        {
                                                                            objetoLog.EstatusSAPTransaccion = "Error";
                                                                            guardaLogGeocom(objetoLog, true);
                                                                            SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ".Error Documento Canje -  Codigo error:" + res2.ToString()); //Error Almacen analizar Nicolas Pecoy
                                                                            res = false;
                                                                        }
                                                                        else
                                                                        {
                                                                            objetoLog.EstatusSAPTransaccion = "OK";
                                                                            res = true;
                                                                            guardaLogGeocom(objetoLog, true);
                                                                            //guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString());
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    objetoLog.EstatusSAPTransaccion = "OK";
                                                                    res = true;
                                                                    guardaLogGeocom(objetoLog, true);
                                                                    //guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString());
                                                                }
                                                            }
                                                        }
                                                        //guardaLogGeocom(objetoLog, true);
                                                        //guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaTitular.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaNombre.ToString(), respuesta.TarjetaTipo.ToString());
                                                    }
                                                    else
                                                    {
                                                        //************DOCUMENTO EFECTIVO PARA INGRESAR PAGO DE TARJETA CREDITO/DEBITO*********//
                                                        #region ***DOCUMENTO PAGO DE TARJETA CREDITO/DEBITO***
                                                        SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                                                        oDoc.CardCode = pDocumentos[0].CardCode;
                                                        oDoc.CardName = pDocumentos[0].CardName;
                                                        oDoc.TransferDate = pDocumentos[0].Fecha;

                                                        // oDoc.DocCurrency = pDocumentos[0].Moneda;
                                                        oDoc.DocCurrency = monedaDoc;

                                                        montoTotalPago = 0;

                                                        foreach (clsPago doc in pDocumentos)
                                                        {
                                                            if (monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU"))
                                                                doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100) / montoTasaCambio; // Para corregir decimales
                                                            //doc.Monto = (TextoaDecimal(respuesta.DatosTransaccion.Monto.ToString()) / 100) * montoTasaCambio; // Para corregir decimales
                                                            else if (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))
                                                                doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100) * montoTasaCambio;
                                                            else
                                                                doc.Monto = TextoaDecimal(respuesta.TotalAmount.ToString()) / 100;

                                                            montoTotalPago += doc.Monto;

                                                            oDoc.Invoices.DocEntry = doc.DocEntry;

                                                            if (configAddOn.Empresa.Equals("ALMACEN"))
                                                            {
                                                                if (monedaDoc.Equals("USD"))
                                                                    oDoc.DocCurrency = "U$S";
                                                                else
                                                                    oDoc.DocCurrency = "$";
                                                            }

                                                            if (monedaDoc.Equals(monedaSistema)) // Si el documento es en Moneda Local
                                                                oDoc.Invoices.SumApplied = doc.Monto;
                                                            else
                                                                oDoc.Invoices.AppliedFC = doc.Monto;

                                                            if (doc.Monto >= 0)
                                                                oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                                                            else
                                                                oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;

                                                            oDoc.Invoices.Add();
                                                        }

                                                        /*SAPbouiCOM.ComboBox oMedioPago = oFormVisor.Items.Item("tjaDesc").Specific; // Tarjeta seleccionada
                                                        if (!String.IsNullOrEmpty(oMedioPago.Value.ToString()))
                                                            oDoc.CreditCards.CreditCard = 1;//Convert.ToInt32(oMedioPago.Selected.Value.ToString());*/
                                                        if (configAddOn.Empresa.Equals("ALMACEN"))
                                                        {
                                                            //Se cambia a una variable temporal, ya que la parametrización de Almacen es diferente al resto
                                                            string monedaDocTemp = "";
                                                            if (oDoc.DocCurrency.Equals("$")) monedaDocTemp = "UYU";
                                                            if (oDoc.DocCurrency.Equals("U$S")) monedaDocTemp = "USD";

                                                            if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))) //ASPL - 2020.08.18
                                                            {
                                                                oDoc.CashSum = montoTotalPago;

                                                                if (monedaDocTemp.Equals(monedaSistema))
                                                                    oDoc.CashAccount = "1113000001";
                                                                else
                                                                    oDoc.CashAccount = "1114000002";
                                                            }

                                                            // obtener id de tarjeta de la tabla OCRC
                                                            if (!String.IsNullOrEmpty(objetoLog.codigoTarjetaSAP))
                                                                oDoc.CreditCards.CreditCard = Convert.ToInt32(objetoLog.codigoTarjetaSAP);

                                                            if (!String.IsNullOrEmpty(objetoLog.cuentaTarjeta.ToString()))
                                                                oDoc.CreditCards.CreditAcct = objetoLog.cuentaTarjeta.ToString();
                                                            else
                                                            {
                                                                if (oDoc.DocCurrency.ToString().Equals(monedaStrISO) || oDoc.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                                                    oDoc.CreditCards.CreditAcct = configAddOn.TarjetaMN;
                                                                else
                                                                    oDoc.CreditCards.CreditAcct = configAddOn.TarjetaME;
                                                            }

                                                            String oNumeroTja = objetoLog.numerotarjeta; // Numero de tarjeta
                                                            oNumeroTja = oNumeroTja.Substring(oNumeroTja.Length - 4, 4);

                                                            if (!String.IsNullOrEmpty(oNumeroTja))
                                                            {
                                                                oDoc.CreditCards.CreditCardNumber = oNumeroTja;
                                                                if (oDoc.CreditCards.CreditCardNumber.Length > 4)
                                                                    oDoc.CreditCards.CreditCardNumber = oDoc.CreditCards.CreditCardNumber.Substring(0, 4);
                                                            }

                                                            string oFecha = objetoLog.fechaTransaccion.ToString(); // Fecha
                                                            if (!String.IsNullOrEmpty(oFecha))
                                                            {
                                                                string dueDateStr = "";
                                                                DateTime fechatransaccion = Convert.ToDateTime(oFecha);
                                                                dueDateStr = String.Format("{0}-{1}-{2}", fechatransaccion.Year, fechatransaccion.Month, fechatransaccion.Day);
                                                                oDoc.CreditCards.CardValidUntil = Convert.ToDateTime(dueDateStr);
                                                            }

                                                            String oCantCuotas = objetoLog.cuotas; // Cant de cuotas
                                                            if (!String.IsNullOrEmpty(oCantCuotas))
                                                                oDoc.CreditCards.NumOfPayments = Convert.ToInt32(oCantCuotas);

                                                            String oVoucherNro = objetoLog.ticket.ToString(); // Nro Certificado
                                                            if (!String.IsNullOrEmpty(oVoucherNro))
                                                                oDoc.CreditCards.VoucherNum = oVoucherNro;

                                                            String oOwnerId = objetoLog.selloCod; // Id Tja
                                                            if (!String.IsNullOrEmpty(oOwnerId))
                                                                oDoc.CreditCards.OwnerIdNum = oOwnerId;

                                                            // se comentan observaciones para pago por tarjeta
                                                            //SAPbouiCOM.EditText oObservaciones = oFormVisor.Items.Item("txObser").Specific; // Se agrega campo Observaciones 28/12/18
                                                            //if (!String.IsNullOrEmpty(oObservaciones.String))
                                                            //    oDoc.Remarks = oObservaciones.String; // Observaciones

                                                            //crear metodo para recibir el codigo
                                                            oDoc.CreditCards.CreditSum = montoTotalPago;
                                                            oDoc.CreditCards.NumOfCreditPayments = oDoc.CreditCards.NumOfPayments;

                                                            //Este codigo sale del metodo de pago creado en SAP, Tabla OCRP
                                                            oDoc.CreditCards.PaymentMethodCode = 3;//oDoc.CreditCards.NumOfPayments;
                                                            oDoc.CreditCards.ConfirmationNum = objetoLog.numerotarjeta;
                                                            //Se guarda en comentarios la moneda en la cual se realizo la transaccion
                                                            oDoc.Remarks = monedaParaPago;
                                                        }
                                                        // se agrega usuario logueado a documento de pago
                                                        oDoc.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;
                                                        //Numero de recibo
                                                        oDoc.CounterReference = numRecibo;
                                                        //Almacen Rural Referencia Pago 2% IVA
                                                        oDoc.UserFields.Fields.Item("U_PagoRef").Value = "No Aplica";

                                                        //else
                                                        //{
                                                        //    if (oDoc.DocCurrency.ToString().Equals(monedaSistema))
                                                        //        oDoc.CashAccount = "1113000001";
                                                        //    else
                                                        //        oDoc.CashAccount = "1114000002";
                                                        //}

                                                        if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                                                        {
                                                            lRetCode = oDoc.Add();
                                                            string ultimoDoc = "";
                                                            oCompany.GetNewObjectCode(out ultimoDoc);

                                                            if (lRetCode != 0)
                                                                SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ". Codigo error:" + lRetCode.ToString());
                                                            else
                                                            {
                                                                if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))) //ASPL - 2021.07.15 - Nuevo proceso por pago en diferente monedas.
                                                                {
                                                                    res = CrearAsiento(monedaDoc, monedaParaPago, dbCambio, montoTotalPago);

                                                                    if (res)
                                                                    {
                                                                        int res2 = 0;
                                                                        //************ CREACIÓN SEGUNDO DOCUMENTO PAGO A CUENTA ************************//
                                                                        SAPbobsCOM.Payments pagoCuenta = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                                                                        pagoCuenta.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                                                                        pagoCuenta.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
                                                                        pagoCuenta.DocDate = DateTime.Now;
                                                                        pagoCuenta.TransferDate = pDocumentos[0].Fecha;

                                                                        pagoCuenta.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
                                                                        pagoCuenta.DocCurrency = monedaParaPago;

                                                                        if (monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) //ASPL - 2021.07.19 - Nuevo proceso por pago en diferente monedas.
                                                                        {
                                                                            pagoCuenta.DocCurrency = "$";
                                                                            pagoCuenta.AccountPayments.AccountCode = "1113000001";
                                                                        }
                                                                        else if (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))
                                                                        {
                                                                            pagoCuenta.DocCurrency = "U$S";
                                                                            pagoCuenta.AccountPayments.AccountCode = "1114000002";
                                                                        }

                                                                        montoTotalPago = 0;

                                                                        foreach (clsPago doc in pDocumentos)
                                                                        {
                                                                            if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD")))
                                                                                doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100);

                                                                            montoTotalPago += doc.Monto;
                                                                        }

                                                                        //Almacen Rural Canje
                                                                        pagoCuenta.AccountPayments.SumPaid = montoTotalPago;
                                                                        //  pagoCuenta.AccountPayments.SumPaid = 10;
                                                                        pagoCuenta.AccountPayments.Decription = "Pago Tarjeta POS - " + pDocumentos[0].CardName;
                                                                        pagoCuenta.AccountPayments.Add();

                                                                        IssuerGeocom issuerAux = ObtenerDatosIssuerGeocom(respuesta.Issuer.ToString(), monedaParaPago, proveedor);
                                                                        if (!String.IsNullOrEmpty(issuerAux.cuentaContable.ToString()))
                                                                            pagoCuenta.CreditCards.CreditAcct = issuerAux.cuentaContable.ToString();
                                                                        else
                                                                        {
                                                                            if (pagoCuenta.DocCurrency.ToString().Equals(monedaStrISO) || pagoCuenta.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                                                                pagoCuenta.CreditCards.CreditAcct = configAddOn.TarjetaMN;
                                                                            else
                                                                                pagoCuenta.CreditCards.CreditAcct = configAddOn.TarjetaME;
                                                                        }

                                                                        // obtener id de tarjeta de la tabla OCRC
                                                                        if (!String.IsNullOrEmpty(objetoLog.codigoTarjetaSAP))
                                                                            pagoCuenta.CreditCards.CreditCard = Convert.ToInt32(objetoLog.codigoTarjetaSAP);

                                                                        string oNumeroTja = objetoLog.numerotarjeta; // Numero de tarjeta
                                                                        oNumeroTja = oNumeroTja.Substring(oNumeroTja.Length - 4, 4);

                                                                        if (!String.IsNullOrEmpty(oNumeroTja))
                                                                        {
                                                                            pagoCuenta.CreditCards.CreditCardNumber = oNumeroTja;

                                                                            if (pagoCuenta.CreditCards.CreditCardNumber.Length > 4)
                                                                                pagoCuenta.CreditCards.CreditCardNumber = pagoCuenta.CreditCards.CreditCardNumber.Substring(0, 4);
                                                                        }

                                                                        string oFecha = objetoLog.fechaTransaccion.ToString(); // Fecha

                                                                        if (!String.IsNullOrEmpty(oFecha))
                                                                        {
                                                                            string dueDateStr = "";
                                                                            DateTime fechatransaccion = Convert.ToDateTime(oFecha);
                                                                            dueDateStr = String.Format("{0}-{1}-{2}", fechatransaccion.Year, fechatransaccion.Month, fechatransaccion.Day);
                                                                            pagoCuenta.CreditCards.CardValidUntil = Convert.ToDateTime(dueDateStr);
                                                                        }

                                                                        string oCantCuotas = objetoLog.cuotas; // Cant de cuotas

                                                                        if (!String.IsNullOrEmpty(oCantCuotas))
                                                                            pagoCuenta.CreditCards.NumOfPayments = Convert.ToInt32(oCantCuotas);

                                                                        string oVoucherNro = objetoLog.ticket.ToString(); // Nro Certificado
                                                                        if (!String.IsNullOrEmpty(oVoucherNro))
                                                                            pagoCuenta.CreditCards.VoucherNum = oVoucherNro;

                                                                        string oOwnerId = objetoLog.selloCod; // Id Tja
                                                                        if (!String.IsNullOrEmpty(oOwnerId))
                                                                            pagoCuenta.CreditCards.OwnerIdNum = oOwnerId;

                                                                        //crear metodo para recibir el codigo
                                                                        pagoCuenta.CreditCards.CreditSum = montoTotalPago - taxAmountTemp;
                                                                        //pagoCuenta.CreditCards.CreditSum = 5;
                                                                        pagoCuenta.CreditCards.NumOfCreditPayments = pagoCuenta.CreditCards.NumOfCreditPayments;//oDoc.CreditCards.NumOfPayments;

                                                                        //Este codigo sale del metodo de pago creado en SAP, Tabla OCRP
                                                                        pagoCuenta.CreditCards.PaymentMethodCode = 3;//oDoc.CreditCards.NumOfPayments;
                                                                        pagoCuenta.CreditCards.ConfirmationNum = objetoLog.numerotarjeta;
                                                                        //Se guarda en comentarios la moneda en la cual se realizo la transaccion
                                                                        //En almacen Rural en este campo guardamos referencia del documento anterior
                                                                        pagoCuenta.Remarks = monedaParaPago;

                                                                        //Almacen Rural Referencia Pago 2% IVA
                                                                        pagoCuenta.UserFields.Fields.Item("U_PagoRef").Value = ultimoDoc;
                                                                        // se agrega usuario logueado a documento de pago
                                                                        pagoCuenta.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;
                                                                        //Numero de recibo
                                                                        pagoCuenta.CounterReference = numRecibo;

                                                                        if (!String.IsNullOrEmpty(respuesta.TaxAmount))
                                                                        {
                                                                            res2 = pagoCuenta.Add();

                                                                            if (res2 != 0)
                                                                            {
                                                                                objetoLog.EstatusSAPTransaccion = "Error";
                                                                                guardaLogGeocom(objetoLog, true);
                                                                                SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ".Error Documento Canje -  Codigo error:" + res2.ToString());
                                                                                res = false;
                                                                            }
                                                                            else
                                                                            {
                                                                                objetoLog.EstatusSAPTransaccion = "OK";
                                                                                res = true;
                                                                                guardaLogGeocom(objetoLog, true);
                                                                                //guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString());
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    res = true; // Pago ingresado correctamente 
                                                                    clienteSeleccionado = new clsCliente();
                                                                    int res2 = 0;
                                                                    //  oObservaciones.String = "";

                                                                    //************CREACIÓN SEGUNDO DOCUMENTO PAGO A CUENTA ************************//
                                                                    SAPbobsCOM.Payments pagoCuenta = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                                                                    pagoCuenta.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
                                                                    pagoCuenta.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
                                                                    pagoCuenta.DocDate = DateTime.Now;
                                                                    pagoCuenta.TransferDate = pDocumentos[0].Fecha;

                                                                    pagoCuenta.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
                                                                    // oDoc.DocCurrency = pDocumentos[0].Moneda;
                                                                    pagoCuenta.DocCurrency = monedaDoc;
                                                                    if (configAddOn.Empresa.Equals("ALMACEN"))
                                                                    {
                                                                        if (monedaDoc.Equals("USD"))
                                                                        {
                                                                            pagoCuenta.DocCurrency = "U$S";
                                                                            pagoCuenta.AccountPayments.AccountCode = "1114000002";

                                                                            if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))) //ASPL - 2021.07.15 - Nuevo proceso por pago en diferente monedas.
                                                                                pagoCuenta.AccountPayments.AccountCode = "1113000001";
                                                                        }
                                                                        else
                                                                        {
                                                                            pagoCuenta.DocCurrency = "$";
                                                                            pagoCuenta.AccountPayments.AccountCode = "1113000001";

                                                                            if ((monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU")) || (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))) //ASPL - 2021.07.15 - Nuevo proceso por pago en diferente monedas.
                                                                                pagoCuenta.AccountPayments.AccountCode = "1114000002";
                                                                        }
                                                                    }

                                                                    montoTotalPago = 0;

                                                                    //oDoc.BPLID = ObtenerSucursal(docEntry.ToString());
                                                                    // oDoc.BPLID = sucursalActiva;
                                                                    foreach (clsPago doc in pDocumentos)
                                                                    {
                                                                        if (monedaDoc.Equals("USD") && monedaParaPago.Equals("UYU"))
                                                                        {
                                                                            doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100) / montoTasaCambio; // Para corregir decimales
                                                                            //doc.Monto = (TextoaDecimal(respuesta.DatosTransaccion.Monto.ToString()) / 100) * montoTasaCambio; // Para corregir decimales
                                                                        }
                                                                        else if (monedaDoc.Equals("UYU") && monedaParaPago.Equals("USD"))
                                                                            doc.Monto = (TextoaDecimal(respuesta.TotalAmount.ToString()) / 100) * montoTasaCambio;
                                                                        else
                                                                            doc.Monto = TextoaDecimal(respuesta.TotalAmount.ToString()) / 100;

                                                                        montoTotalPago += doc.Monto;
                                                                    }

                                                                    //Almacen Rural Canje
                                                                    pagoCuenta.AccountPayments.SumPaid = montoTotalPago;
                                                                    //  pagoCuenta.AccountPayments.SumPaid = 10;
                                                                    pagoCuenta.AccountPayments.Decription = "Pago Tarjeta POS - " + pDocumentos[0].CardName;
                                                                    pagoCuenta.AccountPayments.Add();
                                                                    pagoCuenta.AccountPayments.AccountCode = "1133090000";
                                                                    pagoCuenta.AccountPayments.Decription = "Pago Tarjeta POS - " + pDocumentos[0].CardName;

                                                                    //si se cambio la moneda de pago convertir
                                                                    if (comboMoneda.Equals("Dolares"))
                                                                    {
                                                                        if (monedaDoc.Equals("UYU"))
                                                                            taxAmountTemp = Math.Round(taxAmountTemp * montoTasaCambio, 2);
                                                                    }
                                                                    else if (comboMoneda.Equals("Pesos"))
                                                                    {
                                                                        if (monedaDoc.Equals("USD"))
                                                                            taxAmountTemp = Math.Round(taxAmountTemp / montoTasaCambio, 2);
                                                                    }

                                                                    pagoCuenta.AccountPayments.SumPaid = taxAmountTemp * -1;
                                                                    //pagoCuenta.AccountPayments.SumPaid = -5;
                                                                    pagoCuenta.AccountPayments.Add();

                                                                    if (!String.IsNullOrEmpty(objetoLog.cuentaTarjeta.ToString()))
                                                                        pagoCuenta.CreditCards.CreditAcct = objetoLog.cuentaTarjeta.ToString();
                                                                    else
                                                                    {
                                                                        if (pagoCuenta.DocCurrency.ToString().Equals(monedaStrISO) || pagoCuenta.DocCurrency.ToString().Equals(monedaStrSimbolo))
                                                                            pagoCuenta.CreditCards.CreditAcct = configAddOn.TarjetaMN;
                                                                        else
                                                                            pagoCuenta.CreditCards.CreditAcct = configAddOn.TarjetaME;
                                                                    }

                                                                    // obtener id de tarjeta de la tabla OCRC
                                                                    if (!String.IsNullOrEmpty(objetoLog.codigoTarjetaSAP))
                                                                        pagoCuenta.CreditCards.CreditCard = Convert.ToInt32(objetoLog.codigoTarjetaSAP);

                                                                    String oNumeroTja = objetoLog.numerotarjeta; // Numero de tarjeta
                                                                    oNumeroTja = oNumeroTja.Substring(oNumeroTja.Length - 4, 4);
                                                                    if (!String.IsNullOrEmpty(oNumeroTja))
                                                                    {
                                                                        pagoCuenta.CreditCards.CreditCardNumber = oNumeroTja;

                                                                        if (pagoCuenta.CreditCards.CreditCardNumber.Length > 4)
                                                                            pagoCuenta.CreditCards.CreditCardNumber = pagoCuenta.CreditCards.CreditCardNumber.Substring(0, 4);
                                                                    }

                                                                    string oFecha = objetoLog.fechaTransaccion.ToString(); // Fecha

                                                                    if (!String.IsNullOrEmpty(oFecha))
                                                                    {
                                                                        string dueDateStr = "";
                                                                        DateTime fechatransaccion = Convert.ToDateTime(oFecha);
                                                                        dueDateStr = String.Format("{0}-{1}-{2}", fechatransaccion.Year, fechatransaccion.Month, fechatransaccion.Day);
                                                                        pagoCuenta.CreditCards.CardValidUntil = Convert.ToDateTime(dueDateStr);
                                                                    }

                                                                    String oCantCuotas = objetoLog.cuotas; // Cant de cuotas

                                                                    if (!String.IsNullOrEmpty(oCantCuotas))
                                                                        pagoCuenta.CreditCards.NumOfPayments = Convert.ToInt32(oCantCuotas);

                                                                    String oVoucherNro = objetoLog.ticket.ToString(); // Nro Certificado
                                                                    if (!String.IsNullOrEmpty(oVoucherNro))
                                                                        pagoCuenta.CreditCards.VoucherNum = oVoucherNro;

                                                                    String oOwnerId = objetoLog.selloCod; // Id Tja
                                                                    if (!String.IsNullOrEmpty(oOwnerId))
                                                                        pagoCuenta.CreditCards.OwnerIdNum = oOwnerId;

                                                                    //crear metodo para recibir el codigo
                                                                    pagoCuenta.CreditCards.CreditSum = montoTotalPago - taxAmountTemp;
                                                                    //pagoCuenta.CreditCards.CreditSum = 5;
                                                                    pagoCuenta.CreditCards.NumOfCreditPayments = pagoCuenta.CreditCards.NumOfCreditPayments;//oDoc.CreditCards.NumOfPayments;

                                                                    //Este codigo sale del metodo de pago creado en SAP, Tabla OCRP
                                                                    pagoCuenta.CreditCards.PaymentMethodCode = 3;//oDoc.CreditCards.NumOfPayments;
                                                                    pagoCuenta.CreditCards.ConfirmationNum = objetoLog.numerotarjeta;
                                                                    //Se guarda en comentarios la moneda en la cual se realizo la transaccion
                                                                    //En almacen Rural en este campo guardamos referencia del documento anterior
                                                                    pagoCuenta.Remarks = monedaParaPago;

                                                                    //Almacen Rural Referencia Pago 2% IVA
                                                                    pagoCuenta.UserFields.Fields.Item("U_PagoRef").Value = ultimoDoc;

                                                                    if (configAddOn.Empresa.Equals("ALMACEN"))
                                                                    {
                                                                        // se agrega usuario logueado a documento de pago
                                                                        pagoCuenta.UserFields.Fields.Item("U_Usuario").Value = usuarioPago;
                                                                        //Numero de recibo
                                                                        pagoCuenta.CounterReference = numRecibo;
                                                                    }

                                                                    if (!String.IsNullOrEmpty(respuesta.TaxAmount))
                                                                    {
                                                                        res2 = pagoCuenta.Add();

                                                                        if (res2 != 0)
                                                                        {
                                                                            objetoLog.EstatusSAPTransaccion = "Error";
                                                                            guardaLogGeocom(objetoLog, true);
                                                                            SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ".Error Documento Canje -  Codigo error:" + res2.ToString());
                                                                            res = false;
                                                                        }
                                                                        else
                                                                        {
                                                                            objetoLog.EstatusSAPTransaccion = "OK";
                                                                            res = true;
                                                                            guardaLogGeocom(objetoLog, true);
                                                                            //guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString());
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        // SBO_Application.MessageBox(oCompany.GetLastErrorDescription() + ".Error Documento Canje -  Codigo error:" + res2.ToString());
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        #endregion
                                                        //************FIN DOCUMENTO 1 PAGO EFECTIVO PARA CREAR REALIDAD DE CC*****************//
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    objetoLog.EstatusSAPTransaccion = "Error";
                                                    SBO_Application.MessageBox("Error al ingresar pago en SAP, cancele transacción desde POS.");
                                                    guardaLogGeocom(objetoLog, true);
                                                }
                                            }
                                            else
                                            {
                                                if (respuesta.PosResponseCode.Equals("0"))
                                                {
                                                    SBO_Application.MessageBox("Error al realizar operación en POS.");
                                                    objetoLog.EstatusGeocomTransaccion = "Transacción cancelada POS";
                                                }
                                                else
                                                {
                                                    SBO_Application.MessageBox("Tiempo Agotado (POS)- Error Transacciòn.");
                                                    objetoLog.EstatusGeocomTransaccion = codPosRespuesta.mensaje;
                                                }

                                                guardaLogGeocom(objetoLog, true);
                                            }
                                        }
                                        else if (respuesta.PosResponseCode.Equals("-1"))
                                        {
                                            SBO_Application.MessageBox("Error de comunicación con GEOCOM, la transacción se canceló automáticamente.");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                SBO_Application.MessageBox("Favor ingrese tasa de cambio en el sistema.");
                            }
                        }
                        else
                        {
                            SBO_Application.MessageBox("El RUT cargado en el cliente no es valido para realizar pago.");
                        }
                    }
                    else
                    {
                        SBO_Application.MessageBox("Folio no generado, espere unos instantes y vuelva a intentar.");
                    }
                }
                else
                {
                    SBO_Application.MessageBox("El monto a pagar no puede ser mayor al saldo pendiente");
                }
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al ingresarPagoTarjeta", ex.Message.ToString()); }
            return res;
        }

        public Boolean CrearAsiento(string pMonedaDoc, string pMonedaPago, double pMonto, double pMontoPago)
        {
            bool bResult = false;
            DateTime fechaHoy = DateTime.Now;
            SAPbobsCOM.JournalEntries oJounalEntry;
            oJounalEntry = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            int lRetCode = -1;
            string MonCtaDeb = "";
            string MonCtaCre = "";

            if (pMonedaDoc.Equals("USD") && pMonedaPago.Equals("UYU"))
            {
                MonCtaDeb = "1113000001";
                MonCtaCre = "1114000002";
            }
            if (pMonedaDoc.Equals("UYU") && pMonedaPago.Equals("USD"))
            {
                MonCtaDeb = "1114000002";
                MonCtaCre = "1113000001";
            }

            oJounalEntry.Series = 73;
            oJounalEntry.DueDate = fechaHoy;
            oJounalEntry.ReferenceDate = fechaHoy;
            oJounalEntry.TaxDate = fechaHoy;
            oJounalEntry.Lines.AccountCode = MonCtaDeb;
            if (pMonedaPago.Equals("USD"))
            {
                oJounalEntry.Lines.FCCurrency = "U$S";
                oJounalEntry.Lines.FCDebit = pMonto;
                oJounalEntry.Lines.Debit = pMontoPago;
            }
            else
            {
                oJounalEntry.Lines.FCCurrency = "U$S";
                oJounalEntry.Lines.FCDebit = pMontoPago;
                oJounalEntry.Lines.Debit = pMonto;
            }
            oJounalEntry.Lines.BPLID = 1;
            oJounalEntry.Lines.Add();
            oJounalEntry.Lines.ShortName = MonCtaCre;
            if (pMonedaPago.Equals("USD"))
            {
                oJounalEntry.Lines.FCCurrency = "U$S";
                oJounalEntry.Lines.FCCredit = pMonto;
                oJounalEntry.Lines.Credit = pMontoPago;
            }
            else
            {
                oJounalEntry.Lines.FCCurrency = "U$S";
                oJounalEntry.Lines.FCCredit = pMontoPago;
                oJounalEntry.Lines.Credit = pMonto;
            }
            oJounalEntry.Lines.BPLID = 1;
            //oJounalEntry.Lines.LineMemo = "Vale Nro. " + pNumVal;
            oJounalEntry.Lines.Add();
            lRetCode = oJounalEntry.Add();

            if (lRetCode != 0)
            {
                bResult = false;
                int temp_int = 0;
                string temp_string = string.Empty;
                oCompany.GetLastError(out temp_int, out temp_string);
                SBO_Application.MessageBox(" ERROR en Asiento : " + temp_string);
                throw new Exception(temp_string);
            }
            else
                bResult = true;

            return bResult;
        }

        private void LimpiarDatos()
        {
            try
            {
                SAPbouiCOM.ComboBox oComboCuotas = oFormVisor.Items.Item("cmbCT").Specific;
                oComboCuotas.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.EditText oNumero = oFormVisor.Items.Item("chNumero").Specific; // Numero de cheque
                oNumero.Value = "";

                SAPbouiCOM.EditText oFecha = oFormVisor.Items.Item("chVto").Specific; // Fecha
                oFecha.Value = "";

                SAPbouiCOM.EditText oFechaTran = oFormVisor.Items.Item("fcTran").Specific; // Fecha
                oFechaTran.Value = "";

                SAPbouiCOM.EditText oMontoEf = oFormVisor.Items.Item("efMonto").Specific;
                oMontoEf.Value = "";

                //Seleccion por defecto
                SAPbouiCOM.ComboBox oComboTransf = oFormVisor.Items.Item("transfCta").Specific;
                oComboTransf.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboProveedor = oFormVisor.Items.Item("cmbProv").Specific;
                oComboProveedor.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboSello = oFormVisor.Items.Item("cmbSell").Specific;
                oComboSello.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                SAPbouiCOM.ComboBox oComboEfec2 = oFormVisor.Items.Item("cmbEfec").Specific;
                oComboEfec2.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboBank = oFormVisor.Items.Item("chBanco").Specific;
                oComboBank.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboCta = oFormVisor.Items.Item("chCta").Specific;
                oComboCta.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboProv = oFormVisor.Items.Item("cmbProv").Specific;
                oComboProv.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                SAPbouiCOM.ComboBox oComboSell = oFormVisor.Items.Item("cmbSell").Specific;
                oComboSell.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception)
            {
            }
        }

        private Boolean DevolucionPagoPos(List<clsPago> pDocumentos)
        {
            Boolean res = false;

            int docEntry = 0;
            Boolean tipoTarjeta = true;
            string terminal = "";
            string cardCode = "";
            string cardName = "";
            string moneda = "";
            string operacion = "";
            DateTime fecha;
            double montoGravado = 0;
            string impuesto = "";
            string ticket;
            string tranId;
            int factura;
            string monedaDePago = "";
            string tasaCambio = "";
            string monedaTransaccion = "";
            int digitoVerificadorRespuesta;
            string rutAvalidar = "";
            int digitoAvalidar = -1;
            bool rutValidado = false;
            bool clienteFinal = true;
            int decretoLey = 1;
            string refeenciaPagoCanje = "";
            double impuestoCanjeCuentaIVA = 0;
            CultureInfo culture = new CultureInfo("en-US");

            try
            {
                if (pDocumentos.Count != 0)
                {
                    double montoTotalPago = 0;

                    int lRetCode;
                    string montoTotal = String.Empty;
                    string rut = String.Empty;

                    cardCode = pDocumentos[0].CardCode;
                    cardName = pDocumentos[0].CardName;
                    fecha = pDocumentos[0].Fecha;
                    moneda = pDocumentos[0].Moneda;
                    docEntry = pDocumentos[0].DocEntry;
                    ticket = pDocumentos[0].Ticket;
                    tranId = pDocumentos[0].TranId;
                    factura = pDocumentos[0].Factura;
                    monedaDePago = pDocumentos[0].MonedaPago;
                    tasaCambio = pDocumentos[0].TasaCambio;

                    SAPbobsCOM.Recordset oRSMyTable = null;
                    String query = "";

                    try
                    {
                        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                        {
                            //query += "select VatSum as Impuesto, LicTradNum from OINV where DocEntry = '" + docEntry + "'";
                            query += "SELECT T2.VatSum  as IMPUESTO, T2.LicTradNum, T2.DocTotal as TOTAL FROM ORCT T0 "
                                    + "LEFT JOIN RCT2 T1 ON T1.DocNum = T0.DocEntry "
                                    + "LEFT JOIN OINV T2 ON T2.DocEntry = T1.DocEntry "
                                    + "where T0.DocEntry = '" + docEntry + "'";
                        }
                        else
                        {
                            //Almacen Rural 
                            if (moneda.Equals("U$S") || moneda.Equals("USD"))
                            {
                                query += "SELECT T2.\"VatSumFC\"  as IMPUESTO, T2.\"LicTradNum\",T2.\"DocTotalFC\" as TOTAL, T0.\"U_PagoRef\" FROM \"ORCT\" T0 "
                               + "LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\" = T0.\"DocEntry\" "
                               + "LEFT JOIN \"OINV\" T2 ON T2.\"DocEntry\" = T1.\"DocEntry\" "
                               + "where T0.\"DocEntry\" = '" + docEntry + "'";

                            }
                            else
                            {
                                query += "SELECT T2.\"VatSum\"  as IMPUESTO, T2.\"LicTradNum\",T2.\"DocTotal\" as TOTAL, T0.\"U_PagoRef\" FROM \"ORCT\" T0 "
                                  + "LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\" = T0.\"DocEntry\" "
                                  + "LEFT JOIN \"OINV\" T2 ON T2.\"DocEntry\" = T1.\"DocEntry\" "
                                  + "where T0.\"DocEntry\" = '" + docEntry + "'";

                            }

                        }

                        oRSMyTable.DoQuery(query);
                        impuesto = oRSMyTable.Fields.Item("IMPUESTO").Value.ToString();
                        montoTotal = oRSMyTable.Fields.Item("TOTAL").Value.ToString();
                        docEntry = pDocumentos[0].DocEntry;
                        refeenciaPagoCanje = oRSMyTable.Fields.Item("U_PagoRef").Value.ToString();

                        //Validar si tiene referencia con documento de canje
                        if (!String.IsNullOrEmpty(refeenciaPagoCanje) && !refeenciaPagoCanje.Equals("No Aplica"))
                        {
                            query = "";
                            //Almacen Rural 
                            if (moneda.Equals("U$S") || moneda.Equals("USD"))
                            {
                                query += "SELECT T2.\"VatSumFC\"  as IMPUESTO, T2.\"LicTradNum\",T2.\"DocTotalFC\" as TOTAL, T0.\"U_PagoRef\" FROM \"ORCT\" T0 "
                               + "LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\" = T0.\"DocEntry\" "
                               + "LEFT JOIN \"OINV\" T2 ON T2.\"DocEntry\" = T1.\"DocEntry\" "
                               + "where T0.\"DocEntry\" = '" + refeenciaPagoCanje + "'";

                            }
                            else
                            {
                                query += "SELECT T2.\"VatSum\"  as IMPUESTO, T2.\"LicTradNum\",T2.\"DocTotal\" as TOTAL, T0.\"U_PagoRef\" FROM \"ORCT\" T0 "
                                  + "LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\" = T0.\"DocEntry\" "
                                  + "LEFT JOIN \"OINV\" T2 ON T2.\"DocEntry\" = T1.\"DocEntry\" "
                                  + "where T0.\"DocEntry\" = '" + refeenciaPagoCanje + "'";

                            }
                            oRSMyTable.DoQuery(query);
                            impuesto = oRSMyTable.Fields.Item("IMPUESTO").Value.ToString();
                            montoTotal = oRSMyTable.Fields.Item("TOTAL").Value.ToString();
                            docEntry = pDocumentos[0].DocEntry;

                            //Obtener impuesto impuestoCanjeCuentaIVA
                            string query2 = "";
                            if (moneda.Equals("U$S") || moneda.Equals("USD"))
                            {
                                query2 = "SELECT T1.\"AppliedFC\" AS IVACanje FROM ORCT T0  INNER JOIN RCT4 T1 ON T0.\"DocEntry\" = T1.\"DocNum\" WHERE T0.\"DocEntry\" =" + docEntry + " and  T1.\"AcctCode\" =  1133090000";
                            }
                            else
                            {
                                query2 = "SELECT T1.\"AppliedSys\"  AS IVACanje FROM ORCT T0  INNER JOIN RCT4 T1 ON T0.\"DocEntry\" = T1.\"DocNum\" WHERE T0.\"DocEntry\" =" + docEntry + " and  T1.\"AcctCode\" =  1133090000";
                            }
                            oRSMyTable.DoQuery(query2);
                            string IVACanjeTemp = oRSMyTable.Fields.Item("IVACanje").Value.ToString();

                            if (!String.IsNullOrEmpty(IVACanjeTemp))
                            {
                                impuestoCanjeCuentaIVA = double.Parse(IVACanjeTemp);
                                impuestoCanjeCuentaIVA = impuestoCanjeCuentaIVA * -1;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                    }

                    foreach (clsPago doc in pDocumentos)
                    {
                        double monto, tc;
                        string monedaTemp = "";
                        string montoTemp = "";
                        //Almacen Rural
                        if (moneda.Equals("U$S"))
                        {
                            monedaTemp = "USD";
                        }
                        else
                            monedaTemp = "UYU";

                        if ((moneda.Equals("USD") || moneda.Equals("U$S")) && monedaDePago.Equals("UYU"))
                        {
                            if (doc.Factura == 0)
                            {
                                monto = doc.Monto;
                                string tcTemp = ObtenerCambioAlmacenFecha(monedaDePago, fecha.ToString(configAddOn.FormatoFecha));
                                tc = Convert.ToDouble(tcTemp);
                                doc.Monto = monto * tc; // Para corregir decimales
                            }
                            else
                            {
                                monto = doc.Monto;
                                tc = double.Parse(doc.TasaCambio.ToString(), culture);
                                doc.Monto = monto * tc; // Para corregir decimales
                            }
                            /*
                            monto = doc.Monto;
                            
                            tc = double.Parse(doc.TasaCambio.ToString(), culture);
                         
                            doc.Monto = monto * tc; // Para corregir decimales
                            */
                            //Almacen Rural
                            if (!String.IsNullOrEmpty(refeenciaPagoCanje) && !refeenciaPagoCanje.Equals("No Aplica"))
                            {
                                montoTotalPago += doc.Monto + impuestoCanjeCuentaIVA;
                            }
                            else
                            {
                                montoTotalPago += doc.Monto;
                            }

                            //montoGravado = montoTotalPago - (Convert.ToDouble(impuesto)) * tc;

                            ///////////////////////////////////////////////////
                            ////MODIFICACIONES TN
                            //////////////////////////////////////////////////

                            double montoFactura = double.Parse(montoTotal, culture) * tc;
                            double impuestoFactura = Convert.ToDouble(impuesto) * tc;

                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuestoFactura)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuestoFactura) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();
                            ///////////////////////////////////////////////////////////////
                            monedaTransaccion = monedaDePago;

                        }
                        else if ((moneda.Equals("USD") || moneda.Equals("U$S")) && monedaDePago.Equals("USD"))
                        {
                            monto = doc.Monto;

                            //Almacen Rural
                            if (!String.IsNullOrEmpty(refeenciaPagoCanje) && !refeenciaPagoCanje.Equals("No Aplica"))
                            {
                                montoTotalPago += doc.Monto + impuestoCanjeCuentaIVA;
                            }
                            else
                            {
                                montoTotalPago += doc.Monto;
                            }

                            // montoGravado = montoTotalPago - (Convert.ToDouble(impuesto)) / tc;
                            double montoFactura = Convert.ToDouble(montoTotal);
                            double impuestoFactura = Convert.ToDouble(impuesto);
                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuesto)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuesto) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();

                            ///////////////////////////////////////////////////////////////

                            monedaTransaccion = monedaDePago;
                        }
                        else if ((moneda.Equals("$") || moneda.Equals("UYU")) && monedaDePago.Equals("USD"))
                        {
                            if (doc.Factura == 0)
                            {
                                monto = getDouble(doc.Monto.ToString());
                                string tcTemp = ObtenerCambioAlmacenFecha(monedaDePago, fecha.ToString(configAddOn.FormatoFecha));
                                tc = Convert.ToDouble(tcTemp);
                                doc.Monto = monto / tc; // Para corregir decimales*/
                            }
                            else
                            {
                                monto = getDouble(doc.Monto.ToString());
                                tc = getDouble(doc.TasaCambio.ToString());
                                doc.Monto = monto / tc; // Para corregir decimales
                            }

                            /*   monto = getDouble(doc.Monto.ToString());
                               tc = getDouble(doc.TasaCambio.ToString());
                               doc.Monto = monto / tc; // Para corregir decimales*/

                            //Almacen Rural
                            if (!String.IsNullOrEmpty(refeenciaPagoCanje) && !refeenciaPagoCanje.Equals("No Aplica"))
                            {
                                montoTotalPago += doc.Monto + impuestoCanjeCuentaIVA;
                            }
                            else
                            {
                                montoTotalPago += doc.Monto;
                            }
                            //montoGravado = montoTotalPago - (Convert.ToDouble(impuesto)) * tc;

                            ///////////////////////////////////////////////////
                            ////MODIFICACIONES TN
                            //////////////////////////////////////////////////

                            double montoFactura = Convert.ToDouble(montoTotal) / tc;
                            double impuestoFactura = Convert.ToDouble(impuesto) / tc;

                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuesto)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuesto) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();
                            ///////////////////////////////////////////////////////////////
                            monedaTransaccion = monedaDePago;
                        }
                        else
                        {
                            doc.Monto = getDouble(doc.Monto.ToString()); // Para corregir decimales

                            //Almacen Rural
                            if (!String.IsNullOrEmpty(refeenciaPagoCanje) && !refeenciaPagoCanje.Equals("No Aplica"))
                            {
                                montoTotalPago += doc.Monto + impuestoCanjeCuentaIVA;
                            }
                            else
                            {
                                montoTotalPago += doc.Monto;
                            }

                            //montoTotalPago += doc.Monto;
                            // montoGravado = montoTotalPago - Convert.ToDouble(impuesto);
                            //Calculamos el IVA segun el monto de la factura

                            ///////////////////////////////////////////////////
                            ////MODIFICACIONES TN
                            //////////////////////////////////////////////////

                            double montoFactura = Convert.ToDouble(montoTotal);
                            double impuestoFactura = Convert.ToDouble(impuesto);

                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuesto)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuesto) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();
                            ///////////////////////////////////////////////////////////////

                            monedaTransaccion = moneda;
                        }

                    }

                    //cmdDevT
                    SAPbouiCOM.ComboBox oTerminal = oFormVisor.Items.Item("cmbT1").Specific; //
                    terminal = oTerminal.Selected.Description.ToString();

                    operacion = "DEV";

                    //Mensaje para insertar tarjeta
                    if (proveedorPOS.Equals("GEOCOM"))
                    {
                        SBO_Application.MessageBox("Ingrese tarjeta en POS para realizar cancelación.");

                        ControladorGeocom cg = new ControladorGeocom(this);
                        LogGeocom objetoLog = new LogGeocom();
                        LogGeocom objetoLogVenta = new LogGeocom();
                        //datos de test
                        string PosID = terminal;
                        string systemId = configAddOn.hash;
                        string Branch = "Almacen";
                        string clientAppId = "1";
                        string userId = "1";
                        string ticketTemp = "";

                        objetoLogVenta = ObtenerDatosTicketGeocom(ticket.ToString(), PosID);

                        GeocomWSProductivo.PurchaseQueryResponse respuesta = cg.devolucion(montoTotalPago, PosID, monedaTransaccion, montoGravado, Convert.ToDouble(impuesto), factura.ToString(), 0, decretoLey, montoTotalPago, systemId, Branch, clientAppId, userId, ticket.ToString(), objetoLogVenta);

                        if (!respuesta.PosResponseCode.Equals("-1"))
                        {
                            //buscar en las tablas el codigo de respuesta

                            CodRespuestaPOSGeocom codPosRespuesta = CodigoPOSGeocom(respuesta.PosResponseCode);

                            if (String.IsNullOrEmpty(respuesta.OriginCardType)) objetoLog.cardtype = "-"; else objetoLog.cardtype = respuesta.OriginCardType;
                            if (String.IsNullOrEmpty(respuesta.Ci)) objetoLog.ci = "-"; else objetoLog.ci = respuesta.Ci;
                            if (String.IsNullOrEmpty(respuesta.AuthorizationCode)) objetoLog.codigoAutorizacion = "-"; else objetoLog.codigoAutorizacion = respuesta.AuthorizationCode;
                            if (String.IsNullOrEmpty(codPosRespuesta.codigo)) objetoLog.codigoRespuestaPos = "-"; else objetoLog.codigoRespuestaPos = codPosRespuesta.codigo;
                            if (String.IsNullOrEmpty(codPosRespuesta.estado)) objetoLog.codigoRespuestaPosDescripcion = "-"; else objetoLog.codigoRespuestaPosDescripcion = codPosRespuesta.estado;
                            if (String.IsNullOrEmpty(respuesta.Quota)) objetoLog.cuotas = "-"; else objetoLog.cuotas = respuesta.Quota;
                            // objetoLog.fechaTransaccion = respuesta.TransactionDate;
                            objetoLog.transactionDateTime = respuesta.TransactionHour;
                            if (String.IsNullOrEmpty(respuesta.TaxRefund)) objetoLog.impuestocodigo = "-"; else objetoLog.impuestocodigo = respuesta.TaxRefund;
                            if (String.IsNullOrEmpty(respuesta.Batch)) objetoLog.lote = "-"; else objetoLog.lote = respuesta.Batch;
                            if (String.IsNullOrEmpty(respuesta.Currency)) objetoLog.monedaTransaccionCod = "-"; else objetoLog.monedaTransaccionCod = respuesta.Currency;

                            if (objetoLog.monedaTransaccionCod.Equals("858"))
                            {
                                objetoLog.monedaTransaccionDescrip = "UYU";
                            }
                            else if (objetoLog.monedaTransaccionCod.Equals("840"))
                            {
                                //moneda = "0840"; //dolares
                                objetoLog.monedaTransaccionDescrip = "USD";
                            }
                            else
                            {
                                objetoLog.monedaTransaccionDescrip = "-";
                            }

                            //En esta seccion se buscara en la tabla de tarjetas y se traera el Issuer y su cuenta
                            IssuerGeocom issuerTemp = ObtenerDatosIssuerGeocom(respuesta.Issuer.ToString(), monedaTransaccion, objetoLogVenta.Merchant);
                            objetoLog.issuerCode = respuesta.Issuer.ToString();
                            if (String.IsNullOrEmpty(objetoLog.issuerCode)) objetoLog.issuerCode = "-";
                            if (String.IsNullOrEmpty(issuerTemp.nombreTarjeta)) objetoLog.issuerCodeDescripcion = "-"; else objetoLog.issuerCodeDescripcion = issuerTemp.nombreTarjeta;
                            // objetoLog.issuerCodeDescripcion = issuerTemp.nombreTarjeta;//ir a buscar en tabla nativa tarjetas
                            if (String.IsNullOrEmpty(issuerTemp.cuentaContable)) objetoLog.cuentaTarjeta = "-"; else objetoLog.cuentaTarjeta = issuerTemp.cuentaContable;
                            //objetoLog.cuentaTarjeta = issuerTemp.cuentaContable;
                            if (String.IsNullOrEmpty(issuerTemp.codigoTarjetaSAP)) objetoLog.codigoTarjetaSAP = "-"; else objetoLog.codigoTarjetaSAP = issuerTemp.codigoTarjetaSAP;
                            // objetoLog.codigoTarjetaSAP = issuerTemp.codigoTarjetaSAP;

                            //--------------------------------------------------------------------------------------------------//

                            if (codPosRespuesta.estado.Equals("OK"))
                            {
                                objetoLog.EstatusGeocomTransaccion = "OK";

                                double montoTemp = Convert.ToDouble(respuesta.TotalAmount) / 100;
                                objetoLog.monto = montoTemp.ToString();

                                if (String.IsNullOrEmpty(respuesta.CardOwnerName)) objetoLog.nombre = "-"; else objetoLog.nombre = respuesta.CardOwnerName;
                                if (String.IsNullOrEmpty(respuesta.EmvApplicationName)) objetoLog.nombreTarjeta = "-"; else objetoLog.nombreTarjeta = respuesta.EmvApplicationName;
                                if (String.IsNullOrEmpty(respuesta.CardNumber)) objetoLog.numerotarjeta = "-"; else objetoLog.numerotarjeta = respuesta.CardNumber;
                                if (String.IsNullOrEmpty(respuesta.Plan)) objetoLog.plan = "-"; else objetoLog.plan = respuesta.Plan;
                                if (String.IsNullOrEmpty(respuesta.PosID)) objetoLog.posId = "-"; else objetoLog.posId = respuesta.PosID;
                                if (String.IsNullOrEmpty(respuesta.AcquirerTerminal)) objetoLog.terminal = "-"; else objetoLog.terminal = respuesta.AcquirerTerminal;
                                if (String.IsNullOrEmpty(respuesta.Ticket)) objetoLog.ticket = ""; else objetoLog.ticket = respuesta.Ticket;

                                // ir a buscar en tabla SellosGeocom
                                objetoLog.selloCod = respuesta.Acquirer.ToString();
                                objetoLog.selloDescripcion = ObtenerSelloGeocom(objetoLog.selloCod);
                                if (String.IsNullOrEmpty(objetoLog.selloDescripcion)) objetoLog.selloDescripcion = "-";

                                //se va a buscar a la tabla TipoTranGeocom
                                if (String.IsNullOrEmpty(respuesta.TransactionType)) objetoLog.transaccionType = "-"; else objetoLog.transaccionType = respuesta.TransactionType;
                                objetoLog.transaccionTypeDescripcion = TipoTransaccionGeocom(objetoLog.transaccionType);
                                if (String.IsNullOrEmpty(objetoLog.transaccionTypeDescripcion)) objetoLog.transaccionTypeDescripcion = "-";

                                SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                                SAPbobsCOM.Payments oDocCanje = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago

                                string testcardcode = pDocumentos[0].CardCode;
                                /*
                                oDoc.CardName = pDocumentos[0].CardName;
                                oDoc.TransferDate = pDocumentos[0].Fecha;
                                oDoc.DocCurrency = pDocumentos[0].Moneda;
                                montoTotalPago = 0;
                                string query2 = "";
                                SAPbobsCOM.Recordset oRSMyTable2 = null;
                                */

                                oDoc.GetByKey(pDocumentos[0].DocEntry);
                                string testcardname = oDoc.CardName;
                                lRetCode = oDoc.Cancel();

                                if (lRetCode != 0)
                                {
                                    objetoLog.EstatusSAPTransaccion = "Error";
                                    guardaLogGeocom(objetoLog, false);
                                    string error = oCompany.GetLastErrorDescription();
                                    int codeError = oCompany.GetLastErrorCode();
                                    SBO_Application.MessageBox("ERROR." + error.ToString());
                                }
                                else
                                {
                                    res = true;
                                    objetoLog.EstatusSAPTransaccion = "OK";
                                    if (!refeenciaPagoCanje.Equals("No Aplica") || !String.IsNullOrEmpty(refeenciaPagoCanje))
                                    {
                                        oDocCanje.GetByKey(Convert.ToInt32(refeenciaPagoCanje));
                                        lRetCode = oDocCanje.Cancel();
                                        if (lRetCode != 0)
                                        {
                                            SBO_Application.MessageBox("Error cancelar documento de canje." + refeenciaPagoCanje + " favor realizar cancelación manual.");
                                        }

                                    }

                                    guardaLogGeocom(objetoLog, false);
                                    //guardarLogRappelLineas(pLineasNC);
                                    //SBO_Application.MessageBox("Pago cancelado correctamente");
                                }
                            }
                            else
                            {
                                if (respuesta.PosResponseCode.Equals("CT"))
                                {
                                    SBO_Application.MessageBox("Error al realizar operación en POS.");
                                    objetoLog.EstatusGeocomTransaccion = "Transacción cancelada POS";
                                }
                                else
                                {
                                    SBO_Application.MessageBox(codPosRespuesta.mensaje);
                                    objetoLog.EstatusGeocomTransaccion = codPosRespuesta.mensaje;
                                }

                                guardaLogGeocom(objetoLog, true);
                            }
                        }
                        else if (respuesta.PosResponseCode.Equals("-1"))
                        {
                            SBO_Application.MessageBox("Error de comunicación con GEOCOM, la transacción se canceló automáticamente.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al cancelar pago", ex.Message.ToString());
                // guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaTitular.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaNombre.ToString(), respuesta.TarjetaTipo.ToString());

            }
            return res;
        }

        private Boolean DevolucionPagoDocumentosErrorSAP(string documento)
        {
            Boolean res = false;

            int docEntry = 0;
            Boolean tipoTarjeta = true;
            string terminal = "";
            string cardCode = "";
            string cardName = "";
            string moneda = "";
            string operacion = "";
            DateTime fecha;
            double montoGravado = 0;
            string impuesto = "";
            int ticket = 0;
            string tranId;
            int factura = 0;
            string monedaDePago = "";
            string tasaCambio = "";
            string monedaTransaccion = "";
            int digitoVerificadorRespuesta;
            string rutAvalidar = "";
            int digitoAvalidar = -1;
            bool rutValidado = false;
            bool clienteFinal = true;
            int decretoLey = 1;
            bool retornoUpdate = false;
            CultureInfo culture = new CultureInfo("en-US");

            try
            {
                if (!String.IsNullOrEmpty(documento))
                {
                    double montoTotalPago = 0;
                    int lRetCode;
                    string montoTotal = String.Empty;
                    string rut = String.Empty;
                    SAPbobsCOM.Recordset oRSMyTable = null;
                    String query = "";

                    try
                    {
                        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                        {
                            //Almacen Rural 
                            //Se Buscan datos en Log Geocom
                            query += "SELECT TOP 1000 * FROM \"@LOGGEOCOM\" WHERE \"DocEntry\" = " + documento;
                        }

                        oRSMyTable.DoQuery(query);

                        //terminal = oRSMyTable.Fields.Item("U_Terminal").Value.ToString();
                        string tempMontoPago = oRSMyTable.Fields.Item("U_monto").Value.ToString();
                        montoTotalPago = Convert.ToDouble(tempMontoPago);

                        monedaTransaccion = oRSMyTable.Fields.Item("U_MonedaTransaccionCod").Value.ToString();
                        string tempdecretoLey = oRSMyTable.Fields.Item("U_impuestocodigo").Value.ToString();
                        decretoLey = Convert.ToInt32(tempdecretoLey.Substring(0, 1));
                        ticket = Convert.ToInt32(oRSMyTable.Fields.Item("U_ticket").Value.ToString());

                        SAPbouiCOM.ComboBox oTerminal = oFormVisor.Items.Item("term").Specific; //
                        terminal = oTerminal.Selected.Description.ToString();
                    }
                    catch (Exception e)
                    {
                    }

                    //Mensaje para insertar tarjeta
                    if (proveedorPOS.Equals("GEOCOM"))
                    {
                        SBO_Application.MessageBox("Ingrese tarjeta en POS para realizar cancelación.");

                        ControladorGeocom cg = new ControladorGeocom(this);
                        LogGeocom objetoLog = new LogGeocom();
                        LogGeocom objetoLogVenta = new LogGeocom();
                        //datos de test
                        string PosID = terminal;
                        string systemId = configAddOn.hash;
                        string Branch = "Almacen";
                        string clientAppId = "1";
                        string userId = "1";
                        string ticketTemp = "";

                        objetoLogVenta = ObtenerDatosTicketGeocom(ticket.ToString(), PosID);

                        GeocomWSProductivo.PurchaseQueryResponse respuesta = cg.devolucion(montoTotalPago, PosID, monedaTransaccion, montoGravado, Convert.ToDouble(0), factura.ToString(), 0, decretoLey, montoTotalPago, systemId, Branch, clientAppId, userId, ticket.ToString(), objetoLogVenta);

                        if (!respuesta.PosResponseCode.Equals("-1"))
                        {
                            //buscar en las tablas el codigo de respuesta
                            CodRespuestaPOSGeocom codPosRespuesta = CodigoPOSGeocom(respuesta.PosResponseCode);

                            if (String.IsNullOrEmpty(respuesta.OriginCardType)) objetoLog.cardtype = "-"; else objetoLog.cardtype = respuesta.OriginCardType;

                            if (codPosRespuesta.estado.Equals("OK"))
                            {
                                retornoUpdate = updateLogGeocomAnulacion(Convert.ToInt32(documento));
                                if (retornoUpdate)
                                {
                                    //20200715
                                    DateTime fechaTemp = DateTime.Now;
                                    string dia = "";
                                    string mes = "";

                                    if (fechaTemp.Day < 10)
                                    {
                                        dia = "0" + fechaTemp.Day.ToString();
                                    }
                                    else
                                    {
                                        dia = fechaTemp.Day.ToString();
                                    }
                                    if (fechaTemp.Month < 10)
                                    {
                                        mes = "0" + fechaTemp.Month.ToString();
                                    }
                                    else
                                    {
                                        mes = fechaTemp.Month.ToString();
                                    }

                                    string fechaFiltrado = fechaTemp.Year.ToString() + mes + dia;
                                    cargarGrilla(fechaFiltrado, fechaFiltrado);
                                    SBO_Application.MessageBox("Pago anulado con èxito.");
                                }
                                else
                                {
                                    cargarFormularioError();
                                    SBO_Application.MessageBox("Error al anular operación en POS - Contacte a soporte Tècnico.");
                                }
                            }
                            else
                            {
                                if (respuesta.PosResponseCode.Equals("CT"))
                                {
                                    SBO_Application.MessageBox("Error al realizar operación en POS.");
                                    objetoLog.EstatusGeocomTransaccion = "Transacción cancelada POS";
                                }
                                else
                                {
                                    SBO_Application.MessageBox(codPosRespuesta.mensaje);
                                    objetoLog.EstatusGeocomTransaccion = codPosRespuesta.mensaje;
                                }

                                guardaLogGeocom(objetoLog, true);
                            }
                        }
                        else if (respuesta.PosResponseCode.Equals("-1"))
                        {
                            SBO_Application.MessageBox("Error de comunicación con GEOCOM, la transacción se canceló automáticamente.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al cancelar pago", ex.Message.ToString());
                // guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaTitular.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaNombre.ToString(), respuesta.TarjetaTipo.ToString());

            }
            return res;
        }

        private Boolean cancelarPagoPos(List<clsPago> pDocumentos)
        {
            Boolean res = false;

            int docEntry = 0;
            Boolean tipoTarjeta = true;
            string terminal = "";
            string cardCode = "";
            string cardName = "";
            string moneda = "";
            string operacion = "";
            DateTime fecha;
            double montoGravado = 0;
            string impuesto = "";
            string ticket;
            string tranId;
            int factura;
            string monedaDePago = "";
            string tasaCambio = "";
            string monedaTransaccion = "";
            int digitoVerificadorRespuesta;
            string rutAvalidar = "";
            int digitoAvalidar = -1;
            bool rutValidado = false;
            bool clienteFinal = true;
            int decretoLey = 1;
            string montoTotal = String.Empty;
            string rut = String.Empty;
            string refeenciaPagoCanje = "";
            CultureInfo culture = new CultureInfo("en-US");

            try
            {
                if (pDocumentos.Count != 0)
                {
                    double montoTotalPago = 0;
                    int lRetCode;
                    cardCode = pDocumentos[0].CardCode;
                    cardName = pDocumentos[0].CardName;
                    fecha = pDocumentos[0].Fecha;
                    moneda = pDocumentos[0].Moneda;
                    docEntry = pDocumentos[0].DocEntry;
                    ticket = pDocumentos[0].Ticket;
                    tranId = pDocumentos[0].TranId;
                    factura = pDocumentos[0].Factura;
                    monedaDePago = pDocumentos[0].MonedaPago;
                    tasaCambio = pDocumentos[0].TasaCambio;

                    SAPbobsCOM.Recordset oRSMyTable = null;
                    String query = "";

                    try
                    {
                        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                        {
                            //query += "select VatSum as Impuesto, LicTradNum from OINV where DocEntry = '" + docEntry + "'";
                            query += "SELECT T2.VatSum  as IMPUESTO, T2.LicTradNum, T2.DocTotal as TOTAL FROM ORCT T0 "
                                    + "LEFT JOIN RCT2 T1 ON T1.DocNum = T0.DocEntry "
                                    + "LEFT JOIN OINV T2 ON T2.DocEntry = T1.DocEntry "
                                    + "where T0.DocEntry = '" + docEntry + "'";
                        }
                        else
                        {
                            //Almacen Rural 
                            if (moneda.Equals("U$S") || moneda.Equals("USD"))
                            {
                                query += "SELECT T2.\"VatSumFC\"  as IMPUESTO, T2.\"LicTradNum\",T2.\"DocTotalFC\" as TOTAL, T0.\"U_PagoRef\" FROM \"ORCT\" T0 "
                               + "LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\" = T0.\"DocEntry\" "
                               + "LEFT JOIN \"OINV\" T2 ON T2.\"DocEntry\" = T1.\"DocEntry\" "
                               + "where T0.\"DocEntry\" = '" + docEntry + "'";

                            }
                            else
                            {
                                query += "SELECT T2.\"VatSum\"  as IMPUESTO, T2.\"LicTradNum\",T2.\"DocTotal\" as TOTAL , T0.\"U_PagoRef\" FROM \"ORCT\" T0 "
                                  + "LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\" = T0.\"DocEntry\" "
                                  + "LEFT JOIN \"OINV\" T2 ON T2.\"DocEntry\" = T1.\"DocEntry\" "
                                  + "where T0.\"DocEntry\" = '" + docEntry + "'";

                            }

                        }

                        oRSMyTable.DoQuery(query);
                        impuesto = oRSMyTable.Fields.Item("IMPUESTO").Value.ToString();
                        montoTotal = oRSMyTable.Fields.Item("TOTAL").Value.ToString();
                        refeenciaPagoCanje = oRSMyTable.Fields.Item("U_PagoRef").Value.ToString();
                        docEntry = pDocumentos[0].DocEntry;

                    }
                    catch (Exception e)
                    {
                    }

                    try
                    {
                        if (rut.Length == 12)
                        {
                            rutAvalidar = rut.Substring(0, rut.Length - 1);
                            digitoAvalidar = Convert.ToInt32(rut.Substring(rut.Length - 1, 1));
                        }
                        digitoVerificadorRespuesta = validarRUT(rutAvalidar);

                        if (digitoAvalidar == digitoVerificadorRespuesta)
                        {
                            rutValidado = true;
                            clienteFinal = false;
                            decretoLey = 0;
                        }
                    }
                    catch
                    {
                        rutValidado = false;
                        clienteFinal = false;
                        decretoLey = 1;
                    }

                    //ObtenerCambioAlmacenFecha
                    foreach (clsPago doc in pDocumentos)
                    {
                        double monto, tc;
                        if ((moneda.Equals("USD") || moneda.Equals("U$S")) && monedaDePago.Equals("UYU"))
                        {
                            monto = getDouble(doc.Monto.ToString());
                            string tcTemp = ObtenerCambioAlmacenFecha(monedaDePago, fecha.ToString(configAddOn.FormatoFecha));
                            tc = Convert.ToDouble(tcTemp);
                            doc.Monto = monto * tc; // Para corregir decimales
                            montoTotalPago += doc.Monto;

                            /* else
                             {
                                 monto = getDouble(doc.Monto.ToString());
                                 tc = getDouble(doc.TasaCambio.ToString());
                                 doc.Monto = monto * tc; // Para corregir decimales
                                 montoTotalPago += doc.Monto;
                             }*/
                            //montoGravado = montoTotalPago - (Convert.ToDouble(impuesto)) * tc;

                            ///////////////////////////////////////////////////
                            ////MODIFICACIONES TN
                            //////////////////////////////////////////////////

                            double montoFactura = Convert.ToDouble(montoTotal) * tc;
                            double impuestoFactura = Convert.ToDouble(impuesto) * tc;

                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuestoFactura)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuestoFactura) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();
                            ///////////////////////////////////////////////////////////////
                            monedaTransaccion = monedaDePago;

                        }
                        else if ((moneda.Equals("USD") || moneda.Equals("U$S")) && monedaDePago.Equals("USD"))
                        {
                            monto = getDouble(doc.Monto.ToString());

                            montoTotalPago += doc.Monto;
                            // montoGravado = montoTotalPago - (Convert.ToDouble(impuesto)) / tc;


                            double montoFactura = Convert.ToDouble(montoTotal);
                            double impuestoFactura = Convert.ToDouble(impuesto);

                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuesto)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuesto) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();

                            ///////////////////////////////////////////////////////////////

                            monedaTransaccion = monedaDePago;

                        }
                        else if ((moneda.Equals("$") || moneda.Equals("UYU")) && monedaDePago.Equals("USD"))
                        {

                            monto = getDouble(doc.Monto.ToString());
                            string tcTemp = ObtenerCambioAlmacenFecha(monedaDePago, fecha.ToString(configAddOn.FormatoFecha));
                            tc = Convert.ToDouble(tcTemp);
                            doc.Monto = monto * tc; // Para corregir decimales
                            montoTotalPago += doc.Monto;

                            /*   else
                               {
                                   monto = getDouble(doc.Monto.ToString());
                                   string tcTemp = ObtenerCambioAlmacenFecha(monedaDePago, fecha.ToString(configAddOn.FormatoFecha));
                                   tc = Convert.ToDouble(tcTemp);
                                  // tc = getDouble(doc.TasaCambio.ToString());
                                   doc.Monto = monto / tc; // Para corregir decimales
                                   montoTotalPago += doc.Monto;
                               }*/

                            //montoGravado = montoTotalPago - (Convert.ToDouble(impuesto)) * tc;

                            ///////////////////////////////////////////////////
                            ////MODIFICACIONES TN
                            //////////////////////////////////////////////////

                            double montoFactura = Convert.ToDouble(montoTotal) / tc;
                            double impuestoFactura = Convert.ToDouble(impuesto) / tc;

                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuesto)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuesto) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();
                            ///////////////////////////////////////////////////////////////
                            monedaTransaccion = monedaDePago;

                        }
                        else
                        {
                            doc.Monto = getDouble(doc.Monto.ToString()); // Para corregir decimales
                            montoTotalPago += doc.Monto;
                            // montoGravado = montoTotalPago - Convert.ToDouble(impuesto);
                            //Calculamos el IVA segun el monto de la factura

                            ///////////////////////////////////////////////////
                            ////MODIFICACIONES TN
                            //////////////////////////////////////////////////

                            double montoFactura = Convert.ToDouble(montoTotal);
                            double impuestoFactura = Convert.ToDouble(impuesto);

                            double neto = Math.Round(montoFactura - (Convert.ToDouble(impuesto)), 2);
                            double porcent = Math.Round((Convert.ToDouble(impuesto) * 100) / neto);
                            double impuestoParcial = (montoTotalPago * porcent) / 100;
                            montoGravado = montoTotalPago - impuestoParcial;
                            impuesto = impuestoParcial.ToString();
                            ///////////////////////////////////////////////////////////////

                            monedaTransaccion = moneda;
                        }
                    }

                    //cmdDevT
                    SAPbouiCOM.ComboBox oTerminal = oFormVisor.Items.Item("cmbT1").Specific; //
                    terminal = oTerminal.Selected.Description.ToString();

                    operacion = "DEV";

                    //Mensaje para insertar tarjeta
                    if (proveedorPOS.Equals("GEOCOM"))
                    {
                        SBO_Application.MessageBox("Ingrese tarjeta en POS para realizar cancelación.");

                        ControladorGeocom cg = new ControladorGeocom(this);
                        LogGeocom objetoLog = new LogGeocom();
                        LogGeocom objetoLogVenta = new LogGeocom();
                        //datos de test
                        string PosID = terminal;
                        string systemId = configAddOn.hash;
                        string Branch = "Etarey";
                        string clientAppId = "1";
                        string userId = "1";
                        string ticketTemp = "";
                        /*
                        if (ticket.ToString().Length == 2)
                        {
                            ticketTemp = "00" + ticket.ToString();
                        }
                        else if (ticket.ToString().Length == 3)
                        {
                            ticketTemp = "0" + ticket.ToString();
                        }
                        else
                        {
                            ticketTemp = ticket.ToString();
                        }
                        */
                        objetoLogVenta = ObtenerDatosTicketGeocom(ticket.ToString(), PosID);

                        GeocomWSProductivo.PurchaseQueryResponse respuesta = cg.cancelacion(montoTotalPago, PosID, monedaTransaccion, montoGravado, Convert.ToDouble(impuesto), factura.ToString(), 0, decretoLey, montoTotalPago, systemId, Branch, clientAppId, userId, ticket, objetoLogVenta);

                        if (!respuesta.PosResponseCode.Equals("-1"))
                        {
                            //buscar en las tablas el codigo de respuesta
                            CodRespuestaPOSGeocom codPosRespuesta = CodigoPOSGeocom(respuesta.PosResponseCode);

                            if (String.IsNullOrEmpty(respuesta.OriginCardType)) objetoLog.cardtype = "-"; else objetoLog.cardtype = respuesta.OriginCardType;
                            if (String.IsNullOrEmpty(respuesta.Ci)) objetoLog.ci = "-"; else objetoLog.ci = respuesta.Ci;
                            if (String.IsNullOrEmpty(respuesta.AuthorizationCode)) objetoLog.codigoAutorizacion = "-"; else objetoLog.codigoAutorizacion = respuesta.AuthorizationCode;
                            if (String.IsNullOrEmpty(codPosRespuesta.codigo)) objetoLog.codigoRespuestaPos = "-"; else objetoLog.codigoRespuestaPos = codPosRespuesta.codigo;
                            if (String.IsNullOrEmpty(codPosRespuesta.estado)) objetoLog.codigoRespuestaPosDescripcion = "-"; else objetoLog.codigoRespuestaPosDescripcion = codPosRespuesta.estado;
                            if (String.IsNullOrEmpty(respuesta.Quota)) objetoLog.cuotas = "-"; else objetoLog.cuotas = respuesta.Quota;
                            // objetoLog.fechaTransaccion = respuesta.TransactionDate;
                            objetoLog.transactionDateTime = respuesta.TransactionHour;
                            if (String.IsNullOrEmpty(respuesta.TaxRefund)) objetoLog.impuestocodigo = "-"; else objetoLog.impuestocodigo = respuesta.TaxRefund;
                            if (String.IsNullOrEmpty(respuesta.Batch)) objetoLog.lote = "-"; else objetoLog.lote = respuesta.Batch;
                            if (String.IsNullOrEmpty(respuesta.Currency)) objetoLog.monedaTransaccionCod = "-"; else objetoLog.monedaTransaccionCod = respuesta.Currency;

                            if (objetoLog.monedaTransaccionCod.Equals("858"))
                            {
                                objetoLog.monedaTransaccionDescrip = "UYU";
                            }
                            else if (objetoLog.monedaTransaccionCod.Equals("840"))
                            {
                                //moneda = "0840"; //dolares
                                objetoLog.monedaTransaccionDescrip = "USD";
                            }
                            else
                            {
                                objetoLog.monedaTransaccionDescrip = "-";
                            }

                            //En esta seccion se buscara en la tabla de tarjetas y se traera el Issuer y su cuenta
                            IssuerGeocom issuerTemp = ObtenerDatosIssuerGeocom(respuesta.Issuer.ToString(), monedaTransaccion, objetoLogVenta.Merchant);
                            objetoLog.issuerCode = respuesta.Issuer.ToString();
                            if (String.IsNullOrEmpty(objetoLog.issuerCode)) objetoLog.issuerCode = "-";
                            if (String.IsNullOrEmpty(issuerTemp.nombreTarjeta)) objetoLog.issuerCodeDescripcion = "-"; else objetoLog.issuerCodeDescripcion = issuerTemp.nombreTarjeta;
                            // objetoLog.issuerCodeDescripcion = issuerTemp.nombreTarjeta;//ir a buscar en tabla nativa tarjetas
                            if (String.IsNullOrEmpty(issuerTemp.cuentaContable)) objetoLog.cuentaTarjeta = "-"; else objetoLog.cuentaTarjeta = issuerTemp.cuentaContable;
                            //objetoLog.cuentaTarjeta = issuerTemp.cuentaContable;
                            if (String.IsNullOrEmpty(issuerTemp.codigoTarjetaSAP)) objetoLog.codigoTarjetaSAP = "-"; else objetoLog.codigoTarjetaSAP = issuerTemp.codigoTarjetaSAP;
                            // objetoLog.codigoTarjetaSAP = issuerTemp.codigoTarjetaSAP;
                            //--------------------------------------------------------------------------------------------------//

                            if (codPosRespuesta.estado.Equals("OK"))
                            {
                                objetoLog.EstatusGeocomTransaccion = "OK";
                                double montoTemp = Convert.ToDouble(respuesta.TotalAmount) / 100;
                                objetoLog.monto = montoTemp.ToString();

                                if (String.IsNullOrEmpty(respuesta.CardOwnerName)) objetoLog.nombre = "-"; else objetoLog.nombre = respuesta.CardOwnerName;
                                if (String.IsNullOrEmpty(respuesta.EmvApplicationName)) objetoLog.nombreTarjeta = "-"; else objetoLog.nombreTarjeta = respuesta.EmvApplicationName;
                                if (String.IsNullOrEmpty(respuesta.CardNumber)) objetoLog.numerotarjeta = "-"; else objetoLog.numerotarjeta = respuesta.CardNumber;
                                if (String.IsNullOrEmpty(respuesta.Plan)) objetoLog.plan = "-"; else objetoLog.plan = respuesta.Plan;
                                if (String.IsNullOrEmpty(respuesta.PosID)) objetoLog.posId = "-"; else objetoLog.posId = respuesta.PosID;
                                if (String.IsNullOrEmpty(respuesta.AcquirerTerminal)) objetoLog.terminal = "-"; else objetoLog.terminal = respuesta.AcquirerTerminal;
                                if (String.IsNullOrEmpty(respuesta.Ticket)) objetoLog.ticket = ""; else objetoLog.ticket = respuesta.Ticket;

                                // ir a buscar en tabla SellosGeocom
                                objetoLog.selloCod = respuesta.Acquirer.ToString();
                                objetoLog.selloDescripcion = ObtenerSelloGeocom(objetoLog.selloCod);
                                if (String.IsNullOrEmpty(objetoLog.selloDescripcion)) objetoLog.selloDescripcion = "-";

                                //se va a buscar a la tabla TipoTranGeocom
                                if (String.IsNullOrEmpty(respuesta.TransactionType)) objetoLog.transaccionType = "-"; else objetoLog.transaccionType = respuesta.TransactionType;
                                objetoLog.transaccionTypeDescripcion = TipoTransaccionGeocom(objetoLog.transaccionType);
                                if (String.IsNullOrEmpty(objetoLog.transaccionTypeDescripcion)) objetoLog.transaccionTypeDescripcion = "-";

                                SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                                SAPbobsCOM.Payments oDocCanje = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);

                                string testcardcode = pDocumentos[0].CardCode;
                                /*
                                oDoc.CardName = pDocumentos[0].CardName;
                                oDoc.TransferDate = pDocumentos[0].Fecha;
                                oDoc.DocCurrency = pDocumentos[0].Moneda;
                                montoTotalPago = 0;
                                string query2 = "";
                                SAPbobsCOM.Recordset oRSMyTable2 = null;

                                */
                                //refeenciaPagoCanje

                                oDoc.GetByKey(pDocumentos[0].DocEntry);
                                string testcardname = oDoc.CardName;

                                lRetCode = oDoc.Cancel();

                                if (lRetCode != 0)
                                {
                                    objetoLog.EstatusSAPTransaccion = "Error";
                                    guardaLogGeocom(objetoLog, false);
                                    string error = oCompany.GetLastErrorDescription();
                                    int codeError = oCompany.GetLastErrorCode();
                                    SBO_Application.MessageBox("ERROR." + error.ToString());
                                }
                                else
                                {
                                    res = true;
                                    objetoLog.EstatusSAPTransaccion = "OK";

                                    if (!refeenciaPagoCanje.Equals("No Aplica"))
                                    {
                                        oDocCanje.GetByKey(Convert.ToInt32(refeenciaPagoCanje));
                                        lRetCode = oDocCanje.Cancel();
                                        if (lRetCode != 0)
                                        {
                                            SBO_Application.MessageBox("Error cancelar documento de canje." + refeenciaPagoCanje + " favor realizar cancelación manual.");
                                        }
                                    }

                                    guardaLogGeocom(objetoLog, false);
                                    //guardarLogRappelLineas(pLineasNC);
                                    //SBO_Application.MessageBox("Pago cancelado correctamente");
                                }
                            }
                            else
                            {
                                if (respuesta.PosResponseCode.Equals("CT"))
                                {
                                    SBO_Application.MessageBox("Error al realizar operación en POS.");
                                    objetoLog.EstatusGeocomTransaccion = "Transacción cancelada POS";
                                }
                                else
                                {
                                    SBO_Application.MessageBox(codPosRespuesta.mensaje);
                                    objetoLog.EstatusGeocomTransaccion = codPosRespuesta.mensaje;
                                }

                                guardaLogGeocom(objetoLog, true);
                            }
                        }
                        else if (respuesta.PosResponseCode.Equals("-1"))
                        {
                            SBO_Application.MessageBox("Error de comunicación con GEOCOM, la transacción se canceló automáticamente.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al cancelar pago", ex.Message.ToString());
                // guardaLogPagos(docEntry.ToString(), respuesta.Ticket.ToString(), respuesta.Lote.ToString(), respuesta.NroAutorizacion.ToString(), respuesta.Aprobada.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaTitular.ToString(), respuesta.DatosTransaccion.Extendida.TarjetaNombre.ToString(), respuesta.TarjetaTipo.ToString());

            }
            return res;
        }


        public async Task agregarPagoServiceLayerAsync(SLConnection serviceLayer, Pago p)
        {
            var pagoCreado = await serviceLayer.Request("IncomingPayments").PostAsync<Pago>(p);
        }

        public string obtenerItemName(String pItemCode)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            string res = "";
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select ItemName from OITM where ItemCode =  '" + pItemCode + "'";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select \"ItemName\" from \"OITM\" where \"ItemCode\" = \'" + pItemCode + "\'";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = (string)oRSMyTable.Fields.Item("ItemName").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
        }

        public List<clsLineasDocumentos> obtenerLineaDocumentos(int factura)
        {
            List<clsLineasDocumentos> linDoc = new List<clsLineasDocumentos>();
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "select t1.DocEntry ,t1.CardName, t1.CardCode, t1.DocCur, t2.ItemCode, t2.Quantity, t2.Price,t2.DiscPrcnt,t2.LineTotal, t2.AcctCode, t2.TaxCode, t2.Dscription,  t1.VatSUm from oinv t1, INV1 t2 where t1.DocEntry = t2.DocEntry and t1.DocEntry = " + factura;

            if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                query = "select \"t1.DocEntry\" ,\"t1.CardCode\", \"t1.DocCur\", \"t2.ItemCode\", \"t2.Quantity\", \"t2.Price\",\"t2.DiscPrcnt\",\"t2.LineTotal\", \"t2.AcctCode\", \"t2.TaxCode\", \"t2.Dscription\" from \"oinv\" t1, \"INV1\" t2 where \"t1.DocEntry\" = \"t2.DocEntry\" and \"t1.DocEntry\" = \"" + factura + "\"";

            oRSMyTable.DoQuery(query);


            if (oRSMyTable != null)
            {
                while (!oRSMyTable.EoF)
                {
                    clsLineasDocumentos linea = new clsLineasDocumentos();

                    linea.CardCode = oRSMyTable.Fields.Item("CardCode").Value;
                    linea.CardName = oRSMyTable.Fields.Item("CardName").Value;
                    linea.ItemCode = oRSMyTable.Fields.Item("ItemCode").Value;
                    linea.Descuento = oRSMyTable.Fields.Item("DiscPrcnt").Value;
                    linea.ItemName = obtenerItemName(linea.ItemCode);
                    linea.Cantidad = oRSMyTable.Fields.Item("Quantity").Value;
                    linea.TaxCode = oRSMyTable.Fields.Item("TaxCode").Value;
                    linea.Total = Convert.ToDecimal(oRSMyTable.Fields.Item("LineTotal").Value);
                    linea.TotalIVA = Convert.ToDecimal(oRSMyTable.Fields.Item("VatSUm").Value);
                    linDoc.Add(linea);

                    oRSMyTable.MoveNext();
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
            oRSMyTable = null;


            return linDoc;
        }

        public bool guardaLogPagos(string pFormFactura, string ticket, string lote, string numero, string estado, string nombre, string tarjeta, string tipo)
        {
            try
            {
                SAPbobsCOM.Recordset oRSMyTable = null;

                long docEntry = obtenerDocEntryLogPagos();
                DateTime fechaHoy = DateTime.Now;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = "INSERT INTO [@ADDONCAJAFACTURAS] (Code, Name, U_FACTURA, U_TICKET,U_LOTE,U_NUMERO, U_FECHA, U_ESTADO, U_NOMBRE, U_TARJETA, U_TIPO) VALUES (" + docEntry + ",'" + docEntry + "','" + pFormFactura + "','" + ticket + "','" + lote + "','" + numero.ToString() + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "','" + estado.ToString() + "','" + nombre + "','" + tarjeta + "','" + tipo + "')";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "INSERT INTO \"@ADDONCAJAFACTURAS\" (\"Code\", \"Name\", \"U_FACTURA\", \"U_TICKET\",\"U_LOTE\",\"U_NUMERO\", \"U_FECHA\", \"U_ESTADO\" , \"U_NOMBRE\", \"U_TARJETA\", \"U_TIPO\") VALUES (" + docEntry + ",'" + docEntry + "','" + pFormFactura + "','" + ticket + "','" + lote + "','" + numero.ToString() + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "','" + estado.ToString() + "','" + nombre + "','" + tarjeta + "','" + tipo + "')";

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public long obtenerDocEntryLogPagos() //Obtengo el último DocEntry de la tabla LOGPROCESO
        {
            long res = 1;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = "select case when MAX(CAST(Code AS bigint)) is null then 1 else MAX(CAST(Code AS bigint)) + 1 end as Prox from [@ADDONCAJAFACTURAS]";

                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = "select case when MAX(CAST(\"Code\" AS bigint)) is null then 1 else MAX(CAST(\"Code\" AS bigint)) + 1 end as Prox from \"@ADDONCAJAFACTURAS\"";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = Convert.ToInt64(oRSMyTable.Fields.Item("Prox").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo DocEntry de LogProceso.");
                return res;
            }
        }

        #region "Configuración Addon"

        private void CargarTerminales()
        {
            try
            {

                oFormVisor = SBO_Application.Forms.Item("Conf");


            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "Conf";
                fcp.UniqueID = "Conf";
                try
                {
                    fcp.XmlData = LoadFromXML("Conf.srf");
                    oFormVisor = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                sPath = System.Windows.Forms.Application.StartupPath;
                string imagen = sPath + "\\almacen.jpg";

                SAPbouiCOM.PictureBox oImagen = oFormVisor.Items.Item("imgEmp").Specific;
                oImagen.Picture = imagen;


                //Invenzis

                oImagen = oFormVisor.Items.Item("imgInv").Specific;
                imagen = sPath + "\\Invenzis_logo.jpg";
                oImagen.Picture = imagen;

                SAPbouiCOM.ComboBox oComboTerminal = oFormVisor.Items.Item("cmbTer").Specific;
                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                {
                    llenarCombo(oComboTerminal, "SELECT \"U_TERMINAL\"  FROM \"@ADDONCAJADATOS\"  where \"U_SUCURSALCOD\" = '" + sucursalActiva + "' AND \"U_CODUSUARIO\" = '" + usuarioLogueadoCode + "'", false, false, false, true);
                }
                else
                {
                    llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM [@ADDONCAJADATOS]", false, false, false, true);
                }

                SAPbouiCOM.ComboBox oStaticCombo;
                SAPbouiCOM.EditText oStatic;

                //Se limpian cajas de texto
                oStatic = oFormVisor.Items.Item("txtEmp").Specific;
                oStatic.Value = "";
                oStatic = oFormVisor.Items.Item("txtEmn").Specific;
                oStatic.Value = "";

                oStatic = oFormVisor.Items.Item("txtEme").Specific;
                oStatic.Value = "";

                oStatic = oFormVisor.Items.Item("txtChmn").Specific;
                oStatic.Value = "";

                oStatic = oFormVisor.Items.Item("txtChme").Specific;
                oStatic.Value = "";

                oStatic = oFormVisor.Items.Item("txttrmn").Specific;
                oStatic.Value = "";

                oStatic = oFormVisor.Items.Item("txtTrme").Specific;
                oStatic.Value = "";

                oStatic = oFormVisor.Items.Item("txtTer").Specific;
                oStatic.Value = "";



                oStaticCombo = oFormVisor.Items.Item("cmbFormatF").Specific;
                llenarComboFormatoFecha(oStaticCombo);

                oStaticCombo = oFormVisor.Items.Item("cmbPDF").Specific;
                llenarComboConfiguraciones(oStaticCombo);
            }
            catch (Exception ex)
            {

            }
        }

        private void CargarFormularioConfigAddOn(string terminal)
        {
            clsConfiguracion terminalConfig = obtenerDatosTerminal(terminal);

            try
            {
                SAPbouiCOM.ComboBox oStaticCombo;
                oStaticCombo = oFormVisor.Items.Item("cmbFormatF").Specific;
                llenarComboFormatoFecha(oStaticCombo);
                oStaticCombo.Select(terminalConfig.FormatoFecha, SAPbouiCOM.BoSearchKey.psk_ByValue);

                oStaticCombo = oFormVisor.Items.Item("cmbPDF").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                string valorImprime = terminalConfig.Imprime ? "Si" : "No";
                oStaticCombo.Select(valorImprime, SAPbouiCOM.BoSearchKey.psk_ByValue);

                /*
                oStaticCombo = oFormVisor.Items.Item("cmbPDF").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbQR").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbLog").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbCFE").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbEnvioA").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbCopiasC").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbRemitos").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbResguar").Specific;
                llenarComboConfiguraciones(oStaticCombo);
                oStaticCombo = oFormVisor.Items.Item("cmbFormatF").Specific;
                llenarComboFormatoFecha(oStaticCombo);
                */

                SAPbouiCOM.EditText oStatic;

                oStatic = oFormVisor.Items.Item("txtEmp").Specific;
                if (String.IsNullOrEmpty(terminalConfig.Empresa)) oStatic.Value = ""; else oStatic.Value = terminalConfig.Empresa;

                oStatic = oFormVisor.Items.Item("txtEmn").Specific;
                if (String.IsNullOrEmpty(terminalConfig.CajaMN)) oStatic.Value = ""; else oStatic.Value = terminalConfig.CajaMN;

                oStatic = oFormVisor.Items.Item("txtEme").Specific;
                if (String.IsNullOrEmpty(terminalConfig.CajaME)) oStatic.Value = ""; else oStatic.Value = terminalConfig.CajaME;

                oStatic = oFormVisor.Items.Item("txtChmn").Specific;
                if (String.IsNullOrEmpty(terminalConfig.ChequeMN)) oStatic.Value = ""; else oStatic.Value = terminalConfig.ChequeMN;

                oStatic = oFormVisor.Items.Item("txtChme").Specific;
                if (String.IsNullOrEmpty(terminalConfig.ChequeME)) oStatic.Value = ""; else oStatic.Value = terminalConfig.ChequeME;

                oStatic = oFormVisor.Items.Item("txttrmn").Specific;
                if (String.IsNullOrEmpty(terminalConfig.TransferenciaMN)) oStatic.Value = ""; else oStatic.Value = terminalConfig.TransferenciaMN;


                oStatic = oFormVisor.Items.Item("txtTrme").Specific;
                if (String.IsNullOrEmpty(terminalConfig.TransferenciaME)) oStatic.Value = ""; else oStatic.Value = terminalConfig.TransferenciaME;

                oStatic = oFormVisor.Items.Item("txtHash").Specific;
                if (String.IsNullOrEmpty(terminalConfig.hash)) oStatic.Value = ""; else oStatic.Value = terminalConfig.hash;

                oStatic = oFormVisor.Items.Item("txtrans").Specific;
                if (String.IsNullOrEmpty(terminalConfig.emprTransact)) oStatic.Value = ""; else oStatic.Value = terminalConfig.emprTransact;


                //CargarDatosFormularioConfig();
                oFormVisor.Visible = true;
            }
            catch (Exception ex)
            {
            }
        }


        public clsConfiguracion obtenerDatosTerminal(string terminal)
        {
            string query = "";
            clsConfiguracion confTerminal = new clsConfiguracion();

            try
            {
                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                {
                    query = String.Format("select U_EMPRESA,U_FORMATO_FECHA, U_CAJAMN, U_CAJAME,U_TRANSFMN, U_TRANSFME," +
                       "U_CHEQUEMN,U_CHEQUEME, U_TARJETAMN, U_TARJETAME, U_IMPRIME, U_TERMINAL,U_HASH,U_EMPTRANSACT from [dbo].[@ADDONCAJADATOS] WHERE U_TERMINAL = '{0}'", terminal);

                }
                else
                {
                    query = String.Format("select \"U_EMPRESA\",\"U_FORMATO_FECHA\", \"U_CAJAMN\", \"U_CAJAME\",\"U_TRANSFMN\", \"U_TRANSFME\"," +
                        "\"U_CHEQUEMN\",\"U_CHEQUEME\", \"U_TARJETAMN\", \"U_TARJETAME\", \"U_IMPRIME\", \"U_TERMINAL\",\"U_HASH\",\"U_EMPTRANSACT\" from \"@ADDONCAJADATOS\" WHERE \"U_TERMINAL\" = '{0}'", terminal);
                }
                oRSMyTable.DoQuery(query);

                confTerminal.terminal = terminal;
                string imprime = oRSMyTable.Fields.Item("U_IMPRIME").Value.ToString();
                confTerminal.Imprime = imprime.Equals("1") ? true : false;
                confTerminal.Empresa = oRSMyTable.Fields.Item("U_EMPRESA").Value.ToString();
                confTerminal.FormatoFecha = oRSMyTable.Fields.Item("U_FORMATO_FECHA").Value.ToString();
                confTerminal.CajaMN = oRSMyTable.Fields.Item("U_CAJAMN").Value.ToString();
                confTerminal.ChequeME = oRSMyTable.Fields.Item("U_CAJAME").Value.ToString();
                confTerminal.TransferenciaMN = oRSMyTable.Fields.Item("U_TRANSFMN").Value.ToString();
                confTerminal.TransferenciaME = oRSMyTable.Fields.Item("U_TRANSFME").Value.ToString();
                confTerminal.ChequeMN = oRSMyTable.Fields.Item("U_CHEQUEMN").Value.ToString();
                confTerminal.ChequeME = oRSMyTable.Fields.Item("U_CHEQUEME").Value.ToString();
                confTerminal.hash = oRSMyTable.Fields.Item("U_HASH").Value.ToString();
                confTerminal.emprTransact = oRSMyTable.Fields.Item("U_EMPTRANSACT").Value.ToString();
            }
            catch
            {
                return null;
            }

            return confTerminal;
        }

        private void llenarComboEmpresas(SAPbouiCOM.ComboBox pCombo)
        {
            try
            {
                pCombo.ValidValues.Add("INVEN", "Invenzis");
                pCombo.ValidValues.Add("LAMEC", "Lameco Viajes");
                pCombo.ValidValues.Add("NALIS", "Nalistar");
                pCombo.ValidValues.Add("MULTI", "Multimar");
                pCombo.ValidValues.Add("PONTY", "Pontyn");
                pCombo.ValidValues.Add("SUSVI", "Susviela");
            }
            catch (Exception ex)
            { }
        }

        // Llena el objeto ComboBox con los tipos de documentos
        private void llenarComboConfiguraciones(SAPbouiCOM.ComboBox pCombo)
        {
            try
            {
                pCombo.ValidValues.Add("SI", "Si");
                pCombo.ValidValues.Add("NO", "No");
            }
            catch (Exception ex)
            { }
        }


        // Llena el objeto ComboBox con los formatos de Fechas
        private void llenarComboFormatoFecha(SAPbouiCOM.ComboBox pCombo)
        {
            try
            {
                pCombo.ValidValues.Add("YMD", "yyyy-MM-dd");
                pCombo.ValidValues.Add("DMY", "dd-MM-yyyy");
            }
            catch (Exception ex)
            { }
        }

        #endregion
        #region "cancelacionpagos"

        private void CargarFormularioPagos()
        {
            //crearUserTable("LogCaja", "Log de transacción");
            try
            {
                oFormVisor = SBO_Application.Forms.Item("VisorPagos");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisorPago";
                fcp.UniqueID = "VisorPago";
                try
                {
                    fcp.XmlData = LoadFromXML("VisorPagos.srf");
                    oFormVisor = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                sPath = System.Windows.Forms.Application.StartupPath;
                string imagen = sPath + "\\almacen.jpg";

                SAPbouiCOM.PictureBox oImagen = oFormVisor.Items.Item("Item_3").Specific;
                oImagen.Picture = imagen;

                //Calendar
                oImagen = oFormVisor.Items.Item("Imgcal").Specific;
                imagen = sPath + "\\calendar.jpg";
                //imagen = sPath + "\\bmp.bmp";
                oImagen.Picture = imagen;

                //Ticket

                oImagen = oFormVisor.Items.Item("Item_9").Specific;
                imagen = sPath + "\\ticket.jpg";
                oImagen.Picture = imagen;

                oImagen = oFormVisor.Items.Item("pago").Specific;
                imagen = sPath + "\\credito.jpg";
                oImagen.Picture = imagen;

                //Invenzis


                oImagen = oFormVisor.Items.Item("imgInv2").Specific;
                imagen = sPath + "\\Invenzis_logo.jpg";
                oImagen.Picture = imagen;
                /////////////////////////////

                SAPbouiCOM.EditText oStatic;
                oFormVisor.DataSources.UserDataSources.Add("Date9", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oFormVisor.DataSources.UserDataSources.Add("Date10", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date11", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date12", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date13", SAPbouiCOM.BoDataType.dt_DATE, 10);

                try
                {
                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    oStatic.DataBind.SetBound(true, "", "Date9");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    oStatic.DataBind.SetBound(true, "", "Date10");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");

                }
                catch (Exception ex)
                { }

                //Se carga grilla de pagos
                CargarGrillaPagos();



                SAPbouiCOM.ComboBox oComboTerminal = oFormVisor.Items.Item("cmbT1").Specific;
                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                {
                    llenarCombo(oComboTerminal, "SELECT \"U_TERMINAL\"  FROM \"@ADDONCAJADATOS\"  where \"U_SUCURSALCOD\" = '" + sucursalActiva + "' AND \"U_CODUSUARIO\" = '" + usuarioLogueadoCode + "' ORDER BY \"Code\" desc ", false, false, false, true);
                }
                else
                {
                    //solo Teyma
                    if (!usuarioLogueado.Equals("manager"))
                    {
                        llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM \"@ADDONCAJADATOS\" where \"U_SUCURSALCOD\" = '" + sucursalActiva + "'", false, false, false, true);
                    }
                    else
                    {
                        llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM \"@ADDONCAJADATOS\" ", false, false, false, true);
                    }
                }

                oComboTerminal.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

                SAPbouiCOM.ComboBox oComboCuotas = oFormVisor.Items.Item("cmbCT").Specific;
                oComboCuotas.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                oComboLey.Select("Aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboOperacion = oFormVisor.Items.Item("cmbOperac").Specific;
                oComboOperacion.Select("Venta", SAPbouiCOM.BoSearchKey.psk_ByValue);
                //   SAPbouiCOM.ComboBox oComboTipo = oFormVisor.Items.Item("cmbTarjeta").Specific;
                //  oComboTipo.Select("Credito", SAPbouiCOM.BoSearchKey.psk_ByValue);
                SAPbouiCOM.ComboBox oComboCtaTransf = oFormVisor.Items.Item("transfCta").Specific;
                llenarCombo(oComboCtaTransf, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaTransf\" = '1' order by Name", false, false, false, false);
                SAPbouiCOM.ComboBox oComboCtaTarjeta = oFormVisor.Items.Item("tjaCta").Specific;
                llenarCombo(oComboCtaTarjeta, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaTarjetas\" = '1' /*and \"Segment_0\" <> \'\'*/ order by Name", true, false, false, false);
                SAPbouiCOM.ComboBox oComboDescTarjeta = oFormVisor.Items.Item("tjaDesc").Specific;
                llenarCombo(oComboDescTarjeta, "select  \"CreditCard\" as Code, \"CardName\" as Name from \"OCRC\" order by Name", true, false, false, false);
                SAPbouiCOM.ComboBox oComboCtaCheque = oFormVisor.Items.Item("chCta").Specific;
                llenarCombo(oComboCtaCheque, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaCheques\" = '1' /*and \"Segment_0\" <> \'\'*/ order by Name", false, false, false, false);
                SAPbouiCOM.ComboBox oComboBancoCheque = oFormVisor.Items.Item("chBanco").Specific;
                llenarCombo(oComboBancoCheque, "select \"BankCode\" as Code, \"BankName\" as Name from \"ODSC\" order by Name", false, false, false, false);

                SAPbouiCOM.ComboBox oComboUsr = oFormVisor.Items.Item("cmbUsr").Specific;
                llenarCombo(oComboUsr, "select \"U_NAME\" as Code , \"U_NAME\" as Name from \"OUSR\" WHERE \"U_NAME\" <> ''", false, false, false, false);
                oComboUsr.Select("Administracion Florida", SAPbouiCOM.BoSearchKey.psk_ByValue);

                try
                {
                    oComboCtaTransf.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboCtaTarjeta.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboDescTarjeta.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboCtaCheque.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboBancoCheque.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboUsr.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                }
                catch (Exception ex)
                { }

                oFormVisor.Visible = true;
                if (!String.IsNullOrEmpty(configAddOn.TransferenciaMN))
                    oComboCtaTransf.Select(configAddOn.TransferenciaMN, BoSearchKey.psk_ByDescription);
                if (!String.IsNullOrEmpty(configAddOn.ChequeMN))
                    oComboCtaCheque.Select(configAddOn.ChequeMN, BoSearchKey.psk_ByDescription);
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarFormulario", ex.Message.ToString()); }
        }

        private void CargarFormularioPagosDevolucion()
        {
            //crearUserTable("LogCaja", "Log de transacción");
            try
            {
                oFormVisor = SBO_Application.Forms.Item("VisorPagos");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisorDev";
                fcp.UniqueID = "VisorDev";
                try
                {
                    fcp.XmlData = LoadFromXML("Devolucion.srf");
                    oFormVisor = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                sPath = System.Windows.Forms.Application.StartupPath;
                string imagen = sPath + "\\almacen.jpg";

                SAPbouiCOM.PictureBox oImagen = oFormVisor.Items.Item("Item_3").Specific;
                oImagen.Picture = imagen;

                //Calendar
                oImagen = oFormVisor.Items.Item("Imgcal").Specific;
                imagen = sPath + "\\calendar.jpg";
                //imagen = sPath + "\\bmp.bmp";
                oImagen.Picture = imagen;

                //Ticket
                oImagen = oFormVisor.Items.Item("Item_9").Specific;
                imagen = sPath + "\\ticket.jpg";
                oImagen.Picture = imagen;

                oImagen = oFormVisor.Items.Item("pago").Specific;
                imagen = sPath + "\\credito.jpg";
                oImagen.Picture = imagen;

                //Invenzis
                oImagen = oFormVisor.Items.Item("imgInv2").Specific;
                imagen = sPath + "\\Invenzis_logo.jpg";
                oImagen.Picture = imagen;
                /////////////////////////////
                SAPbouiCOM.EditText oStatic;
                oFormVisor.DataSources.UserDataSources.Add("Date14", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oFormVisor.DataSources.UserDataSources.Add("Date15", SAPbouiCOM.BoDataType.dt_DATE, 10);

                try
                {
                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    oStatic.DataBind.SetBound(true, "", "Date14");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    oStatic.DataBind.SetBound(true, "", "Date15");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                }
                catch (Exception ex)
                { }

                //Se carga grilla de pagos
                CargarGrillaPagosDevoluciones();

                SAPbouiCOM.ComboBox oComboTerminal = oFormVisor.Items.Item("cmbT1").Specific;
                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                {
                    llenarCombo(oComboTerminal, "SELECT \"U_TERMINAL\"  FROM \"@ADDONCAJADATOS\"  where \"U_SUCURSALCOD\" = '" + sucursalActiva + "' AND \"U_CODUSUARIO\" = '" + usuarioLogueadoCode + "' ORDER BY \"Code\" desc ", false, false, false, true);
                }
                else
                {
                    //solo Teyma
                    if (!usuarioLogueado.Equals("manager"))
                    {
                        llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM \"@ADDONCAJADATOS\" where \"U_SUCURSALCOD\" = '" + sucursalActiva + "'", false, false, false, true);
                    }
                    else
                    {
                        llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM \"@ADDONCAJADATOS\" ", false, false, false, true);
                    }
                }

                oComboTerminal.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                oFormVisor.Visible = true;
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarFormulario", ex.Message.ToString()); }
        }

        private void CargarGrillaPagos()
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormVisor != null)
                {
                    matriz = oFormVisor.Items.Item("2").Specific;
                }
                else
                {
                    oFormVisor = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormVisor.Items.Item("2").Specific;
                }

                SAPbouiCOM.EditText oStatic;
                DateTime fechaDesde = Convert.ToDateTime(DateTime.Now);
                DateTime fechaHasta = Convert.ToDateTime(DateTime.Now);
                string numTicket = "";
                string numFac = "";

                try
                {
                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaDesde.ToString("dd/MM/yyyy");

                    fechaDesde = Convert.ToDateTime(oStatic.String);

                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaHasta.ToString("dd/MM/yyyy");

                    fechaHasta = Convert.ToDateTime(oStatic.String);


                    SAPbouiCOM.EditText oTjNumTicket = oFormVisor.Items.Item("nTIcket").Specific; // numero de ticket
                    if (!String.IsNullOrEmpty(oTjNumTicket.String))
                        numTicket = oTjNumTicket.String;

                    SAPbouiCOM.EditText oTjNumFac = oFormVisor.Items.Item("nFac").Specific; // numero de ticket
                    if (!String.IsNullOrEmpty(oTjNumFac.String))
                        numFac = oTjNumFac.String;
                }
                catch (Exception ex)
                { }

                SAPbobsCOM.Recordset ds = obtenerPagos(fechaDesde, fechaHasta, numTicket, numFac);
                oFormVisor.DataSources.DataTables.Item("DatosPag").Rows.Clear();
                oFormVisor.DataSources.DataTables.Item("DatosPag").Rows.Add(ds.RecordCount);
                int cont = 0;
                string monedaTmp = "";
                while (!ds.EoF)
                {
                    try
                    {
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColDocEntry", cont, ds.Fields.Item("DocEntry").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColCardCode", cont, ds.Fields.Item("CardCode").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColCliente", cont, ds.Fields.Item("CardName").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColDir", cont, ds.Fields.Item("Address").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColCuenta", cont, ds.Fields.Item("CreditAcct").Value);

                        //Se trae el monto en USD si el doc es en USD
                        monedaTmp = ds.Fields.Item("CreditCur").Value;

                        if (configAddOn.Empresa.Equals("ALMACEN"))
                        {
                            if (monedaTmp.Equals("U$S"))
                            {
                                monedaTmp = "USD";
                            }
                            else
                            {
                                monedaTmp = "UYU";
                            }
                        }

                        if (monedaTmp.Equals("USD") && monedaSistema.Equals("UYU"))
                        {
                            oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("DocTotalFC").Value);
                        }
                        else
                        {
                            oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("DocTotal").Value);
                        }
                        // oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("DocTotal").Value);

                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColDescripcion", cont, ds.Fields.Item("Canceled").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColIDTran", cont, ds.Fields.Item("TransId").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColFecha", cont, ds.Fields.Item("CreateDate").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColTicket", cont, ds.Fields.Item("VoucherNum").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMoneda", cont, ds.Fields.Item("CreditCur").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColTranID", cont, ds.Fields.Item("ConfNum").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColTC", cont, ds.Fields.Item("Rate").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonP", cont, ds.Fields.Item("Comments").Value);

                        string folio = ds.Fields.Item("FolioPref").Value + ds.Fields.Item("FolioNum").Value;
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColFactura", cont, folio);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("Colse", cont, ds.Fields.Item("tarjeta").Value);

                        //oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMDoc", cont, ds.Fields.Item("DocTotalFC").Value);
                        //ColMonP
                        //Col_50

                        cont++;
                    }
                    catch (Exception ex)
                    { guardaLogProceso("", "", "ERROR al CargarGrilla 02", ex.Message.ToString()); }
                    ds.MoveNext();
                }

                matriz.Columns.Item("V_19").DataBind.Bind("DatosPag", "ColDocEntry");
                matriz.Columns.Item("colFac").DataBind.Bind("DatosPag", "ColFactura");
                matriz.Columns.Item("V_7").DataBind.Bind("DatosPag", "ColCardCode");
                matriz.Columns.Item("V_12").DataBind.Bind("DatosPag", "ColCliente");
                matriz.Columns.Item("V_9").DataBind.Bind("DatosPag", "ColDir");
                matriz.Columns.Item("V_10").DataBind.Bind("DatosPag", "ColCuenta");
                matriz.Columns.Item("V_11").DataBind.Bind("DatosPag", "ColMonto");
                matriz.Columns.Item("V_20").DataBind.Bind("DatosPag", "ColDescripcion");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosPag", "ColIDTran");
                matriz.Columns.Item("V_1").DataBind.Bind("DatosPag", "ColFecha");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosPag", "ColTicket");
                matriz.Columns.Item("V_8").DataBind.Bind("DatosPag", "ColMoneda");
                matriz.Columns.Item("TC").DataBind.Bind("DatosPag", "ColTC");
                matriz.Columns.Item("Col_30").DataBind.Bind("DatosPag", "ColMDoc");
                matriz.Columns.Item("Col_50").DataBind.Bind("DatosPag", "ColMonP");
                matriz.Columns.Item("Col_52").DataBind.Bind("DatosPag", "Colse");

                // Se comentan estas líneas porque se maneja desde el Event
                SAPbouiCOM.LinkedButton oLink;
                oLink = matriz.Columns.Item("V_19").ExtendedObject;
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Receipt;

                //oLink = matriz.Columns.Item("colFac").ExtendedObject;
                //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;
                //matriz.Columns.Item("colFac").Visible = true;
                //matriz.Columns.Item("V_1").Visible = false;
                //matriz.Columns.Item("V_2").Visible = false;
                //matriz.Columns.Item("V_7").Visible = false;
                matriz.Columns.Item("V_20").Visible = false;
                matriz.Columns.Item("V_9").Visible = false;
                matriz.Columns.Item("Col_30").Visible = false;
                // matriz.Columns.Item("V_3").RightJustified = true;
                // matriz.Columns.Item("V_8").RightJustified = true;
                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarGrilla 03", ex.Message.ToString()); }
        }

        private void CargarFormularioPagosError()
        {
            //crearUserTable("LogCaja", "Log de transacción");
            try
            {
                oFormVisor = SBO_Application.Forms.Item("VisorError");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisorError";
                fcp.UniqueID = "VisorError";
                try
                {
                    fcp.XmlData = LoadFromXML("VisorError.srf");
                    oFormVisor = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                sPath = System.Windows.Forms.Application.StartupPath;
                string imagen = sPath + "\\almacen.jpg";

                SAPbouiCOM.PictureBox oImagen = oFormVisor.Items.Item("Item_3").Specific;
                oImagen.Picture = imagen;

                //Calendar
                oImagen = oFormVisor.Items.Item("Imgcal").Specific;
                imagen = sPath + "\\calendar.jpg";
                //imagen = sPath + "\\bmp.bmp";
                oImagen.Picture = imagen;

                //Ticket
                oImagen = oFormVisor.Items.Item("Item_9").Specific;
                imagen = sPath + "\\ticket.jpg";
                oImagen.Picture = imagen;

                oImagen = oFormVisor.Items.Item("pago").Specific;
                imagen = sPath + "\\credito.jpg";
                oImagen.Picture = imagen;

                //Invenzis
                oImagen = oFormVisor.Items.Item("imgInv2").Specific;
                imagen = sPath + "\\Invenzis_logo.jpg";
                oImagen.Picture = imagen;
                /////////////////////////////

                SAPbouiCOM.EditText oStatic;
                oFormVisor.DataSources.UserDataSources.Add("Date16", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oFormVisor.DataSources.UserDataSources.Add("Date17", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date3", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date4", SAPbouiCOM.BoDataType.dt_DATE, 10);
                //oFormVisor.DataSources.UserDataSources.Add("Date5", SAPbouiCOM.BoDataType.dt_DATE, 10);

                try
                {
                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    oStatic.DataBind.SetBound(true, "", "Date16");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    oStatic.DataBind.SetBound(true, "", "Date17");
                    oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
                }
                catch (Exception ex)
                { }

                //Se carga grilla de pagos
                CargarGrillaPagosError();

                SAPbouiCOM.ComboBox oComboTerminal = oFormVisor.Items.Item("cmbT1").Specific;
                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                {
                    llenarCombo(oComboTerminal, "SELECT \"U_TERMINAL\"  FROM \"@ADDONCAJADATOS\"  where \"U_SUCURSALCOD\" = '" + sucursalActiva + "' AND \"U_CODUSUARIO\" = '" + usuarioLogueadoCode + "'", false, false, false, true);
                }
                else
                {
                    llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM [@ADDONCAJADATOS]", false, false, false, true);
                }

                oComboTerminal.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                SAPbouiCOM.ComboBox oComboCuotas = oFormVisor.Items.Item("cmbCT").Specific;
                oComboCuotas.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboLey = oFormVisor.Items.Item("cmdLey").Specific;
                oComboLey.Select("Aplicar devolución", SAPbouiCOM.BoSearchKey.psk_ByValue);

                SAPbouiCOM.ComboBox oComboOperacion = oFormVisor.Items.Item("cmbOperac").Specific;
                oComboOperacion.Select("Venta", SAPbouiCOM.BoSearchKey.psk_ByValue);
                //   SAPbouiCOM.ComboBox oComboTipo = oFormVisor.Items.Item("cmbTarjeta").Specific;
                //  oComboTipo.Select("Credito", SAPbouiCOM.BoSearchKey.psk_ByValue);
                SAPbouiCOM.ComboBox oComboCtaTransf = oFormVisor.Items.Item("transfCta").Specific;
                llenarCombo(oComboCtaTransf, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"Finanse\" = 'Y' and \"U_CtaTransf\" = '1' /*and \"Segment_0\" <> \'\'*/ order by Name", false, false, false, false);
                SAPbouiCOM.ComboBox oComboCtaTarjeta = oFormVisor.Items.Item("tjaCta").Specific;
                llenarCombo(oComboCtaTarjeta, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaTarjetas\" = '1' /*and \"Segment_0\" <> \'\'*/ order by Name", true, false, false, false);
                SAPbouiCOM.ComboBox oComboDescTarjeta = oFormVisor.Items.Item("tjaDesc").Specific;
                llenarCombo(oComboDescTarjeta, "select  \"CreditCard\" as Code, \"CardName\" as Name from \"OCRC\" order by Name", true, false, false, false);
                SAPbouiCOM.ComboBox oComboCtaCheque = oFormVisor.Items.Item("chCta").Specific;
                llenarCombo(oComboCtaCheque, "select \"AcctCode\" as Code, \"AcctName\" as Name from \"OACT\" where \"U_CtaCheques\" = '1' /*and \"Segment_0\" <> \'\'*/ order by Name", false, false, false, false);
                SAPbouiCOM.ComboBox oComboBancoCheque = oFormVisor.Items.Item("chBanco").Specific;
                llenarCombo(oComboBancoCheque, "select \"BankCode\" as Code, \"BankName\" as Name from \"ODSC\" order by Name", false, false, false, false);
                SAPbouiCOM.ComboBox oComboUsr = oFormVisor.Items.Item("cmbUsr").Specific;
                llenarCombo(oComboUsr, "select \"U_NAME\" as Code , \"U_NAME\" as Name from \"OUSR\" WHERE \"U_NAME\" <> ''", false, false, false, false);
                oComboUsr.Select("Administracion Florida", SAPbouiCOM.BoSearchKey.psk_ByValue);

                try
                {
                    oComboCtaTransf.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboCtaTarjeta.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboDescTarjeta.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboCtaCheque.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboBancoCheque.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oComboUsr.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                }
                catch (Exception ex)
                { }

                oFormVisor.Visible = true;
                if (!String.IsNullOrEmpty(configAddOn.TransferenciaMN))
                    oComboCtaTransf.Select(configAddOn.TransferenciaMN, BoSearchKey.psk_ByDescription);
                if (!String.IsNullOrEmpty(configAddOn.ChequeMN))
                    oComboCtaCheque.Select(configAddOn.ChequeMN, BoSearchKey.psk_ByDescription);
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarFormulario", ex.Message.ToString()); }
        }

        private void CargarGrillaPagosDevoluciones()
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormVisor != null)
                {
                    matriz = oFormVisor.Items.Item("3").Specific;
                }
                else
                {
                    oFormVisor = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormVisor.Items.Item("3").Specific;
                }

                SAPbouiCOM.EditText oStatic;
                DateTime fechaDesde = Convert.ToDateTime(DateTime.Now);
                DateTime fechaHasta = Convert.ToDateTime(DateTime.Now);
                string numTicket = "";
                string numFac = "";

                try
                {
                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaDesde.ToString("dd/MM/yyyy");

                    fechaDesde = Convert.ToDateTime(oStatic.String);

                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaHasta.ToString("dd/MM/yyyy");

                    fechaHasta = Convert.ToDateTime(oStatic.String);

                    SAPbouiCOM.EditText oTjNumTicket = oFormVisor.Items.Item("nTIcket").Specific; // numero de ticket
                    if (!String.IsNullOrEmpty(oTjNumTicket.String))
                        numTicket = oTjNumTicket.String;

                    SAPbouiCOM.EditText oTjNumFac = oFormVisor.Items.Item("nFac").Specific; // numero de ticket
                    if (!String.IsNullOrEmpty(oTjNumFac.String))
                        numFac = oTjNumFac.String;
                }
                catch (Exception ex)
                { }

                SAPbobsCOM.Recordset ds = obtenerPagos(fechaDesde, fechaHasta, numTicket, numFac);
                oFormVisor.DataSources.DataTables.Item("DatosPag").Rows.Clear();
                oFormVisor.DataSources.DataTables.Item("DatosPag").Rows.Add(ds.RecordCount);
                int cont = 0;
                string monedaTmp = "";

                while (!ds.EoF)
                {
                    try
                    {
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColDocEntry", cont, ds.Fields.Item("DocEntry").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColCardCode", cont, ds.Fields.Item("CardCode").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColCliente", cont, ds.Fields.Item("CardName").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColDir", cont, ds.Fields.Item("Address").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColCuenta", cont, ds.Fields.Item("CreditAcct").Value);

                        //Se trae el monto en USD si el doc es en USD
                        monedaTmp = ds.Fields.Item("CreditCur").Value;

                        if (configAddOn.Empresa.Equals("ALMACEN"))
                        {
                            if (monedaTmp.Equals("U$S"))
                            {
                                monedaTmp = "USD";
                            }
                            else
                            {
                                monedaTmp = "UYU";
                            }
                        }

                        if (monedaTmp.Equals("USD") && monedaSistema.Equals("UYU"))
                        {
                            oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("DocTotalFC").Value);
                        }
                        else
                        {
                            oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("DocTotal").Value);
                        }
                        // oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("DocTotal").Value);

                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColDescripcion", cont, ds.Fields.Item("Canceled").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColIDTran", cont, ds.Fields.Item("TransId").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColFecha", cont, ds.Fields.Item("CreateDate").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColTicket", cont, ds.Fields.Item("VoucherNum").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMoneda", cont, ds.Fields.Item("CreditCur").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColTranID", cont, ds.Fields.Item("ConfNum").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColTC", cont, ds.Fields.Item("Rate").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonP", cont, ds.Fields.Item("Comments").Value);

                        string folio = ds.Fields.Item("FolioPref").Value + ds.Fields.Item("FolioNum").Value;
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColFactura", cont, folio);
                        oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("Colse", cont, ds.Fields.Item("tarjeta").Value);

                        //oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMDoc", cont, ds.Fields.Item("DocTotalFC").Value);
                        //ColMonP
                        //Col_50

                        cont++;
                    }
                    catch (Exception ex)
                    { guardaLogProceso("", "", "ERROR al CargarGrilla 02", ex.Message.ToString()); }
                    ds.MoveNext();
                }

                matriz.Columns.Item("V_19").DataBind.Bind("DatosPag", "ColDocEntry");
                matriz.Columns.Item("colFac").DataBind.Bind("DatosPag", "ColFactura");
                matriz.Columns.Item("V_7").DataBind.Bind("DatosPag", "ColCardCode");
                matriz.Columns.Item("V_12").DataBind.Bind("DatosPag", "ColCliente");
                matriz.Columns.Item("V_9").DataBind.Bind("DatosPag", "ColDir");
                matriz.Columns.Item("V_10").DataBind.Bind("DatosPag", "ColCuenta");
                matriz.Columns.Item("V_11").DataBind.Bind("DatosPag", "ColMonto");
                matriz.Columns.Item("V_20").DataBind.Bind("DatosPag", "ColDescripcion");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosPag", "ColIDTran");
                matriz.Columns.Item("V_1").DataBind.Bind("DatosPag", "ColFecha");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosPag", "ColTicket");
                matriz.Columns.Item("V_8").DataBind.Bind("DatosPag", "ColMoneda");
                matriz.Columns.Item("TC").DataBind.Bind("DatosPag", "ColTC");
                matriz.Columns.Item("Col_30").DataBind.Bind("DatosPag", "ColMDoc");
                matriz.Columns.Item("Col_50").DataBind.Bind("DatosPag", "ColMonP");
                matriz.Columns.Item("Col_52").DataBind.Bind("DatosPag", "Colse");

                // Se comentan estas líneas porque se maneja desde el Event
                SAPbouiCOM.LinkedButton oLink;
                oLink = matriz.Columns.Item("V_19").ExtendedObject;
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Receipt;

                //oLink = matriz.Columns.Item("colFac").ExtendedObject;
                //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;
                //matriz.Columns.Item("colFac").Visible = true;
                //matriz.Columns.Item("V_1").Visible = false;
                //matriz.Columns.Item("V_2").Visible = false;
                //matriz.Columns.Item("V_7").Visible = false;
                matriz.Columns.Item("V_20").Visible = false;
                matriz.Columns.Item("V_9").Visible = false;
                matriz.Columns.Item("Col_30").Visible = false;
                // matriz.Columns.Item("V_3").RightJustified = true;
                // matriz.Columns.Item("V_8").RightJustified = true;
                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarGrilla 03", ex.Message.ToString()); }
        }

        private void CargarGrillaPagosError()
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormVisor != null)
                {
                    matriz = oFormVisor.Items.Item("2").Specific;
                }
                else
                {
                    oFormVisor = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormVisor.Items.Item("2").Specific;
                }

                SAPbouiCOM.EditText oStatic;
                DateTime fechaDesde = Convert.ToDateTime(DateTime.Now);
                DateTime fechaHasta = Convert.ToDateTime(DateTime.Now);
                string numTicket = "";
                string numFac = "";

                try
                {
                    oStatic = oFormVisor.Items.Item("dtFechaD").Specific; // Desde Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaDesde.ToString("dd/MM/yyyy");

                    fechaDesde = Convert.ToDateTime(oStatic.String);

                    oStatic = oFormVisor.Items.Item("dtFechaH").Specific; // Hasta Fecha
                    if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.String = fechaHasta.ToString("dd/MM/yyyy");

                    fechaHasta = Convert.ToDateTime(oStatic.String);


                    SAPbouiCOM.EditText oTjNumTicket = oFormVisor.Items.Item("nTIcket").Specific; // numero de ticket
                    if (!String.IsNullOrEmpty(oTjNumTicket.String))
                        numTicket = oTjNumTicket.String;

                    SAPbouiCOM.EditText oTjNumFac = oFormVisor.Items.Item("nFac").Specific; // numero de ticket
                    if (!String.IsNullOrEmpty(oTjNumFac.String))
                        numFac = oTjNumFac.String;
                }
                catch (Exception ex)
                { }

                SAPbobsCOM.Recordset ds = obtenerPagosError(fechaDesde, fechaHasta, numTicket, numFac);
                oFormVisor.DataSources.DataTables.Item("DatosErr").Rows.Clear();
                oFormVisor.DataSources.DataTables.Item("DatosErr").Rows.Add(ds.RecordCount);
                int cont = 0;
                string monedaTmp = "";

                while (!ds.EoF)
                {
                    try
                    {
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("ColNom", cont, ds.Fields.Item("U_nombre").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("NumTarj", cont, ds.Fields.Item("U_numerotarjeta").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("MonTran", cont, ds.Fields.Item("U_monedaTransaccionDescrip").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("Sello", cont, ds.Fields.Item("U_issuerCodeDescripcion").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("PosId", cont, ds.Fields.Item("U_posId").Value);
                        string test = ds.Fields.Item("U_nombre").Value;
                        test = ds.Fields.Item("U_numerotarjeta").Value;
                        test = ds.Fields.Item("U_monedaTransaccionDescrip").Value;
                        test = ds.Fields.Item("U_issuerCodeDescripcion").Value;
                        test = ds.Fields.Item("U_posId").Value;

                        if (monedaTmp.Equals("USD") && monedaSistema.Equals("UYU"))
                            oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("U_monto").Value);
                        else
                            oFormVisor.DataSources.DataTables.Item("DatosPag").SetValue("ColMonto", cont, ds.Fields.Item("U_monto").Value);

                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("Ticket", cont, ds.Fields.Item("U_ticket").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("Tarjeta", cont, ds.Fields.Item("U_nombreTarjeta").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("ColFecha", cont, ds.Fields.Item("U_TransactionDateTime").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("ColTicket", cont, ds.Fields.Item("U_ticket").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("EstatusS", cont, ds.Fields.Item("U_EstatusSAP").Value);
                        oFormVisor.DataSources.DataTables.Item("DatosErr").SetValue("EstatusG", cont, ds.Fields.Item("U_EstatusGeocom").Value);

                        cont++;
                    }
                    catch (Exception ex)
                    { guardaLogProceso("", "", "ERROR al CargarGrilla 02", ex.Message.ToString()); }
                    ds.MoveNext();
                }

                matriz.Columns.Item("V_19").DataBind.Bind("DatosErr", "ColNom");
                matriz.Columns.Item("colFac").DataBind.Bind("DatosErr", "NumTarj");
                matriz.Columns.Item("V_7").DataBind.Bind("DatosErr", "MonTran");
                matriz.Columns.Item("V_12").DataBind.Bind("DatosErr", "Sello");
                matriz.Columns.Item("V_9").DataBind.Bind("DatosErr", "PosId");
                matriz.Columns.Item("V_11").DataBind.Bind("DatosErr", "ColMonto");
                matriz.Columns.Item("V_10").DataBind.Bind("DatosErr", "Ticket");
                matriz.Columns.Item("V_20").DataBind.Bind("DatosErr", "Tarjeta");
                matriz.Columns.Item("V_1").DataBind.Bind("DatosErr", "ColFecha");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosErr", "ColTicket");
                matriz.Columns.Item("V_8").DataBind.Bind("DatosErr", "EstatusS");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosErr", "EstatusG");

                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al CargarGrilla 03", ex.Message.ToString()); }
        }

        #region validacion RUT

        public int validarRUT(string numero)
        {
            bool retorno = false;
            int[] digitos = new int[numero.Length];
            int factor;
            int suma = 0;
            int modulo = 0;
            int digitoVerificador = -1;
            try
            {
                factor = 2;
                int total = digitos.Length - 1;
                for (int i = total; i >= 0; i--)
                {
                    digitos[i] = Convert.ToInt32("" + numero[i]);
                    suma = suma + (digitos[i] * factor);
                    factor = factor == 9 ? 2 : (factor + 1);
                }
                //calculo el modulo 11 de la suma
                modulo = suma % 11;
                digitoVerificador = 11 - modulo;
                if (digitoVerificador == 11)
                {
                    digitoVerificador = 0;
                }
                if (digitoVerificador == 10)
                {
                    digitoVerificador = 1;
                }
            }
            catch (Exception e)
            {
                digitoVerificador = -1;
            }

            return digitoVerificador;
        }
        #endregion


        public SAPbobsCOM.Recordset obtenerPagos(DateTime pDesde, DateTime pHasta, string numTicket, string numFac)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            string query = "";

            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (tipoConexionBaseDatos.Equals("HANNA"))
                {
                    if (monedaSistema.Equals("UYU"))
                    {
                        if (configAddOn.Empresa.Equals("ALMACEN"))
                        {
                            query = "select DISTINCT t6.\"DocEntry\", t6.\"CardCode\", t6.\"CardName\", t6.\"Address\", t6.\"CreditAcct\", t6.\"DocTotal\",t6.\"Canceled\",t6.\"TransId\", t6.\"CreateDate\",t6.\"VoucherNum\",t6.\"CreditCur\",T6.\"ConfNum\", t6.\"Ref1\", t6.\"DocRate\", t6.\"DocTotalFC\", t6.\"Comments\", t6.\"Rate\", t6.\"DocNum\", t2.\"FolioPref\", t2.\"FolioNum\" , t6.\"Tarjeta\" from(" +
                                "SELECT t5.\"CardName\" as \"Tarjeta\", t1.\"DocEntry\", t1.\"CardCode\", t1.\"CardName\", t1.\"Address\", t2.\"CreditAcct\", t1.\"DocTotal\",t1.\"Canceled\",t1.\"TransId\", t1.\"CreateDate\",t2.\"VoucherNum\",t2.\"CreditCur\",T2.\"ConfNum\", t3.\"Ref1\", t1.\"DocRate\", t1.\"DocTotalFC\", t1.\"Comments\", t4.\"Rate\", t1.\"DocNum\"" +
                                " FROM \"ORCT\" T1, \"RCT3\" T2, (SELECT t0.\"DocEntry\", t2.\"DocNum\", t0.\"Ref1\", t2.\"FolioPref\", t2.\"FolioNum\"  FROM \"ORCT\" T0 LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\"=T0.\"DocNum\" LEFT JOIN \"OINV\" T2 ON T2.\"DocNum\"=T1.\"DocEntry\") T3," +
                                " (select * from \"ORTT\" where \"Currency\" = 'U$S') T4, \"OCRC\" t5 " +
                                "where T1.\"DocEntry\" = T2.\"DocNum\" and T1.\"DocEntry\" = t3.\"DocEntry\" and t1.\"Canceled\" <> 'Y' and  t1.\"CreateDate\" = t4.\"RateDate\" ";
                        }
                        else
                        {
                            query = "select DISTINCT t6.\"DocEntry\", t6.\"CardCode\", t6.\"CardName\", t6.\"Address\", t6.\"CreditAcct\", t6.\"DocTotal\",t6.\"Canceled\",t6.\"TransId\", t6.\"CreateDate\",t6.\"VoucherNum\",t6.\"CreditCur\",T6.\"ConfNum\", t6.\"Ref1\", t6.\"DocRate\", t6.\"DocTotalFC\", t6.\"Comments\", t6.\"Rate\", t6.\"DocNum\", t2.\"FolioPref\", t2.\"FolioNum\" , t6.\"Tarjeta\" from(" +
                                "SELECT t5.\"CardName\" as \"Tarjeta\", t1.\"DocEntry\", t1.\"CardCode\", t1.\"CardName\", t1.\"Address\", t2.\"CreditAcct\", t1.\"DocTotal\",t1.\"Canceled\",t1.\"TransId\", t1.\"CreateDate\",t2.\"VoucherNum\",t2.\"CreditCur\",T2.\"ConfNum\", t3.\"Ref1\", t1.\"DocRate\", t1.\"DocTotalFC\", t1.\"Comments\", t4.\"Rate\", t1.\"DocNum\"" +
                                " FROM \"ORCT\" T1, \"RCT3\" T2, (SELECT t0.\"DocEntry\", t2.\"DocNum\", t0.\"Ref1\", t2.\"FolioPref\", t2.\"FolioNum\"  FROM \"ORCT\" T0 LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\"=T0.\"DocNum\" LEFT JOIN \"OINV\" T2 ON T2.\"DocNum\"=T1.\"DocEntry\") T3," +
                                " (select * from \"ORTT\" where \"Currency\" = 'USD') T4, \"OCRC\" t5 " +
                                "where T1.\"DocEntry\" = T2.\"DocNum\" and T1.\"DocEntry\" = t3.\"DocEntry\" and t1.\"Canceled\" <> 'Y' and  t1.\"CreateDate\" = t4.\"RateDate\" ";
                        }
                    }
                    else
                    {
                        if (configAddOn.Empresa.Equals("ALMACEN"))
                        {
                            query = "select DISTINCT t6.\"DocEntry\", t6.\"CardCode\", t6.\"CardName\", t6.\"Address\", t6.\"CreditAcct\", t6.\"DocTotal\",t6.\"Canceled\",t6.\"TransId\", t6.\"CreateDate\",t6.\"VoucherNum\",t6.\"CreditCur\",T6.\"ConfNum\", t6.\"Ref1\", t6.\"DocRate\", t6.\"DocTotalFC\", t6.\"Comments\", t6.\"Rate\", t6.\"DocNum\", t2.\"FolioPref\", t2.\"FolioNum\" , t6.\"Tarjeta\" from(" +
                                "SELECT t5.\"CardName\" as \"Tarjeta\", t1.\"DocEntry\", t1.\"CardCode\", t1.\"CardName\", t1.\"Address\", t2.\"CreditAcct\", t1.\"DocTotal\",t1.\"Canceled\",t1.\"TransId\", t1.\"CreateDate\",t2.\"VoucherNum\",t2.\"CreditCur\",T2.\"ConfNum\", t3.\"Ref1\", t1.\"DocRate\", t1.\"DocTotalFC\", t1.\"Comments\", t4.\"Rate\", t1.\"DocNum\"" +
                                " FROM \"ORCT\" T1, \"RCT3\" T2, (SELECT t0.\"DocEntry\", t2.\"DocNum\", t0.\"Ref1\", t2.\"FolioPref\", t2.\"FolioNum\"  FROM \"ORCT\" T0 LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\"=T0.\"DocNum\" LEFT JOIN \"OINV\" T2 ON T2.\"DocNum\"=T1.\"DocEntry\") T3," +
                                " (select * from \"ORTT\" where \"Currency\" = '$') T4, \"OCRC\" t5 " +
                                "where T1.\"DocEntry\" = T2.\"DocNum\" and T1.\"DocEntry\" = t3.\"DocEntry\" and t1.\"Canceled\" <> 'Y' and  t1.\"CreateDate\" = t4.\"RateDate\" ";
                        }
                        else
                        {
                            query = "select DISTINCT t6.\"DocEntry\", t6.\"CardCode\", t6.\"CardName\", t6.\"Address\", t6.\"CreditAcct\", t6.\"DocTotal\",t6.\"Canceled\",t6.\"TransId\", t6.\"CreateDate\",t6.\"VoucherNum\",t6.\"CreditCur\",T6.\"ConfNum\", t6.\"Ref1\", t6.\"DocRate\", t6.\"DocTotalFC\", t6.\"Comments\", t6.\"Rate\", t6.\"DocNum\", t2.\"FolioPref\", t2.\"FolioNum\" , t6.\"Tarjeta\" from(" +
                                "SELECT t5.\"CardName\" as \"Tarjeta\", t1.\"DocEntry\", t1.\"CardCode\", t1.\"CardName\", t1.\"Address\", t2.\"CreditAcct\", t1.\"DocTotal\",t1.\"Canceled\",t1.\"TransId\", t1.\"CreateDate\",t2.\"VoucherNum\",t2.\"CreditCur\",T2.\"ConfNum\", t3.\"Ref1\", t1.\"DocRate\", t1.\"DocTotalFC\", t1.\"Comments\", t4.\"Rate\", t1.\"DocNum\"" +
                                " FROM \"ORCT\" T1, \"RCT3\" T2, (SELECT t0.\"DocEntry\", t2.\"DocNum\", t0.\"Ref1\", t2.\"FolioPref\", t2.\"FolioNum\"  FROM \"ORCT\" T0 LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\"=T0.\"DocNum\" LEFT JOIN \"OINV\" T2 ON T2.\"DocNum\"=T1.\"DocEntry\") T3," +
                                " (select * from \"ORTT\" where \"Currency\" = '$') T4, \"OCRC\" t5 " +
                                "where T1.\"DocEntry\" = T2.\"DocNum\" and T1.\"DocEntry\" = t3.\"DocEntry\" and t1.\"Canceled\" <> 'Y' and  t1.\"CreateDate\" = t4.\"RateDate\" ";
                        }
                    }

                    if (!String.IsNullOrEmpty(pDesde.ToString()) && !String.IsNullOrEmpty(pHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                        query += " and T1.\"DocDate\" >='" + pDesde.ToString(configAddOn.FormatoFecha) + "' and T1.\"DocDate\" <='" + pHasta.ToString(configAddOn.FormatoFecha) + "'";

                    if (!String.IsNullOrEmpty(numTicket))
                        query += " and T2.\"VoucherNum\" = " + "'" + numTicket + "'";
                    if (!String.IsNullOrEmpty(numFac))
                        query += " and T1.\"DocEntry\" = " + "'" + numFac + "'";

                    query += " and t2.\"CreditCard\" = t5.\"CreditCard\")t6 LEFT JOIN \"RCT2\" T1 ON T1.\"DocNum\" = T6.\"DocEntry\" LEFT JOIN \"OINV\" T2 ON T2.\"DocEntry\" = T1.\"DocEntry\"";
                }
                else
                {
                    if (monedaSistema.Equals("UYU"))
                    {
                        query = "select DISTINCT t6.DocEntry, t6.CardCode, t6.CardName, t6.Address, t6.CreditAcct, t6.DocTotal,t6.Canceled,t6.TransId, t6.CreateDate,t6.VoucherNum,t6.CreditCur,T6.ConfNum, t6.Ref1, t6.DocRate, t6.DocTotalFC, t6.Comments, t6.Rate, t6.DocNum, t2.FolioPref, t2.FolioNum , t6.tarjeta from(" +
                                 "SELECT t5.CardName as tarjeta, t1.DocEntry, t1.CardCode, t1.CardName, t1.Address, t2.CreditAcct, t1.DocTotal,t1.Canceled,t1.TransId, t1.CreateDate,t2.VoucherNum,t2.CreditCur,T2.ConfNum, t3.Ref1, t1.DocRate, t1.DocTotalFC, t1.Comments, t4.Rate, t1.DocNum" +
                             " FROM ORCT T1, RCT3 T2, (SELECT t0.DocEntry, t2.DocNum, t0.Ref1, t2.FolioPref, t2.FolioNum  FROM ORCT T0 LEFT JOIN RCT2 T1 ON T1.DocNum=T0.DocNum LEFT JOIN OINV T2 ON T2.DocNum=T1.DocEntry) T3," +
                            " (select * from ORTT where Currency = 'USD') T4, OCRC t5 " +
                              "where T1.DocNum = T2.DocNum and T1.DocNum = t3.DocEntry  and t1.Canceled <> 'Y' and  t1.CreateDate = t4.RateDate ";
                    }
                    else
                    {
                        query = "select DISTINCT t6.DocEntry, t6.CardCode, t6.CardName, t6.Address, t6.CreditAcct, t6.DocTotal,t6.Canceled,t6.TransId, t6.CreateDate,t6.VoucherNum,t6.CreditCur,T6.ConfNum, t6.Ref1, t6.DocRate, t6.DocTotalFC, t6.Comments, t6.Rate, t6.DocNum, t2.FolioPref, t2.FolioNum,  t6.tarjeta from(" +
                                   "SELECT t5.CardName as tarjeta, t1.DocEntry, t1.CardCode, t1.CardName, t1.Address, t2.CreditAcct, t1.DocTotal,t1.Canceled,t1.TransId, t1.CreateDate,t2.VoucherNum,t2.CreditCur,T2.ConfNum, t3.Ref1, t1.DocRate, t1.DocTotalFC, t1.Comments, t4.Rate, t1.DocNum" +
                                " FROM ORCT T1, RCT3 T2, (SELECT t0.DocEntry, t2.DocNum, t0.Ref1, t2.FolioPref, t2.FolioNum  FROM ORCT T0 LEFT JOIN RCT2 T1 ON T1.DocNum=T0.DocNum LEFT JOIN OINV T2 ON T2.DocNum=T1.DocEntry') T3," +
                                " (select * from ORTT where Currency = 'UYU') T4, OCRC t5 " +
                                   "where T1.DocNum = T2.DocNum and T1.DocNum = t3.DocEntry  and t1.Canceled <> 'Y' and  t1.CreateDate = t4.RateDate ";
                    }

                    if (!String.IsNullOrEmpty(pDesde.ToString()) && !String.IsNullOrEmpty(pHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                    {
                        if (tipoConexionBaseDatos.ToString().Equals("SQL"))
                            query += " and T1.DocDate >='" + pDesde.ToString(configAddOn.FormatoFecha) + "' and T1.DocDate <='" + pHasta.ToString(configAddOn.FormatoFecha) + "'";
                    }

                    if (!String.IsNullOrEmpty(numTicket))
                        query += " and T2.VoucherNum = " + "'" + numTicket + "'";
                    if (!String.IsNullOrEmpty(numFac))
                        query += " and T1.DocEntry = " + "'" + numFac + "'";

                    query += " and t2.CreditCard = t5.CreditCard)t6 LEFT JOIN RCT2 T1 ON T1.DocNum = T6.DocEntry LEFT JOIN OINV T2 ON T2.DocEntry = T1.DocEntry";
                }

                oRSMyTable.DoQuery(query);
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al obtenerFacturasPendientes", ex.Message.ToString()); }

            return oRSMyTable;
        }

        public SAPbobsCOM.Recordset obtenerPagosError(DateTime pDesde, DateTime pHasta, string numTicket, string numFac)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;

            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "select T0.U_nombre,T0.U_numerotarjeta,T0.U_monto,T0.U_monedaTransaccionDescrip,T0.U_TransactionDateTime,T0.U_posId,T0.U_ticket,T0.U_nombreTarjeta,T0.U_issuerCodeDescripcion, T0.U_EstatusSAP, T0.U_EstatusGeocom" +
                                " from[dbo].[@LOGGEOCOM] T0 WHERE T0.U_EstatusSAP = 'Error'";


                oRSMyTable.DoQuery(query);
            }
            catch (Exception ex)
            { guardaLogProceso("", "", "ERROR al obtenerFacturasPendientes", ex.Message.ToString()); }

            return oRSMyTable;
        }
        #endregion
        #endregion

        Boolean existe = false;
        Boolean agregoDatos = false;
        public void InitDeclareUdfs()
        {
            crearUserTable("ADDONCAJA", "DATOS CAJA");

            if (!existe)
            {
                Thread.Sleep(5000);
                CreateUserDefinedField("EMPRESA", "EMPRESA", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("FORMATO_FECHA", "FORMATO_FECHA", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, "yyyyMMdd");
                CreateUserDefinedField("CAJAMN", "CAJAMN", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("CAJAME", "CAJAME", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("TRANSFMN", "TRANSFMN", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("TRANSFME", "TRANSFME", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("CHEQUEMN", "CHEQUEMN", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("CHEQUEME", "CHEQUEME", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("TARJETAMN", "TARJETAMN", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("TARJETAME", "TARJETAME", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
                CreateUserDefinedField("IMPRIME", "IMPRIME", BoFieldTypes.db_Numeric, 50, "ADDONCAJA", null, "0");
                CreateUserDefinedField("TERMINAL", "TERMINAL", BoFieldTypes.db_Alpha, 50, "ADDONCAJA", null, null);
            }
            existe = false;
            crearUserTable("ADDONLOGS", "LOG DE DATOS");

            if (!existe)
            {
                Thread.Sleep(5000);
                CreateUserDefinedField("PANTALLA", "PANTALLA", BoFieldTypes.db_Alpha, 50, "ADDONLOGS", null, null);
                CreateUserDefinedField("CODIGO", "CODIGO", BoFieldTypes.db_Alpha, 50, "ADDONLOGS", null, null);
                CreateUserDefinedField("ACCION", "ACCION", BoFieldTypes.db_Alpha, 50, "ADDONLOGS", null, null);
                CreateUserDefinedField("LOGXML", "LOGXML", BoFieldTypes.db_Alpha, 200, "ADDONLOGS", null, null);
                CreateUserDefinedField("FECHA", "FECHA", BoFieldTypes.db_Date, 0, "ADDONLOGS", null, null);
                CreateUserDefinedField("CREATE_DATE", "CREATE_DATE", BoFieldTypes.db_Date, 0, "ADDONLOGS", null, null);
            }

            existe = false;
            crearUserTable("ADDONCAJAFACTURAS", "LOG DE PAGOS EFECTUADOS");

            if (!existe)
            {
                Thread.Sleep(5000);

                CreateUserDefinedField("NOMBRE", "NOMBRE TITULAR", BoFieldTypes.db_Alpha, 50, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("TARJETA", "TARJETA", BoFieldTypes.db_Alpha, 50, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("TIPO", "TIPO TARJETA", BoFieldTypes.db_Alpha, 50, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("FACTURA", "FACTURA", BoFieldTypes.db_Alpha, 50, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("TICKET", "TICKET", BoFieldTypes.db_Alpha, 50, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("LOTE", "LOTE", BoFieldTypes.db_Alpha, 50, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("NUMERO", "NUMERO", BoFieldTypes.db_Alpha, 200, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("FECHA", "FECHA", BoFieldTypes.db_Date, 0, "ADDONCAJAFACTURAS", null, null);
                CreateUserDefinedField("ESTADO", "ESTADO", BoFieldTypes.db_Alpha, 50, "ADDONCAJAFACTURAS", null, null);
            }

            if (agregoDatos)
            {
                SBO_Application.MessageBox("Mensaje CAJA: " + "Reiniciar SAP antes de realizar cualquier operación con el AddOn");
            }
        }

        private void CreateUserDefinedField(string Name, string Descreption, BoFieldTypes dataType, int size, string tableName, Dictionary<string, string> dictionary, string defaultValue = "")
        {
            GC.Collect();
            var SboCompany = oCompany;
            var recordset = (Recordset)SboCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            var userField = (UserFieldsMD)SboCompany.GetBusinessObject(BoObjectTypes.oUserFields);
            recordset.DoQuery("SELECT FieldId FROM cufd where AliasId ='" + Name + "' and tableid = '" + tableName + "'");
            int Count = recordset.RecordCount;
            int ufId = 0;
            if (Count != 0)
            {
                ufId = Convert.ToInt32(recordset.Fields.Item(0).Value);
            }
            Marshal.ReleaseComObject(recordset);

            if (Count == 0)
            {
                userField.TableName = tableName;
                userField.Name = Name;
                userField.Description = Descreption;
                userField.Type = dataType;

                var vv = userField.ValidValues;
                if (dictionary != null)
                {
                    var valids = new List<string>();
                    for (int i = 0; i < vv.Count; i++)
                    {
                        vv.SetCurrentLine(i);
                        valids.Add(vv.Value);
                    }
                    foreach (var pair in dictionary)
                    {
                        if (valids.Contains(pair.Key))
                        {
                            continue;
                        }
                        userField.ValidValues.Value = pair.Key;
                        userField.ValidValues.Description = pair.Value;
                        userField.ValidValues.Add();
                    }
                }

                userField.DefaultValue = defaultValue;
                if (dataType != BoFieldTypes.db_Numeric)
                    userField.Size = size;

                Marshal.ReleaseComObject(recordset);

                if (userField.Add() != 0)
                {
                    //MessageBox.Show("ERROR " + "No se pudieron agregar los campos de usuario");
                }
                else
                {
                    agregoDatos = true;
                    //MessageBox.Show("Mensaje " + "Se agregaron los campos de usuario correctamente");
                }

                Marshal.ReleaseComObject(userField);
            }
        }


        public void cargarFormularioError()
        {
            try
            {

                oFormVisor = SBO_Application.Forms.Item("VisorGeo");


            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisorGeo";
                fcp.UniqueID = "VisorGeo";
                try
                {
                    fcp.XmlData = LoadFromXML("VisorGeo.srf");
                    oFormVisor = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            oFormVisor.DataSources.UserDataSources.Add("Date7", SAPbouiCOM.BoDataType.dt_DATE, 10);
            oFormVisor.DataSources.UserDataSources.Add("Date8", SAPbouiCOM.BoDataType.dt_DATE, 10);

            sPath = System.Windows.Forms.Application.StartupPath;
            string imagen = sPath + "\\almacen.jpg";

            SAPbouiCOM.PictureBox oImagen = oFormVisor.Items.Item("img1").Specific;
            oImagen.Picture = imagen;

            //Invenzis
            oImagen = oFormVisor.Items.Item("img2").Specific;
            imagen = sPath + "\\Invenzis_logo.jpg";
            oImagen.Picture = imagen;

            SAPbouiCOM.ComboBox oComboTerminal = oFormVisor.Items.Item("term").Specific;
            if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
            {
                llenarCombo(oComboTerminal, "SELECT \"U_TERMINAL\"  FROM \"@ADDONCAJADATOS\"  where \"U_SUCURSALCOD\" = '" + sucursalActiva + "' AND \"U_CODUSUARIO\" = '" + usuarioLogueadoCode + "'", false, false, false, true);
            }
            else
            {
                //solo Teyma
                if (!usuarioLogueado.Equals("manager"))
                {
                    llenarCombo(oComboTerminal, "SELECT U_TERMINAL FROM \"@ADDONCAJADATOS\" where \"U_SUCURSALCOD\" = '" + sucursalActiva + "'", false, false, false, true);
                }
                else
                {
                    llenarCombo(oComboTerminal, "SELECT U_TERMINAL  FROM \"@ADDONCAJADATOS\" ", false, false, false, true);
                }
            }

            oComboTerminal.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

            SAPbouiCOM.EditText oStatic;

            oStatic = oFormVisor.Items.Item("fhcDesde").Specific; // Desde Fecha
            oStatic.DataBind.SetBound(true, "", "Date7");

            oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");
            oStatic = oFormVisor.Items.Item("Item_4").Specific; // Hasta Fecha
            oStatic.DataBind.SetBound(true, "", "Date8");
            oStatic.String = DateTime.Now.ToString("dd/MM/yyyy");

            oStatic = oFormVisor.Items.Item("fhcDesde").Specific;
            string desde = oStatic.Value.ToString();
            oStatic = oFormVisor.Items.Item("Item_4").Specific;
            string hasta = oStatic.Value.ToString(); //fecha hasta
            cargarGrilla(desde, hasta);
        }

        public void cargarGrilla(string pFechaDesde, string pFechaHasta)
        {
            SAPbouiCOM.Grid oGrid;
            SAPbouiCOM.DataTable tabla = oFormVisor.DataSources.DataTables.Item("dt");
            oGrid = (SAPbouiCOM.Grid)oFormVisor.Items.Item("grillaLog").Specific;
            oFormVisor.DataSources.DataTables.Item("dt").Clear();
            oFormVisor.Freeze(true);
            string where = "";
            string query = "SELECT \"Code\", \"DocEntry\",\"U_lote\" AS \"Lote\", \"U_cuentaTarjeta\" AS \"Cuenta\", \"U_ticket\" as \"Ticket\", \"U_monto\" as \"Monto\", \"U_nombreTarjeta\" AS \"Tarjeta\", \"U_posId\" AS \"ID del POS\", \"U_transaccionTypeDescripcion\" AS \"Tipo de Transaccion\", \"U_fechaTransaccion\" AS \"Fecha de Transaccion\", \"U_EstatusSAP\" AS \"Estado\"" +
                                " FROM \"@LOGGEOCOM\"";


            if (string.IsNullOrEmpty(where) && !string.IsNullOrEmpty(pFechaDesde) && !string.IsNullOrEmpty(pFechaHasta))
                where += "WHERE \"U_fechaTransaccion\" >= '" + pFechaDesde + "' AND \"U_fechaTransaccion\" <= '" + pFechaHasta + "'" + " AND \"U_EstatusSAP\" <> 'OK' and \"U_EstatusGeocom\" <> 'Anulado'"; //and \"U_EstatusGeocom\" <> 'Anulado
            else if (!string.IsNullOrEmpty(where) && !string.IsNullOrEmpty(pFechaDesde) && !string.IsNullOrEmpty(pFechaHasta))
                where += "AND \"U_fechaTransaccion\" >= '" + pFechaDesde + "' AND \"U_fechaTransaccion\" <= '" + pFechaHasta + "' " + " AND \"U_EstatusSAP\" <> 'OK' and \"U_EstatusGeocom\" <> 'Anulado"; //and \"U_EstatusGeocom\" <> 'Anulado'

            if (!string.IsNullOrEmpty(where))
                query += where;

            tabla.ExecuteQuery(query);
            oFormVisor.Freeze(false);

        }

        private bool insertarConfiguracion(string empresa, string formatofecha, string efectivoMN, string efectivoME, string chequeMN, string chequeME, string trasfMe, string transfMN, int imprime, string agregarTerminal, string hash, string empresatransact)
        {
            bool retorno = false;

            Random random = new Random();
            int num = random.Next(100, 500);

            try
            {

                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = String.Format("insert into @ADDONCAJADATOS  (Code,Name,U_EMPRESA, U_FORMATO_FECHA, U_CAJAMN,U_CAJAME, U_TRANSFMN,U_TRANSFME,U_CHEQUEMN, U_CHEQUEME, U_TARJETAMN, U_TARJETAME, U_IMPRIME, U_TERMINAL,U_HASH,U_EMPTRANSACT)" +
                                            "values({12},{13},'{0}','{1}',{2},'{3}','{4}',{5},{6},'{7}','{8}',{9},'{10}',{11},'{14}','{15}')",
                                            empresa, formatofecha, efectivoMN, efectivoME, transfMN, trasfMe, chequeMN, chequeME, "", "", imprime, agregarTerminal, num, num, hash, empresatransact);
                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = String.Format("insert into \"@ADDONCAJADATOS\"  (\"Code\", \"DocEntry\", \"U_EMPRESA\", \"U_FORMATO_FECHA\", \"U_CAJAMN\",\"U_CAJAME\", \"U_TRANSFMN\",\"U_TRANSFME\",\"U_CHEQUEMN\", \"U_CHEQUEME\", \"U_TARJETAMN\", \"U_TARJETAME\", \"U_IMPRIME\", \"U_TERMINAL\",\"U_HASH\",\"U_EMPTRANSACT\")" +
                                            "values({12},{13},'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',{10},'{11}','{14}','{15}')",
                                            empresa, formatofecha, efectivoMN, efectivoME, transfMN, trasfMe, chequeMN, chequeME, "", "", imprime, agregarTerminal, num, num, hash, empresatransact);

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                CargarTerminales();

                retorno = true;
            }
            catch (Exception ex)
            {
                retorno = false;
            }

            return retorno;
        }

        private bool modificarConfiguracion(string empresa, string formatofecha, string efectivoMN, string efectivoME, string chequeMN, string chequeME, string trasfMe, string transfMN, int imprime, string agregarTerminal, string resComboTerminal, string hash, string empresatransact)
        {
            bool retorno = false;

            try
            {

                SAPbobsCOM.Recordset oRSMyTable = null;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = String.Format(" UPDATE [@ADDONCAJADATOS] set U_EMPRESA = '{0}', U_FORMATO_FECHA = '{1}' ,U_CAJAMN ='{2}', U_CAJAME ='{3}' ,U_CHEQUEMN ='{4}' ,U_CHEQUEME = '{5}' ,U_TRANSFMN ='{6}', U_TRANSFME ='{7}' ,U_IMPRIME ={8},U_HASH ='{10}',U_EMPTRANSACT ='{11}'  WHERE U_TERMINAL = '{9}'",
                                               empresa, formatofecha, efectivoMN, efectivoME, chequeMN, chequeME, transfMN, trasfMe, imprime, resComboTerminal, hash, empresatransact);
                if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                    query = String.Format(" UPDATE \"@ADDONCAJADATOS\" set \"U_EMPRESA\" = '{0}', \"U_FORMATO_FECHA\" = '{1}' ,\"U_CAJAMN\" ='{2}', \"U_CAJAME\" ='{3}' ,\"U_CHEQUEMN\" ='{4}' ,\"U_CHEQUEME\" = '{5}' ,\"U_TRANSFMN\" ='{6}', \"U_TRANSFME\" ='{7}' ,\"U_IMPRIME\" ={8},\"U_HASH\" ={10},\"U_EMPTRANSACT\" ={11}  WHERE \"U_TERMINAL\" = '{9}'",
                                               empresa, formatofecha, efectivoMN, efectivoME, chequeMN, chequeME, transfMN, trasfMe, imprime, resComboTerminal, hash, empresatransact);

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                CargarTerminales();

                retorno = true;
            }
            catch (Exception ex)
            {
                retorno = false;
            }

            return retorno;
        }

        private void crearUserTable(string nombre, string descripcion)
        {
            SAPbobsCOM.UserTablesMD tabla;

            // Create an instance of the metadata object
            tabla = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            try
            {
                // Set the table properties
                tabla.GetByKey(nombre);
                //tabla.Remove();
                tabla.TableName = nombre;
                tabla.TableDescription = descripcion;
                tabla.TableType = SAPbobsCOM.BoUTBTableType.bott_NoObject;

                // Add the table to SBO
                int iVal = tabla.Add();
                if (iVal != 0)
                {
                    existe = true;
                    //SBO_Application.SetStatusBarMessage("Error: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Long, true);
                }
                else
                {
                    SBO_Application.StatusBar.SetText("Instalando herramientas del AddOn de Caja", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.SetStatusBarMessage("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, true);
            }
            finally
            {
                // Release the object and call garbage collection
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tabla);
                tabla = null;
            }
        }

        public string ObtenerSucActiva()
        {
            string query = "";
            string respuesta = "";
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //Solo Teyma

            if (tipoConexionBaseDatos.Equals("HANNA"))
            {
                query = "select * from \"OUSR\" t1 where \"U_NAME\" = '" + usuarioLogueado + "'";
            }

            else
            {
                // query = "select * from \"[@USUARIOSADDONCAJA]\" t0, \"OUSR\" t1 where \"t0.U_Usuario\" = \"t1.U_NAME\" and \"U_NAME\" = '" + usuarioLogueado + "'";
            }

            oRSMyTable.DoQuery(query);

            try
            {
                respuesta = (oRSMyTable.Fields.Item("Branch").Value.ToString());
            }
            catch (Exception)
            {

                throw;
            }


            return respuesta;
        }

        public int ObtenerSucursal(string docEntry)
        {
            string query = "";
            int respuesta;
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            Console.WriteLine(usuarioLogueado);

            if (tipoConexionBaseDatos.Equals("SQL"))
            {
                query = "select BPLId from OINV WHERE DocEntry = '" + docEntry + "'";
            }

            else
            {
                query = "select \"BPLId\" from \"OINV\" WHERE \"DocEntry \" = \'" + docEntry + "\'";
            }


            oRSMyTable.DoQuery(query);
            respuesta = Convert.ToInt32(oRSMyTable.Fields.Item("BPLId").Value.ToString());

            return respuesta;

        }

        public CodRespuestaPOSGeocom CodigoPOSGeocom(string codigo)
        {
            string query = "";
            int respuesta;
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            CodRespuestaPOSGeocom codPosRespuesta = new CodRespuestaPOSGeocom();
            Console.WriteLine(usuarioLogueado);

            try
            {
                if (tipoConexionBaseDatos.Equals("SQL"))
                {
                    query = "SELECT [U_Codigo],[U_Mensaje],[U_Respuesta],[U_Estado] FROM [@RESPUESTAPOS] where [U_Codigo] = '" + codigo + "'";
                }

                else
                {
                    query = "SELECT \"U_Codigo\",\"U_Mensaje\",\"U_Respuesta\",\"U_Estado\" FROM \"@RESPUESTAPOS\" where \"U_Codigo\" = '" + codigo + "'";
                }

                oRSMyTable.DoQuery(query);


                codPosRespuesta.codigo = (oRSMyTable.Fields.Item("U_Codigo").Value.ToString());
                codPosRespuesta.mensaje = (oRSMyTable.Fields.Item("U_Mensaje").Value.ToString());
                codPosRespuesta.respuesta = (oRSMyTable.Fields.Item("U_Respuesta").Value.ToString());
                codPosRespuesta.estado = (oRSMyTable.Fields.Item("U_Estado").Value.ToString());
            }
            catch (Exception)
            {

                return codPosRespuesta;
            }



            return codPosRespuesta;

        }

        //Solo ALmacen Rural
        public bool ObtenerCuentaEfectivoSucursal(string monedaDoc, string cuentaSeleccionada)
        {
            bool retorno = false;
            string query = "";
            string respuesta = "";
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {

                query = "select \"U_Moneda\" from \"@CUENTASPAGSEFECTIVO\" where \"U_Cuenta\" = " + cuentaSeleccionada; ;


                oRSMyTable.DoQuery(query);

                respuesta = (oRSMyTable.Fields.Item("U_Moneda").Value.ToString());
                if (monedaDoc.Equals(respuesta))
                {
                    retorno = true;
                }



            }
            catch (Exception)
            {

                return false;
            }

            return retorno;

        }

        public IssuerGeocom ObtenerDatosIssuerGeocom(string codigo, string moneda, string proveedor)
        {

            string query = "";
            string temp = "";
            string temp1 = "";
            int respuesta;
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            IssuerGeocom issuerTarjeta = new IssuerGeocom();

            try
            {
                if (tipoConexionBaseDatos.Equals("SQL"))
                    query = "select CreditCard,CardName,AcctCode,CompanyId,Phone from OCRC where CompanyId = '" + codigo + "-" + proveedor + "' and Phone = '" + moneda + "'";
                else
                    query = "select \"CreditCard\",\"CardName\",\"AcctCode\",\"CompanyId\",\"Phone\" from \"OCRC\" where \"CompanyId\" = '" + codigo + "-" + proveedor + "' and \"Phone\" = '" + moneda + "'";
                //   query = "SELECT TOP 1 * FROM \"@TARJETASADDON\" where \"U_ISSUER\" = '" + codigo + "' AND \"U_MONEDA\" = '" + moneda + "' AND \"U_PROVEEDOR\" = '" + proveedor + "'";

                oRSMyTable.DoQuery(query);

                issuerTarjeta.codigoTarjetaSAP = (oRSMyTable.Fields.Item("CreditCard").Value.ToString());
                temp = (oRSMyTable.Fields.Item("CompanyId").Value.ToString());
                temp1 = temp.Substring(0, 2);
                if (Convert.ToInt32(temp1.Substring(0, 1)) == 0)
                {
                    temp1 = temp1.Substring(1, 1);
                }

                issuerTarjeta.codigoTarejeta = temp1;
                issuerTarjeta.cuentaContable = (oRSMyTable.Fields.Item("AcctCode").Value.ToString());
                issuerTarjeta.nombreTarjeta = (oRSMyTable.Fields.Item("CardName").Value.ToString());
                issuerTarjeta.moneda = (oRSMyTable.Fields.Item("Phone").Value.ToString());
                // issuerTarjeta.proveedor = temp.Substring(3, temp.Length-3);
            }
            catch (Exception)
            {

                return issuerTarjeta;
            }



            return issuerTarjeta;

        }

        public string TipoTransaccionGeocom(string codigo)
        {
            string query = "";
            string respuesta = "";
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (tipoConexionBaseDatos.Equals("SQL"))
                {
                    query = "SELECT [U_Valor],[U_Descripcion] FROM [dbo].[@TIPOTRANGEOCOM] where [U_Valor]  = '" + codigo + "'";
                }

                else
                {
                    query = "SELECT \"U_Valor\",\"U_Descripcion\" FROM \"@TIPOTRANGEOCOM\" where \"U_Valor\"  = '" + codigo + "'";
                }

                oRSMyTable.DoQuery(query);
                respuesta = (oRSMyTable.Fields.Item("U_Descripcion").Value.ToString());

            }
            catch (Exception)
            {

                return respuesta;
            }

            return respuesta;

        }

        public string ObtenerSelloGeocom(string codigo)
        {
            string query = "";
            string respuesta = "";
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (tipoConexionBaseDatos.Equals("SQL"))
                {
                    query = "SELECT [U_Sello],[U_Codigo] FROM [dbo].[@SELLOSGEOCOM] where [U_Codigo] = '" + codigo + "'";
                }

                else
                {
                    query = "SELECT \"U_Sello\",\"U_Codigo\" FROM \"@SELLOSGEOCOM\" where \"U_Codigo\" = '" + codigo + "'";
                }

                oRSMyTable.DoQuery(query);
                respuesta = (oRSMyTable.Fields.Item("U_Sello").Value.ToString());

            }
            catch (Exception)
            {

                return respuesta;
            }

            return respuesta;

        }

        public LogGeocom ObtenerDatosTicketGeocom(string ticket, string terminal)
        {
            string query = "";
            string respuesta = "";
            SAPbobsCOM.Recordset oRSMyTable = null;
            LogGeocom respuestaLog = new LogGeocom();
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (tipoConexionBaseDatos.Equals("SQL"))
                {
                    query = "select * from [dbo].[@LOGGEOCOM]  where U_ticket='" + ticket + "'";
                }
                else
                {
                    query = "select * from \"@LOGGEOCOM\"  where \"U_ticket\"='" + ticket + "' and \"U_Terminal\"='" + terminal + "'";
                }

                oRSMyTable.DoQuery(query);
                respuestaLog.TransactionDateTime = (oRSMyTable.Fields.Item("U_TransactionDateTime").Value.ToString());
                respuestaLog.ticket = oRSMyTable.Fields.Item("U_ticket").Value.ToString();//Convert.ToInt32(oRSMyTable.Fields.Item("U_ticket").Value.ToString());
                respuestaLog.monto = (oRSMyTable.Fields.Item("U_monto").Value.ToString());
                respuestaLog.monedaTransaccionCod = (oRSMyTable.Fields.Item("U_monedaTransaccionCod").Value.ToString());
                respuestaLog.cuotas = (oRSMyTable.Fields.Item("U_cuotas").Value.ToString());
                respuestaLog.plan = (oRSMyTable.Fields.Item("U_plan").Value.ToString());
                respuestaLog.TaxableAmount = (oRSMyTable.Fields.Item("U_TaxableAmount").Value.ToString());
                respuestaLog.TaxRefund = (oRSMyTable.Fields.Item("U_TaxRefund").Value.ToString());
                respuestaLog.InvoiceAmount = (oRSMyTable.Fields.Item("U_InvoiceAmount").Value.ToString());
                respuestaLog.transactionDateTime = (oRSMyTable.Fields.Item("U_horaTransaccion").Value.ToString());
                respuestaLog.Merchant = (oRSMyTable.Fields.Item("U_Merchant").Value.ToString());



            }
            catch (Exception ex)
            {

                return respuestaLog;
            }

            return respuestaLog;

        }

        public bool guardaLogGeocom(LogGeocom objLog, bool venta)
        {
            try
            {

                SAPbobsCOM.Recordset oRSMyTable = null;

                long docEntry = obtenerDocEntryLogPagos();
                DateTime fechaHoy = DateTime.Now;
                DateTime fechaOperacion;
                Random random = new Random();
                int num = random.Next(10, 999);
                string temp = objLog.codigoAutorizacion + num;

                int fecha = Convert.ToInt32(fechaHoy.Day) + Convert.ToInt32(fechaHoy.Month) + Convert.ToInt32(fechaHoy.Year) + Convert.ToInt32(fechaHoy.Second) + Convert.ToInt32(fechaHoy.Minute) + Convert.ToInt32(fechaHoy.Millisecond);
                if (objLog.codigoAutorizacion.Equals("-")) objLog.codigoAutorizacion = fecha.ToString();

                if (venta)
                {
                    fechaOperacion = objLog.fechaTransaccion;
                }
                else
                {
                    fechaOperacion = DateTime.Now;
                }




                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = "";
                if (tipoConexionBaseDatos.Equals("SQL"))
                {
                    query = String.Format("INSERT INTO [dbo].[@LOGGEOCOM]([Code],[Name],[DocEntry],[U_Terminal],[U_codigoAutorizacion],[U_lote],[U_numerotarjeta],[U_nombre],[U_ci],[U_monedaTransaccionCod],[U_monedaTransaccionDescrip],[U_cuentaTarjeta],[U_nombreTarjeta]" +
                                 ",[U_selloCod],[U_selloDescripcion],[U_issuerCode],[U_issuerCodeDescripcion],[U_cardtype],[U_plan],[U_posId],[U_codigoRespuestaPos],[U_codigoRespuestaPosDescripcion],[U_cuotas],[U_impuestocodigo]" +
                                 ",[U_ticket],[U_monto],[U_fechaTransaccion],[U_transaccionType],[U_transaccionTypeDescripcion],[U_codigoTarjetaSAP],[U_horaTransaccion],[U_TaxableAmount],[U_TaxRefund],[U_InvoiceAmount],[U_TransactionDateTime],[U_EstatusGeocom],[U_EstatusSAP])" +
                                  "VALUES('{0}', '{1}', {2}, '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}'," +
                                  "'{15}', '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', '{22}', '{23}', '{24}', '{25}', '{26}', '{27}', '{28}', '{29}','{30}','{31}','{32}','{33}','{34}','{35}','{36}')",
                                  objLog.codigoAutorizacion, objLog.codigoAutorizacion, Convert.ToInt32(objLog.codigoAutorizacion), objLog.terminal, objLog.codigoAutorizacion, objLog.lote, objLog.numerotarjeta,
                                  objLog.nombre, objLog.ci, objLog.monedaTransaccionCod, objLog.monedaTransaccionDescrip, objLog.cuentaTarjeta, objLog.nombreTarjeta,
                                  objLog.selloCod, objLog.selloDescripcion, objLog.issuerCode, objLog.issuerCodeDescripcion, objLog.cardtype, objLog.plan, objLog.posId, objLog.codigoRespuestaPos, objLog.codigoRespuestaPosDescripcion,
                                  objLog.cuotas, objLog.impuestocodigo, objLog.ticket, objLog.monto, fechaOperacion, objLog.transaccionType, objLog.transaccionTypeDescripcion, objLog.codigoTarjetaSAP, objLog.transactionDateTime, objLog.TaxableAmount, objLog.TaxRefund, objLog.InvoiceAmount, objLog.TransactionDateTime,
                                  objLog.EstatusGeocomTransaccion, objLog.EstatusSAPTransaccion);
                }
                else
                {
                    //    query = "INSERT INTO \"@ADDONLOGS\" (\"Code\", \"Name\", \"U_PANTALLA\", \"U_CODIGO\",\"U_ACCION\",\"U_LOGXML\", \"U_FECHA\", \"U_CREATE_DATE\") VALUES (" + docEntry + ",'" + docEntry + "','" + pFormFactura + "','" + pCodigoFactura + "','" + pAccion + "','" + pXML.ToString() + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "','" + fechaHoy.ToString(configAddOn.FormatoFecha + " HH:mm:ss") + "')";
                    string fechaTemp = fechaOperacion.Year + "-" + fechaOperacion.Month + "-" + fechaOperacion.Day;
                    query = String.Format("INSERT INTO \"@LOGGEOCOM\"(\"Code\",\"Name\",\"DocEntry\",\"U_Terminal\",\"U_codigoAutorizacion\",\"U_lote\",\"U_numerotarjeta\",\"U_nombre\",\"U_ci\",\"U_monedaTransaccionCod\",\"U_monedaTransaccionDescrip\",\"U_cuentaTarjeta\",\"U_nombreTarjeta\"" +
                                ",\"U_selloCod\",\"U_selloDescripcion\",\"U_issuerCode\",\"U_issuerCodeDescripcion\",\"U_cardtype\",\"U_plan\",\"U_posId\",\"U_codigoRespuestaPos\",\"U_codigoRespuestaPosDescripcion\",\"U_cuotas\",\"U_impuestocodigo\"" +
                                ",\"U_ticket\",\"U_monto\",\"U_fechaTransaccion\",\"U_transaccionType\",\"U_transaccionTypeDescripcion\",\"U_codigoTarjetaSAP\",\"U_horaTransaccion\",\"U_TaxableAmount\",\"U_TaxRefund\",\"U_InvoiceAmount\",\"U_TransactionDateTime\",\"U_EstatusGeocom\",\"U_EstatusSAP\", \"U_Merchant\")" +
                                 "VALUES('{0}', '{1}', {2}, '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}'," +
                                 "'{15}', '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', '{22}', '{23}', '{24}', '{25}', '{26}', '{27}', '{28}', '{29}','{30}','{31}','{32}','{33}','{34}','{35}','{36}','{37}')",
                                 objLog.codigoAutorizacion + num, objLog.codigoAutorizacion + num, Convert.ToInt32(temp), objLog.terminal, objLog.codigoAutorizacion, objLog.lote, objLog.numerotarjeta,
                                 objLog.nombre, objLog.ci, objLog.monedaTransaccionCod, objLog.monedaTransaccionDescrip, objLog.cuentaTarjeta, objLog.nombreTarjeta,
                                 objLog.selloCod, objLog.selloDescripcion, objLog.issuerCode, objLog.issuerCodeDescripcion, objLog.cardtype, objLog.plan, objLog.posId, objLog.codigoRespuestaPos, objLog.codigoRespuestaPosDescripcion,
                                 objLog.cuotas, objLog.impuestocodigo, objLog.ticket.ToString(), objLog.monto, fechaTemp, objLog.transaccionType, objLog.transaccionTypeDescripcion, objLog.codigoTarjetaSAP, objLog.transactionDateTime, objLog.TaxableAmount, objLog.TaxRefund, objLog.InvoiceAmount, objLog.TransactionDateTime,
                                 objLog.EstatusGeocomTransaccion, objLog.EstatusSAPTransaccion, objLog.Merchant);
                }


                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;


                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool existeUsuarioRegistrado()
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            //string query = "select t1.DocEntry ,t1.CardName, t1.CardCode, t1.DocCur, t2.ItemCode, t2.Quantity, t2.Price,t2.DiscPrcnt,t2.LineTotal, t2.AcctCode, t2.TaxCode, t2.Dscription,  t1.VatSUm from oinv t1, INV1 t2 where t1.DocEntry = t2.DocEntry and t1.DocEntry = " + factura;

            if (tipoConexionBaseDatos.ToString().Equals("HANNA"))
                query = "select TOP 1 \"Code\" ,\"Name\" from \"@USUARIOSCAJA\"  where \"Code\" = '" + usuarioLogueado + "' ";

            oRSMyTable.DoQuery(query);


            if (oRSMyTable != null)
            {
                while (!oRSMyTable.EoF)
                {
                    contrasena = oRSMyTable.Fields.Item("Name").Value;
                    contrasena = Decrypt(contrasena);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                    oRSMyTable = null;
                    return true;
                }
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
            oRSMyTable = null;


            return false;
        }

        public void agregarAUDT(string nombreDeUsuario, string contrasena)
        {
            SAPbobsCOM.UserTable tabla = oCompany.UserTables.Item("USUARIOSCAJA");

            try
            {
                tabla.Code = nombreDeUsuario;
                tabla.Name = contrasena;
                int nErr = tabla.Add();

                string md = oCompany.GetLastErrorDescription();
            }
            catch (Exception ex)
            {

            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tabla);
                tabla = null;
                GC.Collect();
            }
        }

        public string encrypt(string encryptString)
        {
            string EncryptionKey = "EYTENEDZA7P7F3UNRQ4QADNTZ4RVEX8QGP9T";
            byte[] clearBytes = Encoding.Unicode.GetBytes(encryptString);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
            0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76
        });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(clearBytes, 0, clearBytes.Length);
                        cs.Close();
                    }
                    encryptString = Convert.ToBase64String(ms.ToArray());
                }
            }
            return encryptString;
        }

        public void CargarFormularioLogin()
        {
            try
            {
                oFormLogin = SBO_Application.Forms.Item("TuneUp2");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.UniqueID = "TuneUp2";
                try
                {
                    fcp.XmlData = LoadFromXML("TuneUp2.srf");

                    oFormLogin = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                SAPbouiCOM.EditText ed = oFormLogin.Items.Item("txtUsu").Specific;
                ed.Value = usuarioLogueado;
                oFormLogin.Visible = true;
            }
            catch (Exception ex)
            {

            }
        }

        public string Decrypt(string cipherText)
        {
            string EncryptionKey = "EYTENEDZA7P7F3UNRQ4QADNTZ4RVEX8QGP9T";
            cipherText = cipherText.Replace(" ", "+");
            byte[] cipherBytes = Convert.FromBase64String(cipherText);
            using (Aes encryptor = Aes.Create())
            {
                Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(EncryptionKey, new byte[] {
            0x49, 0x76, 0x61, 0x6e, 0x20, 0x4d, 0x65, 0x64, 0x76, 0x65, 0x64, 0x65, 0x76
        });
                encryptor.Key = pdb.GetBytes(32);
                encryptor.IV = pdb.GetBytes(16);
                using (MemoryStream ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(cipherBytes, 0, cipherBytes.Length);
                        cs.Close();
                    }
                    cipherText = Encoding.Unicode.GetString(ms.ToArray());
                }
            }
            return cipherText;
        }

        public bool updateLogGeocomAnulacion(int docEntry)
        {
            try
            {

                SAPbobsCOM.Recordset oRSMyTable = null;

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = "";
                if (tipoConexionBaseDatos.Equals("HANNA"))
                {
                    query = "UPDATE \"@LOGGEOCOM\" SET \"U_EstatusSAP\" = 'Anulado' where \"DocEntry\" = " + docEntry;
                }


                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;


                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


    }

}