using SDKCONTPAQNGLib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Recepcion_Mercancia
{
    public partial class Form1 : Form
    {
        // Conexión SQL
        private string connectionString;
        private readonly string iconDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "iconos"); // Directorio donde se almacenan los iconos de la aplicación
        // ===== SDK CONTPAQi =====
        private SDKCONTPAQNGLib.TSdkSesion ses = new TSdkSesion();
        private SDKCONTPAQNGLib.TSdkListaEmpresas LE = new TSdkListaEmpresas();
        SDKCONTPAQNGLib.TSdkCuenta cta = new TSdkCuenta();
        private string _empresaActual = null;

        // ===== Variables para pólizas =====
        private List<PolizaDTO> _polizasGenerar = new List<PolizaDTO>();

        // ===== Variables para Bitácora =====
        private string rutaBitacoraBase = @"D:\Users\Expiriti\Documents\Bitacora RC"; /*@"C:\Users\Soporte2\Documents\Bitacora RC";*/
        private DataTable _datosOriginales = null;

        // ===== Variables para proceso automático =====
        private System.Threading.Timer _timerAutomatico;
        private bool _procesoAutomaticoActivo = false;

        public Form1()
        {
            InitializeComponent();

            // Conectar el evento Click del botón automático
            this.BtnAutomatico.Click += new System.EventHandler(this.BtnAutomatico_Click);

            CargarCadenaConexion();
            ConfigurarDataGridView();
            InicializarSDKContpaqi();
            ConfigurarBitacora();

            if (System.IO.File.Exists(System.IO.Path.Combine(iconDir, "logo_icon.ico")))
            {
                this.Icon = new System.Drawing.Icon(System.IO.Path.Combine(iconDir, "logo_icon.ico"));
            }
        }

        #region Inicialización y Configuración

        private void CargarCadenaConexion()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["ClipDB_Connection"].ConnectionString;
                if (string.IsNullOrEmpty(connectionString))
                    throw new Exception("No se encontró la cadena de conexión en app.config");
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error al cargar cadena de conexión. Usando configuración por defecto.";
                lblEstado.ForeColor = Color.Orange;
                connectionString = @"Server=.\SQLEXPRESS;Database=ComercialSP;Integrated Security=True;TrustServerCertificate=True;";
            }
        }

        private void ConfigurarDataGridView()
        {
            if (dgvResultados == null) return;

            dgvResultados.AutoGenerateColumns = true;
            dgvResultados.AllowUserToAddRows = false;
            dgvResultados.AllowUserToDeleteRows = false;
            dgvResultados.ReadOnly = true;
            dgvResultados.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvResultados.MultiSelect = false;
            dgvResultados.RowHeadersVisible = false;
            dgvResultados.EnableHeadersVisualStyles = true;
            dgvResultados.ColumnHeadersHeight = 25;
            dgvResultados.RowTemplate.Height = 22;
        }

        private void ConfigurarBitacora()
        {
            try
            {
                // Crear la carpeta si no existe
                if (!Directory.Exists(rutaBitacoraBase))
                {
                    Directory.CreateDirectory(rutaBitacoraBase);
                }

                // Mostrar la ruta en el TextBox
                TBBitacora.Text = rutaBitacoraBase;
                TBBitacora.ReadOnly = true;
                TBBitacora.BackColor = Color.WhiteSmoke;
            }
            catch (Exception ex)
            {
                lblEstado.Text = $"Error al configurar bitácora: {ex.Message}";
                lblEstado.ForeColor = Color.Red;
            }
        }

        #endregion

        #region SDK CONTPAQi

        private void InicializarSDKContpaqi()
        {
            try
            {
                // Inicializar conexión SDK
                if (ses.conexionActiva == 0)
                    ses.iniciaConexion();

                if (ses.conexionActiva == 1 && ses.ingresoUsuario == 0)
                    ses.firmaUsuario();

                if (ses.conexionActiva == 1 && ses.ingresoUsuario == 1)
                {
                    // Conectar automáticamente a la empresa por defecto
                    ConectarEmpresaAutomatica();
                }
                else
                {
                    lblEstadoSDK.Text = "SDK CONTPAQi no disponible";
                    lblEstadoSDK.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                lblEstadoSDK.Text = $"Error SDK: {ex.Message}";
                lblEstadoSDK.ForeColor = Color.Red;
            }
        }

        private void ConectarEmpresaAutomatica()
        {
            try
            {
                // Obtener el nombre de la empresa por defecto desde app.config
                string empresaDefault = ConfigurationManager.AppSettings["CONTPAQi_EmpresaDefault"];

                if (string.IsNullOrEmpty(empresaDefault))
                {
                    lblEstadoSDK.Text = "No se encontró empresa por defecto en configuración";
                    lblEstadoSDK.ForeColor = Color.Red;
                    return;
                }

                // Buscar la empresa en la lista de empresas disponibles
                bool empresaEncontrada = false;
                int error = LE.buscaPrimero();

                while (error > 0)
                {
                    if (LE.NombreBDD.Equals(empresaDefault, StringComparison.OrdinalIgnoreCase))
                    {
                        empresaEncontrada = true;
                        break;
                    }
                    error = LE.buscaSiguiente();
                }

                if (!empresaEncontrada)
                {
                    lblEstadoSDK.Text = $"Empresa '{empresaDefault}' no encontrada";
                    lblEstadoSDK.ForeColor = Color.Red;
                    return;
                }

                // Cerrar empresa si hay una abierta
                if (!string.IsNullOrEmpty(_empresaActual))
                    ses.cierraEmpresa();

                // Abrir la empresa por defecto
                ses.abreEmpresa(empresaDefault);
                _empresaActual = empresaDefault;

                lblEstadoSDK.Text = $"Empresa conectada: {empresaDefault}";
                lblEstadoSDK.ForeColor = Color.Green;

                // Habilitar botón de generación si hay datos cargados
                btnGenerarPolizas.Enabled = (_polizasGenerar.Count > 0);
            }
            catch (Exception ex)
            {
                lblEstadoSDK.Text = $"Error al conectar empresa: {ex.Message}";
                lblEstadoSDK.ForeColor = Color.Red;
                lblEstado.Text = $"Error al conectar a CONTPAQi: {ex.Message}";
                lblEstado.ForeColor = Color.Red;
            }
        }

        #endregion

        #region Proceso Automático

        private void BtnAutomatico_Click(object sender, EventArgs e)
        {
            if (!_procesoAutomaticoActivo)
            {
                // Iniciar proceso automático
                IniciarProcesoAutomatico();
            }
            else
            {
                // Detener proceso automático
                DetenerProcesoAutomatico();
            }

        }

        private void IniciarProcesoAutomatico()
        {
            try
            {
                // Calcular el tiempo hasta el próximo día 20
                DateTime ahora = DateTime.Now;
                DateTime proximoDia20 = new DateTime(ahora.Year, ahora.Month, 20);

                // Si ya pasó el día 20 de este mes, programar para el próximo mes
                if (ahora.Day > 20 || (ahora.Day == 20 && ahora.Hour >= 0)) // Si ya pasó o es hoy pero después de la hora actual
                {
                    proximoDia20 = proximoDia20.AddMonths(1);
                }
                else if (ahora.Day == 20)
                {
                    // Si es exactamente el día 20, establecer la hora actual
                    proximoDia20 = ahora;
                }

                // Calcular los milisegundos hasta el próximo día 20
                TimeSpan tiempoHastaProximo = proximoDia20 - ahora;
                int milisegundosHastaProximo = (int)tiempoHastaProximo.TotalMilliseconds;

                // Asegurar que no sea negativo
                if (milisegundosHastaProximo < 0)
                {
                    milisegundosHastaProximo = 0;
                }

                // Configurar el timer para que se ejecute el día 20 de cada mes
                _timerAutomatico = new System.Threading.Timer(
                    EjecutarProcesoAutomatico,
                    null,
                    milisegundosHastaProximo, // Esperar hasta el próximo día 20
                    Timeout.Infinite // No repetir automáticamente
                );

                _procesoAutomaticoActivo = true;
                BtnAutomatico.Text = "Detener Automático";
                BtnAutomatico.BackColor = Color.LightCoral;

                lblEstado.Text = $"Proceso automático iniciado - Se ejecutará el {proximoDia20:dd/MM/yyyy HH:mm}";
                lblEstado.ForeColor = Color.Green;

                // Registrar en bitácora
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Proceso automático INICIADO - Próxima ejecución: {proximoDia20:yyyy-MM-dd HH:mm}\n");
            }
            catch (Exception ex)
            {
                lblEstado.Text = $"Error al iniciar proceso automático: {ex.Message}";
                lblEstado.ForeColor = Color.Red;
            }
        }

        private void DetenerProcesoAutomatico()
        {
            try
            {
                if (_timerAutomatico != null)
                {
                    _timerAutomatico.Dispose();
                    _timerAutomatico = null;
                }

                _procesoAutomaticoActivo = false;
                BtnAutomatico.Text = "Automático";
                BtnAutomatico.BackColor = SystemColors.Control;

                lblEstado.Text = "Proceso automático detenido";
                lblEstado.ForeColor = Color.Black;

                // Registrar en bitácora
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Proceso automático DETENIDO\n");
            }
            catch (Exception ex)
            {
                lblEstado.Text = $"Error al detener proceso automático: {ex.Message}";
                lblEstado.ForeColor = Color.Red;
            }
        }

        private async void EjecutarProcesoAutomatico(object state)
        {
            try
            {
                // Registrar inicio de ejecución
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Ejecutando proceso automático...\n");

                // Necesitamos invocar en el hilo principal de la UI
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(async () =>
                    {
                        await EjecutarSecuenciaAutomatica();
                    }));
                }
                else
                {
                    await EjecutarSecuenciaAutomatica();
                }
            }
            catch (Exception ex)
            {
                // Registrar error
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {ex.Message}\n");
            }
        }

        private async Task EjecutarSecuenciaAutomatica()
        {
            try
            {
                // Verificar que el proceso automático sigue activo
                if (!_procesoAutomaticoActivo) return;

                // Mostrar estado
                lblEstado.Text = "Ejecutando proceso automático...";
                lblEstado.ForeColor = Color.Blue;
                Application.DoEvents();

                // PASO 1: Ejecutar Cargar Datos
                if (_procesoAutomaticoActivo)
                {
                    btnCargar_Click(null, null);

                    // Esperar un momento para que termine la carga
                    await Task.Delay(2000);
                }

                // PASO 2: Verificar que hay pólizas para generar
                if (_procesoAutomaticoActivo && _polizasGenerar != null && _polizasGenerar.Count > 0)
                {
                    // Ejecutar Generar Pólizas
                    btnGenerarPolizas_Click(null, null);

                    // Registrar éxito
                    string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Proceso automático COMPLETADO - {_polizasGenerar.Count} pólizas procesadas\n");
                }
                else if (_procesoAutomaticoActivo)
                {
                    lblEstado.Text = "Proceso automático: No hay pólizas para generar";
                    lblEstado.ForeColor = Color.Orange;

                    // Registrar advertencia
                    string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Proceso automático: No hay pólizas para generar\n");
                }
            }
            catch (Exception ex)
            {
                // Registrar error
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Error en ejecución automática: {ex.Message}\n");

                lblEstado.Text = $"Error en proceso automático: {ex.Message}";
                lblEstado.ForeColor = Color.Red;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_timerAutomatico != null)
            {
                _timerAutomatico.Dispose();
                _timerAutomatico = null;
            }
            base.OnFormClosing(e);
        }

        #endregion

        #region Carga y Procesamiento de Datos

        private void btnCargar_Click(object sender, EventArgs e)
        {
            CargarDatosSQL();
        }

        private void CargarDatosSQL()
        {
            try
            {
                lblEstado.Text = "Conectando a la base de datos...";
                lblEstado.ForeColor = Color.Black;
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
                btnCargar.Enabled = false;
                Application.DoEvents();

                // ===== NUEVA LÓGICA: Calcular período desde 21 del mes anterior hasta 20 del mes actual =====
                DateTime ahora = DateTime.Now;

                // Calcular fecha de corte: 20 del mes actual
                DateTime fechaCorte = new DateTime(ahora.Year, ahora.Month, 20);

                // Calcular fecha de inicio: 21 del mes anterior
                DateTime fechaInicio = fechaCorte.AddMonths(-1).AddDays(1); // 21 del mes anterior

                // Si hoy es después del 20, ajustamos para que el corte sea el próximo 20
                if (ahora.Day > 20)
                {
                    fechaCorte = fechaCorte.AddMonths(1);
                    fechaInicio = fechaCorte.AddMonths(-1).AddDays(1);
                }

                // Mostrar en el label el período que se está consultando
                lblEstado.Text = $"Consultando período: {fechaInicio:dd/MM/yyyy} al {fechaCorte:dd/MM/yyyy}";
                Application.DoEvents();

                /* 
                 * CONSULTA MODIFICADA: Usa rango de fechas en lugar de mes/año específico
                 * Esto capturará documentos desde el 21 del mes anterior hasta el 20 del mes actual
                 */
                string consultaSQL = @"
            SELECT 
                ddi.Total AS neto,
                p.SegmentoCAT1 AS segcat1,
                dd.Title AS titulo,
                dd.FolioPrefix AS serie,
                dd.Folio,
                dd.DateDocument
            FROM docDocumentItem AS ddi
            LEFT JOIN docDocument AS dd 
                ON ddi.DocumentID = dd.DocumentID
            LEFT JOIN orgProduct AS p 
                ON ddi.ProductID = p.ProductID
            WHERE dd.ModuleID IN (1253, 1327)
            AND dd.DateDocument >= @FechaInicio
            AND dd.DateDocument <= @FechaCorte
            ORDER BY dd.Folio";

                DataTable tablaResultados = new DataTable();

                using (SqlConnection conexion = new SqlConnection(connectionString))
                {
                    conexion.Open();
                    lblEstado.Text = $"Ejecutando consulta para período {fechaInicio:dd/MM/yyyy} - {fechaCorte:dd/MM/yyyy}...";
                    Application.DoEvents();

                    using (SqlCommand comando = new SqlCommand(consultaSQL, conexion))
                    {
                        // Parámetros de fecha en lugar de mes/año
                        comando.Parameters.AddWithValue("@FechaInicio", fechaInicio.Date);
                        comando.Parameters.AddWithValue("@FechaCorte", fechaCorte.Date);

                        using (SqlDataAdapter adaptador = new SqlDataAdapter(comando))
                        {
                            adaptador.Fill(tablaResultados);
                        }
                    }
                }

                // Guardar los datos originales para la bitácora
                _datosOriginales = tablaResultados.Copy();

                dgvResultados.DataSource = tablaResultados;
                FormatearColumnasDataGridView();

                // Verificar estructura de datos antes de procesar
                VerificarEstructuraDatos(tablaResultados);

                // ===== VALIDACIÓN DE DUPLICADOS EN CONTPAQi =====
                if (!string.IsNullOrEmpty(_empresaActual))
                {
                    lblEstado.Text = "Validando pólizas duplicadas en CONTPAQi...";
                    Application.DoEvents();

                    // Obtener títulos existentes en CONTPAQi
                    List<string> titulosExistentes = ObtenerTitulosPolizasExistentes(_empresaActual);

                    if (titulosExistentes != null)
                    {
                        // Preparar pólizas con validación mejorada
                        PrepararPolizasDesdeDatosConValidacionMejorada(tablaResultados, titulosExistentes);
                    }
                    else
                    {
                        // Si no se pudo conectar, continuar sin validación
                        lblEstado.Text = "No se pudo conectar a CONTPAQi. Procesando sin validación.";
                        lblEstado.ForeColor = Color.Orange;
                        PrepararPolizasDesdeDatos(tablaResultados);
                    }
                }
                else
                {
                    lblEstado.Text = "No hay empresa CONTPAQi abierta. Procesando sin validación.";
                    lblEstado.ForeColor = Color.Orange;

                    // Preparar pólizas sin validación
                    PrepararPolizasDesdeDatos(tablaResultados);
                }

                int totalRegistros = tablaResultados.Rows.Count;
                lblEstado.Text = $"Carga completada: {totalRegistros} registros del período {fechaInicio:dd/MM/yyyy} - {fechaCorte:dd/MM/yyyy} ({_polizasGenerar.Count} pólizas)";
                lblEstado.ForeColor = totalRegistros > 0 ? Color.Green : Color.Orange;

                // Habilitar botón de generación si hay datos y SDK está conectado
                btnGenerarPolizas.Enabled = (_polizasGenerar.Count > 0 && !string.IsNullOrEmpty(_empresaActual));

                // Registrar en el log automático el período procesado
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "automatico.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Período procesado: {fechaInicio:yyyy-MM-dd} a {fechaCorte:yyyy-MM-dd} - {totalRegistros} registros\n");
            }
            catch (SqlException ex)
            {
                ManejarErrorSQL(ex);
            }
            catch (InvalidCastException castEx)
            {
                lblEstado.Text = "Error en conversión de datos";
                lblEstado.ForeColor = Color.Red;
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error al procesar la solicitud";
                lblEstado.ForeColor = Color.Red;
            }
            finally
            {
                progressBar.Visible = false;
                btnCargar.Enabled = true;
            }
        }

        // MÉTODO PARA VALIDAR PÓLIZAS DUPLICADAS EN CONTPAQi
        private List<string> ObtenerTitulosPolizasExistentes(string nombreBaseDatosContpaqi)
        {
            List<string> titulosExistentes = new List<string>();

            try
            {
                if (string.IsNullOrEmpty(nombreBaseDatosContpaqi))
                {
                    lblEstado.Text = "No se especificó base de datos CONTPAQi";
                    lblEstado.ForeColor = Color.Orange;
                    return titulosExistentes;
                }

                string consultaValidacion = $@"
                    SELECT [Concepto]
                    FROM [{nombreBaseDatosContpaqi}].[dbo].[Polizas]
                    WHERE CHARINDEX('Entrada de Mercancía', [Concepto]) > 0";

                // Obtener cadena de conexión para CONTPAQi
                string connectionStringContpaqi = ObtenerCadenaConexionContpaqi(nombreBaseDatosContpaqi);

                if (string.IsNullOrEmpty(connectionStringContpaqi))
                {
                    lblEstado.Text = "No se pudo obtener conexión para CONTPAQi";
                    lblEstado.ForeColor = Color.Orange;
                    return titulosExistentes;
                }

                Console.WriteLine($"Intentando conectar a: {connectionStringContpaqi}");

                using (SqlConnection conexion = new SqlConnection(connectionStringContpaqi))
                {
                    conexion.Open();
                    Console.WriteLine($"Conexión exitosa a {nombreBaseDatosContpaqi}");

                    using (SqlCommand comando = new SqlCommand(consultaValidacion, conexion))
                    using (SqlDataReader reader = comando.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string concepto = reader["Concepto"].ToString();
                            // Extraer el título del concepto usando el método mejorado
                            string titulo = ExtraerTituloDeConceptoMejorado(concepto);
                            if (!string.IsNullOrEmpty(titulo))
                            {
                                titulosExistentes.Add(titulo);
                            }
                        }
                    }
                }

                Console.WriteLine($"Se encontraron {titulosExistentes.Count} títulos existentes en CONTPAQi.");
            }
            catch (SqlException sqlEx)
            {
                // Error específico de SQL Server
                Console.WriteLine($"Error SQL al validar pólizas: {sqlEx.Message}");

                // Intentar conectar de otra manera
                titulosExistentes = ObtenerTitulosAlternativo(nombreBaseDatosContpaqi);

                if (titulosExistentes == null)
                {
                    lblEstado.Text = "Error de conexión a CONTPAQi";
                    lblEstado.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error general al obtener títulos existentes: {ex.Message}");
                lblEstado.Text = "Error al validar pólizas existentes";
                lblEstado.ForeColor = Color.Red;
            }

            return titulosExistentes;
        }

        // Método alternativo para obtener títulos
        private List<string> ObtenerTitulosAlternativo(string nombreBaseDatosContpaqi)
        {
            try
            {
                // Intentar diferentes formatos de conexión
                List<string> posiblesCadenas = new List<string>
                {
                    // Opción 1: Usar la misma conexión que tu base de datos ComercialSP
                    connectionString.Replace("ComercialSP", nombreBaseDatosContpaqi),
                    
                    // Opción 2: Conexión local con SQLEXPRESS
                    $@"Server=.\SQLEXPRESS;Database={nombreBaseDatosContpaqi};Integrated Security=True;TrustServerCertificate=True;",
                    
                    // Opción 3: Conexión local sin instancia específica
                    $@"Server=(local);Database={nombreBaseDatosContpaqi};Integrated Security=True;TrustServerCertificate=True;",
                    
                    // Opción 4: Conexión con localhost
                    $@"Server=localhost;Database={nombreBaseDatosContpaqi};Integrated Security=True;TrustServerCertificate=True;"
                };

                foreach (string cadena in posiblesCadenas)
                {
                    try
                    {
                        Console.WriteLine($"Intentando conexión alternativa: {cadena}");

                        using (SqlConnection conexion = new SqlConnection(cadena))
                        {
                            conexion.Open();

                            string consulta = $@"
                                SELECT [Concepto]
                                FROM [{nombreBaseDatosContpaqi}].[dbo].[Polizas]
                                WHERE CHARINDEX('Entrada de Mercancía', [Concepto]) > 0";

                            List<string> resultados = new List<string>();

                            using (SqlCommand comando = new SqlCommand(consulta, conexion))
                            using (SqlDataReader reader = comando.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string concepto = reader["Concepto"].ToString();
                                    string titulo = ExtraerTituloDeConceptoMejorado(concepto);
                                    if (!string.IsNullOrEmpty(titulo))
                                    {
                                        resultados.Add(titulo);
                                    }
                                }
                            }

                            Console.WriteLine($"Conexión exitosa con cadena alternativa. Encontrados: {resultados.Count} títulos.");
                            return resultados;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Fallo con cadena alternativa: {ex.Message}");
                        continue; // Intentar siguiente cadena
                    }
                }

                return null; // Todas las conexiones fallaron
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en método alternativo: {ex.Message}");
                return null;
            }
        }

        private string ExtraerTituloDeConceptoMejorado(string concepto)
        {
            if (string.IsNullOrEmpty(concepto))
                return string.Empty;

            string titulo = concepto
                .Replace("Entrada de Mercancía", "")
                .Replace("Entrada de Mercancia", "")
                .Replace("Entrada Mercancía", "")
                .Replace("Entrada Mercancia", "")
                .Trim();

            titulo = titulo.TrimStart(':', '-', '.', ' ');
            titulo = titulo.Trim();

            return titulo;
        }

        private string ObtenerCadenaConexionContpaqi(string nombreBaseDatos)
        {
            try
            {
                try
                {
                    string conexionConfig = ConfigurationManager.ConnectionStrings["CONTPAQi_Connection"]?.ConnectionString;
                    if (!string.IsNullOrEmpty(conexionConfig))
                    {
                        var builder = new SqlConnectionStringBuilder(conexionConfig)
                        {
                            InitialCatalog = nombreBaseDatos
                        };
                        return builder.ConnectionString;
                    }
                }
                catch { }

                if (!string.IsNullOrEmpty(connectionString))
                {
                    var builderPrincipal = new SqlConnectionStringBuilder(connectionString);

                    var builderContpaqi = new SqlConnectionStringBuilder
                    {
                        DataSource = builderPrincipal.DataSource,
                        InitialCatalog = nombreBaseDatos,
                        IntegratedSecurity = builderPrincipal.IntegratedSecurity,
                        TrustServerCertificate = builderPrincipal.TrustServerCertificate
                    };

                    if (!builderPrincipal.IntegratedSecurity)
                    {
                        builderContpaqi.UserID = builderPrincipal.UserID;
                        builderContpaqi.Password = builderPrincipal.Password;
                    }

                    return builderContpaqi.ConnectionString;
                }

                return $@"Server=.\SQLEXPRESS;Database={nombreBaseDatos};Integrated Security=True;TrustServerCertificate=True;";
            }
            catch
            {
                return $@"Server=.\SQLEXPRESS;Database={nombreBaseDatos};Integrated Security=True;TrustServerCertificate=True;";
            }
        }

        private string ObtenerNombreServidorDesdeCadena()
        {
            try
            {
                var builder = new SqlConnectionStringBuilder(connectionString);
                return builder.DataSource;
            }
            catch
            {
                return "Desconocido";
            }
        }

        private void VerificarEstructuraDatos(DataTable datos)
        {
            if (datos.Rows.Count == 0)
            {
                lblEstado.Text = "No se encontraron registros en la consulta.";
                lblEstado.ForeColor = Color.Orange;
                return;
            }

            StringBuilder verificacion = new StringBuilder();
            verificacion.AppendLine("Verificación de estructura de datos:");
            verificacion.AppendLine($"Total de filas: {datos.Rows.Count}");
            verificacion.AppendLine($"Total de columnas: {datos.Columns.Count}");
            verificacion.AppendLine();

            foreach (DataColumn columna in datos.Columns)
            {
                verificacion.AppendLine($"Columna: {columna.ColumnName}");
                verificacion.AppendLine($"  Tipo: {columna.DataType}");

                if (datos.Rows.Count > 0)
                {
                    object valorEjemplo = datos.Rows[0][columna];
                    verificacion.AppendLine($"  Valor ejemplo: {valorEjemplo ?? "(null)"}");
                    verificacion.AppendLine($"  Tipo real: {(valorEjemplo != null ? valorEjemplo.GetType().ToString() : "null")}");
                }

                verificacion.AppendLine();
            }

            Console.WriteLine(verificacion.ToString());

            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "datos_verificacion.log");
                File.WriteAllText(logPath, verificacion.ToString());
            }
            catch { }
        }

        private void PrepararPolizasDesdeDatosConValidacionMejorada(DataTable datos, List<string> titulosExistentes)
        {
            _polizasGenerar.Clear();

            if (datos.Rows.Count == 0)
            {
                lblEstado.Text = "No hay datos para preparar pólizas.";
                lblEstado.ForeColor = Color.Orange;
                return;
            }

            try
            {
                int pólizasDuplicadas = 0;
                int pólizasValidas = 0;
                int pólizasSimilares = 0;
                int movimientosOmitidos = 0;

                List<string> titulosProcesados = new List<string>();
                StringBuilder mensajesAdvertencia = new StringBuilder();

                var gruposPorTitulo = datos.AsEnumerable()
                    .GroupBy(row => ObtenerStringSeguro(row, "titulo"))
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key));

                foreach (var grupo in gruposPorTitulo)
                {
                    var primeraFila = grupo.First();

                    string tituloDocumento = grupo.Key;
                    string tituloLimpio = LimpiarTituloParaValidacion(tituloDocumento);
                    string conceptoCompleto = $"Entrada de Mercancía {tituloDocumento}";

                    bool esDuplicado = false;
                    string motivoDuplicado = "";

                    if (titulosExistentes.Any(t => CompararTitulos(t, tituloLimpio)))
                    {
                        esDuplicado = true;
                        motivoDuplicado = $"Título ya existe en CONTPAQi: '{tituloDocumento}'";
                    }
                    else if (titulosProcesados.Any(t => CompararTitulos(t, tituloLimpio)))
                    {
                        esDuplicado = true;
                        motivoDuplicado = $"Título duplicado en esta sesión: '{tituloDocumento}'";
                    }
                    else if (BuscarSimilitudes(tituloLimpio, titulosExistentes))
                    {
                        pólizasSimilares++;
                        // AUTOMÁTICO: Ya no preguntamos, solo registramos y continuamos
                        mensajesAdvertencia.AppendLine($"• Título similar encontrado (se procesará): '{tituloDocumento}'");
                    }

                    if (esDuplicado)
                    {
                        pólizasDuplicadas++;
                        Console.WriteLine($"Póliza duplicada - Título: {tituloDocumento}, Motivo: {motivoDuplicado}");
                        LogDuplicado(0, tituloDocumento, motivoDuplicado);
                        continue;
                    }

                    titulosProcesados.Add(tituloLimpio);

                    PolizaDTO poliza = new PolizaDTO
                    {
                        Folio = ObtenerIntSeguro(primeraFila, "Folio"),
                        Fecha = ObtenerFechaSegura(primeraFila, "DateDocument"),
                        Concepto = conceptoCompleto,
                        Titulo = tituloDocumento,
                        Serie = ObtenerStringSeguro(primeraFila, "serie"),
                        Movimientos = new List<MovimientoPolizaDTO>()
                    };

                    int consecutivoMov = 1;
                    decimal sumaNetos = 0;

                    foreach (var fila in grupo)
                    {
                        decimal neto = ObtenerDecimalSeguro(fila, "neto");
                        sumaNetos += neto;

                        string cuentaContable = ObtenerStringSeguro(fila, "segcat1");

                        Console.WriteLine($"Título {tituloDocumento}: Cuenta contable obtenida = '{cuentaContable}'");

                        if (string.IsNullOrWhiteSpace(cuentaContable))
                        {
                            movimientosOmitidos++;
                            mensajesAdvertencia.AppendLine($"• Movimiento omitido en '{tituloDocumento}': Cuenta contable vacía");
                            continue;
                        }

                        string cuentaNumerica = FormatearCuentaParaSDK(cuentaContable);
                        Console.WriteLine($"Cuenta convertida para SDK: '{cuentaContable}' -> '{cuentaNumerica}'");

                        var movimiento = new MovimientoPolizaDTO
                        {
                            NumMovimiento = consecutivoMov++,
                            CuentaContable = cuentaContable,
                            CuentaSDK = cuentaNumerica,
                            Monto = neto,
                            TipoMovimiento = ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_CARGO,
                            Referencia = $"{poliza.Serie}{poliza.Folio}",
                            Concepto = conceptoCompleto
                        };

                        poliza.Movimientos.Add(movimiento);
                    }

                    if (poliza.Movimientos.Count == 0)
                    {
                        mensajesAdvertencia.AppendLine($"• No se crearon movimientos para '{tituloDocumento}': todas las cuentas vacías");
                        continue;
                    }

                    string cuentaAbono = ObtenerCuentaContableAbono();
                    string cuentaAbonoSDK = FormatearCuentaParaSDK(cuentaAbono);

                    Console.WriteLine($"Cuenta de abono para título {tituloDocumento}: '{cuentaAbono}' -> SDK: '{cuentaAbonoSDK}'");

                    var abono = new MovimientoPolizaDTO
                    {
                        NumMovimiento = consecutivoMov,
                        CuentaContable = cuentaAbono,
                        CuentaSDK = cuentaAbonoSDK,
                        Monto = sumaNetos,
                        TipoMovimiento = ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_ABONO,
                        Referencia = $"{poliza.Serie}{poliza.Folio}",
                        Concepto = poliza.Concepto
                    };

                    poliza.Movimientos.Add(abono);
                    poliza.Total = sumaNetos;

                    _polizasGenerar.Add(poliza);
                    pólizasValidas++;
                }

                // Mostrar resumen en el panel de estado
                StringBuilder resumenEstado = new StringBuilder();
                resumenEstado.Append($"Validación: {pólizasValidas} válidas, {pólizasDuplicadas} duplicadas");
                if (pólizasSimilares > 0) resumenEstado.Append($", {pólizasSimilares} similares");
                if (movimientosOmitidos > 0) resumenEstado.Append($", {movimientosOmitidos} movs omitidos");

                lblEstado.Text = resumenEstado.ToString();
                lblEstado.ForeColor = pólizasValidas > 0 ? Color.Green : Color.Orange;

                // Mostrar advertencias detalladas en consola
                if (mensajesAdvertencia.Length > 0)
                {
                    Console.WriteLine("Advertencias durante el proceso:");
                    Console.WriteLine(mensajesAdvertencia.ToString());
                }
            }
            catch (Exception ex)
            {
                lblEstado.Text = $"Error al preparar pólizas: {ex.Message}";
                lblEstado.ForeColor = Color.Red;
            }
        }

        private string LimpiarTituloParaValidacion(string titulo)
        {
            if (string.IsNullOrWhiteSpace(titulo))
                return string.Empty;

            string limpio = titulo.ToUpper().Trim();

            limpio = limpio.Replace("-", " ")
                           .Replace("_", " ")
                           .Replace(".", " ")
                           .Replace(",", " ")
                           .Replace(";", " ")
                           .Replace(":", " ");

            while (limpio.Contains("  "))
                limpio = limpio.Replace("  ", " ");

            return limpio.Trim();
        }

        private bool CompararTitulos(string titulo1, string titulo2)
        {
            if (string.IsNullOrWhiteSpace(titulo1) || string.IsNullOrWhiteSpace(titulo2))
                return false;

            string limpio1 = LimpiarTituloParaValidacion(titulo1);
            string limpio2 = LimpiarTituloParaValidacion(titulo2);

            return string.Equals(limpio1, limpio2, StringComparison.OrdinalIgnoreCase);
        }

        private bool BuscarSimilitudes(string tituloBuscar, List<string> titulosExistentes)
        {
            if (string.IsNullOrWhiteSpace(tituloBuscar) || titulosExistentes == null || titulosExistentes.Count == 0)
                return false;

            string tituloLimpio = LimpiarTituloParaValidacion(tituloBuscar);

            foreach (string existente in titulosExistentes)
            {
                string existenteLimpio = LimpiarTituloParaValidacion(existente);

                if (tituloLimpio == existenteLimpio)
                    return true;

                if (tituloLimpio.Contains(existenteLimpio) || existenteLimpio.Contains(tituloLimpio))
                    return true;
            }

            return false;
        }

        private void LogDuplicado(int folio, string titulo, string motivo)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "duplicados.log");
                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Folio: {folio}, Título: '{titulo}', Motivo: {motivo}\n";

                File.AppendAllText(logPath, logEntry);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al escribir en log: {ex.Message}");
            }
        }

        private void PrepararPolizasDesdeDatos(DataTable datos)
        {
            _polizasGenerar.Clear();

            if (datos.Rows.Count == 0)
            {
                lblEstado.Text = "No hay datos para preparar pólizas.";
                lblEstado.ForeColor = Color.Orange;
                return;
            }

            try
            {
                int movimientosOmitidos = 0;
                StringBuilder mensajesAdvertencia = new StringBuilder();

                var gruposPorTitulo = datos.AsEnumerable()
                    .GroupBy(row => ObtenerStringSeguro(row, "titulo"))
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key));

                foreach (var grupo in gruposPorTitulo)
                {
                    var primeraFila = grupo.First();
                    string tituloDocumento = grupo.Key;

                    PolizaDTO poliza = new PolizaDTO
                    {
                        Folio = ObtenerIntSeguro(primeraFila, "Folio"),
                        Fecha = ObtenerFechaSegura(primeraFila, "DateDocument"),
                        Concepto = $"Entrada de Mercancía {tituloDocumento}",
                        Titulo = tituloDocumento,
                        Serie = ObtenerStringSeguro(primeraFila, "serie"),
                        Movimientos = new List<MovimientoPolizaDTO>()
                    };

                    int consecutivoMov = 1;
                    decimal sumaNetos = 0;

                    foreach (var fila in grupo)
                    {
                        decimal neto = ObtenerDecimalSeguro(fila, "neto");
                        sumaNetos += neto;

                        string cuentaContable = ObtenerStringSeguro(fila, "segcat1");

                        if (string.IsNullOrWhiteSpace(cuentaContable))
                        {
                            movimientosOmitidos++;
                            mensajesAdvertencia.AppendLine($"• Movimiento omitido en '{tituloDocumento}': Cuenta contable vacía");
                            continue;
                        }

                        string cuentaNumerica = FormatearCuentaParaSDK(cuentaContable);

                        var movimiento = new MovimientoPolizaDTO
                        {
                            NumMovimiento = consecutivoMov++,
                            CuentaContable = cuentaContable,
                            CuentaSDK = cuentaNumerica,
                            Monto = neto,
                            TipoMovimiento = ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_CARGO,
                            Referencia = $"{poliza.Serie}{poliza.Folio}",
                            Concepto = $"Entrada de Mercancía {tituloDocumento}"
                        };

                        poliza.Movimientos.Add(movimiento);
                    }

                    if (poliza.Movimientos.Count == 0)
                    {
                        mensajesAdvertencia.AppendLine($"• No se crearon movimientos para '{tituloDocumento}': todas las cuentas vacías");
                        continue;
                    }

                    string cuentaAbono = ObtenerCuentaContableAbono();
                    string cuentaAbonoSDK = FormatearCuentaParaSDK(cuentaAbono);

                    var abono = new MovimientoPolizaDTO
                    {
                        NumMovimiento = consecutivoMov,
                        CuentaContable = cuentaAbono,
                        CuentaSDK = cuentaAbonoSDK,
                        Monto = sumaNetos,
                        TipoMovimiento = ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_ABONO,
                        Referencia = $"{poliza.Serie}{poliza.Folio}",
                        Concepto = poliza.Concepto
                    };

                    poliza.Movimientos.Add(abono);
                    poliza.Total = sumaNetos;

                    _polizasGenerar.Add(poliza);
                }

                string estadoTexto = $"Preparadas {_polizasGenerar.Count} pólizas para generación";
                if (movimientosOmitidos > 0)
                    estadoTexto += $", {movimientosOmitidos} movimientos omitidos";

                lblEstado.Text = estadoTexto;
                lblEstado.ForeColor = _polizasGenerar.Count > 0 ? Color.Green : Color.Orange;

                if (movimientosOmitidos > 0)
                {
                    Console.WriteLine("Movimientos omitidos por cuentas vacías:");
                    Console.WriteLine(mensajesAdvertencia.ToString());
                }
            }
            catch (Exception ex)
            {
                lblEstado.Text = $"Error al preparar pólizas: {ex.Message}";
                lblEstado.ForeColor = Color.Red;
            }
        }

        private string FormatearCuentaParaSDK(string cuentaConFormato)
        {
            if (string.IsNullOrWhiteSpace(cuentaConFormato))
                return string.Empty;

            string cuentaOriginal = cuentaConFormato.Trim();

            Console.WriteLine($"DEBUG: Cuenta original: '{cuentaOriginal}'");

            if (cuentaOriginal.All(char.IsDigit))
            {
                Console.WriteLine($"DEBUG: Ya es numérica, devolviendo: '{cuentaOriginal}'");
                return cuentaOriginal;
            }

            if (cuentaOriginal.Contains("-"))
            {
                string cuentaSinGuiones = cuentaOriginal.Replace("-", "");
                Console.WriteLine($"DEBUG: Con guiones, sin guiones: '{cuentaSinGuiones}'");

                if (cuentaSinGuiones.All(char.IsDigit))
                {
                    return cuentaSinGuiones;
                }
            }

            StringBuilder soloNumeros = new StringBuilder();
            foreach (char c in cuentaOriginal)
            {
                if (char.IsDigit(c))
                    soloNumeros.Append(c);
            }

            string resultado = soloNumeros.ToString();
            Console.WriteLine($"DEBUG: Después de filtrar solo dígitos: '{resultado}'");

            return resultado;
        }

        private string ObtenerCuentaContableAbono()
        {
           // return "20103000";
            return "02105100001";
        }

        private DateTime ObtenerFechaSegura(DataRow fila, string nombreColumna)
        {
            try
            {
                if (!fila.Table.Columns.Contains(nombreColumna))
                    return DateTime.Now;

                object valor = fila[nombreColumna];
                if (valor == null || Convert.IsDBNull(valor))
                    return DateTime.Now;

                if (valor is DateTime)
                    return (DateTime)valor;

                if (valor is string fechaStr)
                {
                    if (DateTime.TryParse(fechaStr, out DateTime fecha))
                        return fecha;
                }

                return Convert.ToDateTime(valor);
            }
            catch
            {
                return DateTime.Now;
            }
        }

        private string ObtenerStringSeguro(DataRow fila, string nombreColumna)
        {
            try
            {
                if (!fila.Table.Columns.Contains(nombreColumna))
                    return string.Empty;

                object valor = fila[nombreColumna];
                if (valor == null || Convert.IsDBNull(valor))
                    return string.Empty;

                return valor.ToString().Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        private decimal ObtenerDecimalSeguro(DataRow fila, string nombreColumna)
        {
            try
            {
                if (!fila.Table.Columns.Contains(nombreColumna))
                    return 0;

                object valor = fila[nombreColumna];
                if (valor == null || Convert.IsDBNull(valor))
                    return 0;

                decimal resultado;

                if (valor is decimal)
                    resultado = (decimal)valor;
                else if (valor is double)
                    resultado = Convert.ToDecimal((double)valor);
                else if (valor is float)
                    resultado = Convert.ToDecimal((float)valor);
                else if (valor is int)
                    resultado = Convert.ToDecimal((int)valor);
                else if (valor is string strValor)
                {
                    if (decimal.TryParse(strValor, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.CurrentCulture, out decimal parseado))
                        resultado = parseado;
                    else
                        resultado = 0;
                }
                else
                    resultado = Convert.ToDecimal(valor);

                return Math.Round(resultado, 2, MidpointRounding.AwayFromZero);
            }
            catch
            {
                return 0;
            }
        }

        private int ObtenerIntSeguro(DataRow fila, string nombreColumna)
        {
            try
            {
                if (!fila.Table.Columns.Contains(nombreColumna))
                    return 0;

                object valor = fila[nombreColumna];
                if (valor == null || Convert.IsDBNull(valor))
                    return 0;

                if (valor is int)
                    return (int)valor;

                if (valor is decimal)
                    return Convert.ToInt32((decimal)valor);

                if (valor is string strValor)
                {
                    if (int.TryParse(strValor, out int resultado))
                        return resultado;
                }

                return Convert.ToInt32(valor);
            }
            catch
            {
                return 0;
            }
        }

        private decimal FormatearMontoParaSDK(decimal monto)
        {
            return Math.Round(monto, 2, MidpointRounding.AwayFromZero);
        }

        private bool ValidarRedondeoPoliza(PolizaDTO poliza)
        {
            try
            {
                foreach (var movimiento in poliza.Movimientos)
                {
                    decimal montoRedondeado = Math.Round(movimiento.Monto, 2);
                    if (Math.Abs(movimiento.Monto - montoRedondeado) > 0.0001m)
                    {
                        movimiento.Monto = montoRedondeado;
                        Console.WriteLine($"Ajustado redondeo en movimiento {movimiento.NumMovimiento}: {movimiento.Monto:C}");
                    }
                }

                decimal totalCalculado = poliza.Movimientos
                    .Where(m => m.TipoMovimiento == ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_CARGO)
                    .Sum(m => m.Monto);

                decimal totalRedondeado = Math.Round(totalCalculado, 2);

                if (Math.Abs(poliza.Total - totalRedondeado) > 0.0001m)
                {
                    poliza.Total = totalRedondeado;
                    Console.WriteLine($"Ajustado redondeo del total: {poliza.Total:C}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error en validación de redondeo: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Generación de Pólizas en CONTPAQi

        private void btnGenerarPolizas_Click(object sender, EventArgs e)
        {
            if (_polizasGenerar.Count == 0)
            {
                lblEstado.Text = "No hay pólizas preparadas para generar.";
                lblEstado.ForeColor = Color.Orange;
                return;
            }

            if (string.IsNullOrEmpty(_empresaActual))
            {
                lblEstado.Text = "No hay empresa de CONTPAQi abierta.";
                lblEstado.ForeColor = Color.Orange;
                return;
            }

            List<PolizaDTO> polizasExitosas = new List<PolizaDTO>();
            List<PolizaDTO> polizasFallidas = new List<PolizaDTO>();

            try
            {
                progressBar.Visible = true;
                progressBar.Maximum = _polizasGenerar.Count;
                progressBar.Value = 0;
                btnGenerarPolizas.Enabled = false;
                Application.DoEvents();

                int pólizasCreadas = 0;
                int errores = 0;
                List<string> erroresDetallados = new List<string>();

                foreach (var poliza in _polizasGenerar)
                {
                    lblEstado.Text = $"Creando póliza para {poliza.Serie}{poliza.Folio} - {poliza.Titulo}...";
                    Application.DoEvents();

                    if (!ValidarRedondeoPoliza(poliza))
                    {
                        errores++;
                        polizasFallidas.Add(poliza);
                        erroresDetallados.Add($"Folio {poliza.Folio}: Error en validación de redondeo");
                        continue;
                    }

                    bool resultado = CrearPolizaContpaqi(poliza, out string mensajeError);

                    if (resultado)
                    {
                        pólizasCreadas++;
                        polizasExitosas.Add(poliza);
                    }
                    else
                    {
                        errores++;
                        polizasFallidas.Add(poliza);
                        erroresDetallados.Add($"Folio {poliza.Folio}: {mensajeError}");
                    }

                    progressBar.Value++;
                    Application.DoEvents();
                }

                progressBar.Visible = false;
                btnGenerarPolizas.Enabled = true;

                // ===== GENERAR BITÁCORA EN EXCEL =====
                if (polizasExitosas.Count > 0 || polizasFallidas.Count > 0)
                {
                    GenerarBitacoraExcel(_datosOriginales, polizasExitosas, polizasFallidas);
                }

                lblEstado.Text = $"Proceso completado: {pólizasCreadas} creadas, {errores} errores";
                lblEstado.ForeColor = errores == 0 ? Color.Green : Color.Orange;

                if (erroresDetallados.Count > 0)
                {
                    Console.WriteLine("═══════════════════════════════════════");
                    Console.WriteLine("ERRORES DETALLADOS EN GENERACIÓN DE PÓLIZAS:");
                    foreach (string error in erroresDetallados)
                    {
                        Console.WriteLine(error);
                    }
                }
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error en generación de pólizas";
                lblEstado.ForeColor = Color.Red;
            }
            finally
            {
                progressBar.Visible = false;
                btnGenerarPolizas.Enabled = true;
            }
        }

        private bool CrearPolizaContpaqi(PolizaDTO poliza, out string mensajeError)
        {
            mensajeError = string.Empty;

            try
            {
                var sdkPoliza = new TSdkPoliza();
                sdkPoliza.setSesion(ses);
                sdkPoliza.iniciarInfo();

                sdkPoliza.Tipo = ETIPOPOLIZA.TIPO_DIARIO;
                sdkPoliza.Clase = ECLASEPOLIZA.CLASE_AFECTAR;
                sdkPoliza.Impresa = 0;
                sdkPoliza.Fecha = poliza.Fecha;
                sdkPoliza.SistOrigen = ESISTORIGEN.ORIG_CONTPAQNG;
                sdkPoliza.Concepto = poliza.Concepto.Length > 60 ?
                    poliza.Concepto.Substring(0, 60) : poliza.Concepto;
                sdkPoliza.Guid = Guid.NewGuid().ToString().ToUpper();

                foreach (var movimiento in poliza.Movimientos)
                {
                    var sdkMovimiento = new TSdkMovimientoPoliza();
                    sdkMovimiento.setSesion(ses);
                    sdkMovimiento.iniciarInfo();

                    sdkMovimiento.NumMovto = movimiento.NumMovimiento;
                    sdkMovimiento.CodigoCuenta = movimiento.CuentaSDK;
                    sdkMovimiento.TipoMovto = movimiento.TipoMovimiento;
                    sdkMovimiento.Importe = movimiento.Monto;
                    sdkMovimiento.SegmentoNegocio = "1005";

                    if (!string.IsNullOrEmpty(movimiento.Referencia))
                        sdkMovimiento.Referencia = movimiento.Referencia;

                    if (!string.IsNullOrEmpty(movimiento.Concepto) && movimiento.Concepto.Length <= 60)
                        sdkMovimiento.Concepto = movimiento.Concepto;

                    sdkMovimiento.Guid = Guid.NewGuid().ToString().ToUpper();

                    int resultado = sdkPoliza.agregaMovimiento(sdkMovimiento);
                    if (resultado == 0)
                    {
                        string errorSDK = sdkMovimiento.getMensajeError();
                        mensajeError = $"Error en movimiento {movimiento.NumMovimiento}: " +
                                     $"Cuenta: {movimiento.CuentaContable} -> {movimiento.CuentaSDK}, " +
                                     $"Monto: {movimiento.Monto:C}, " +
                                     $"SDK: {errorSDK}";
                        return false;
                    }
                }

                int ret = sdkPoliza.crea();
                if (ret == 0)
                {
                    mensajeError = $"Error SDK: {sdkPoliza.getMensajeError()}";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                mensajeError = $"Excepción: {ex.Message}";
                return false;
            }
        }

        #endregion

        #region Bitácora Excel

        private void GenerarBitacoraExcel(DataTable datosOriginales, List<PolizaDTO> polizasExitosas, List<PolizaDTO> polizasFallidas)
        {
            try
            {
                lblEstado.Text = "Generando bitácora en Excel...";
                Application.DoEvents();

                string fechaArchivo = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string nombreArchivo = $"Bitacora_RC_{fechaArchivo}.xlsx";
                string rutaCompleta = Path.Combine(rutaBitacoraBase, nombreArchivo);

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                // Configurar título
                worksheet.Cells[1, 1] = "BITÁCORA DE GENERACIÓN DE PÓLIZAS - RECEPCIÓN DE MERCANCÍA";
                worksheet.Cells[2, 1] = $"Fecha de generación: {DateTime.Now:dd/MM/yyyy HH:mm:ss}";
                worksheet.Cells[3, 1] = $"Empresa CONTPAQi: {_empresaActual}";
                worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 6]].Merge();
                worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 6]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 6]].Font.Size = 14;

                int filaActual = 5;

                // ===== SECCIÓN 1: RESUMEN =====
                worksheet.Cells[filaActual, 1] = "RESUMEN DE PROCESO";
                worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 6]].Font.Bold = true;
                worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 6]].Font.Size = 12;
                filaActual += 2;

                worksheet.Cells[filaActual, 1] = "Total documentos originales:";
                worksheet.Cells[filaActual, 2] = datosOriginales?.Rows.Count ?? 0;
                worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 2]].Font.Bold = true;
                filaActual++;

                worksheet.Cells[filaActual, 1] = "Total pólizas generadas:";
                worksheet.Cells[filaActual, 2] = polizasExitosas.Count;
                worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 2]].Font.Bold = true;
                filaActual++;

                worksheet.Cells[filaActual, 1] = "Total pólizas con error:";
                worksheet.Cells[filaActual, 2] = polizasFallidas.Count;
                worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 2]].Font.Bold = true;
                filaActual += 2;

                // ===== SECCIÓN 2: PÓLIZAS EXITOSAS =====
                if (polizasExitosas.Count > 0)
                {
                    worksheet.Cells[filaActual, 1] = "PÓLIZAS GENERADAS EXITOSAMENTE";
                    worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 6]].Font.Bold = true;
                    worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 6]].Interior.Color = Color.LightGreen;
                    filaActual++;

                    // Encabezados
                    string[] encabezadosExitosas = { "Serie", "Folio", "Título", "Fecha", "Total", "Concepto" };
                    for (int i = 0; i < encabezadosExitosas.Length; i++)
                    {
                        worksheet.Cells[filaActual, i + 1] = encabezadosExitosas[i];
                        worksheet.Cells[filaActual, i + 1].Font.Bold = true;
                        worksheet.Cells[filaActual, i + 1].Interior.Color = Color.LightGray;
                    }
                    filaActual++;

                    foreach (var poliza in polizasExitosas)
                    {
                        worksheet.Cells[filaActual, 1] = poliza.Serie;
                        worksheet.Cells[filaActual, 2] = poliza.Folio;
                        worksheet.Cells[filaActual, 3] = poliza.Titulo;
                        worksheet.Cells[filaActual, 4] = poliza.Fecha.ToString("dd/MM/yyyy");
                        worksheet.Cells[filaActual, 5] = poliza.Total;
                        worksheet.Cells[filaActual, 6] = poliza.Concepto;
                        filaActual++;
                    }
                    filaActual += 2;
                }

                // ===== SECCIÓN 3: PÓLIZAS CON ERROR =====
                if (polizasFallidas.Count > 0)
                {
                    worksheet.Cells[filaActual, 1] = "PÓLIZAS CON ERRORES";
                    worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 6]].Font.Bold = true;
                    worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 6]].Interior.Color = Color.LightCoral;
                    filaActual++;

                    string[] encabezadosError = { "Serie", "Folio", "Título", "Fecha", "Total", "Estado" };
                    for (int i = 0; i < encabezadosError.Length; i++)
                    {
                        worksheet.Cells[filaActual, i + 1] = encabezadosError[i];
                        worksheet.Cells[filaActual, i + 1].Font.Bold = true;
                        worksheet.Cells[filaActual, i + 1].Interior.Color = Color.LightGray;
                    }
                    filaActual++;

                    foreach (var poliza in polizasFallidas)
                    {
                        worksheet.Cells[filaActual, 1] = poliza.Serie;
                        worksheet.Cells[filaActual, 2] = poliza.Folio;
                        worksheet.Cells[filaActual, 3] = poliza.Titulo;
                        worksheet.Cells[filaActual, 4] = poliza.Fecha.ToString("dd/MM/yyyy");
                        worksheet.Cells[filaActual, 5] = poliza.Total;
                        worksheet.Cells[filaActual, 6] = "ERROR";
                        filaActual++;
                    }
                    filaActual += 2;
                }

                // ===== SECCIÓN 4: DETALLE DE MOVIMIENTOS =====
                if (polizasExitosas.Count > 0)
                {
                    worksheet.Cells[filaActual, 1] = "DETALLE DE MOVIMIENTOS - PÓLIZAS EXITOSAS";
                    worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 8]].Font.Bold = true;
                    worksheet.Range[worksheet.Cells[filaActual, 1], worksheet.Cells[filaActual, 8]].Interior.Color = Color.LightBlue;
                    filaActual++;

                    string[] encabezadosMov = { "Serie", "Folio", "Título", "Mov.", "Cuenta", "Monto", "Tipo" };
                    for (int i = 0; i < encabezadosMov.Length; i++)
                    {
                        worksheet.Cells[filaActual, i + 1] = encabezadosMov[i];
                        worksheet.Cells[filaActual, i + 1].Font.Bold = true;
                        worksheet.Cells[filaActual, i + 1].Interior.Color = Color.LightGray;
                    }
                    filaActual++;

                    foreach (var poliza in polizasExitosas)
                    {
                        foreach (var mov in poliza.Movimientos)
                        {
                            worksheet.Cells[filaActual, 1] = poliza.Serie;
                            worksheet.Cells[filaActual, 2] = poliza.Folio;
                            worksheet.Cells[filaActual, 3] = poliza.Titulo;
                            worksheet.Cells[filaActual, 4] = mov.NumMovimiento;
                            worksheet.Cells[filaActual, 5] = mov.CuentaContable;
                            worksheet.Cells[filaActual, 6] = mov.Monto;
                            worksheet.Cells[filaActual, 7] = mov.TipoMovimiento == ETIPOIMPORTEMOVPOLIZA.MOVPOLIZA_CARGO ? "CARGO" : "ABONO";
                            filaActual++;
                        }
                    }
                }

                // Ajustar columnas
                worksheet.Columns.AutoFit();

                // Guardar archivo
                workbook.SaveAs(rutaCompleta);
                workbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                // Actualizar el TextBox con la ruta del archivo generado
                TBBitacora.Text = rutaCompleta;

                // Mostrar mensaje detallado en el panel de estado
                lblEstado.Text = $"✓ Bitácora Excel creada exitosamente: {nombreArchivo}";
                lblEstado.ForeColor = Color.Green;

                // Mostrar también en el toolstrip status
                toolStripStatusLabel2.Text = $"Bitácora: {nombreArchivo}";
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error al generar bitácora Excel. Generando respaldo...";
                lblEstado.ForeColor = Color.Orange;

                // Guardar respaldo en texto plano
                GenerarBitacoraTexto(datosOriginales, polizasExitosas, polizasFallidas);
            }
        }

        private void GenerarBitacoraTexto(DataTable datosOriginales, List<PolizaDTO> polizasExitosas, List<PolizaDTO> polizasFallidas)
        {
            try
            {
                string fechaArchivo = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string nombreArchivo = $"Bitacora_RC_{fechaArchivo}.txt";
                string rutaCompleta = Path.Combine(rutaBitacoraBase, nombreArchivo);

                using (StreamWriter writer = new StreamWriter(rutaCompleta))
                {
                    writer.WriteLine("==========================================");
                    writer.WriteLine("BITÁCORA DE GENERACIÓN DE PÓLIZAS - RECEPCIÓN DE MERCANCÍA");
                    writer.WriteLine($"Fecha: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                    writer.WriteLine($"Empresa CONTPAQi: {_empresaActual}");
                    writer.WriteLine("==========================================");
                    writer.WriteLine();
                    writer.WriteLine($"Total documentos originales: {datosOriginales?.Rows.Count ?? 0}");
                    writer.WriteLine($"Total pólizas generadas: {polizasExitosas.Count}");
                    writer.WriteLine($"Total pólizas con error: {polizasFallidas.Count}");
                    writer.WriteLine();

                    writer.WriteLine("=== PÓLIZAS EXITOSAS ===");
                    foreach (var poliza in polizasExitosas)
                    {
                        writer.WriteLine($"{poliza.Serie}{poliza.Folio} - {poliza.Titulo} - ${poliza.Total:C}");
                    }
                    writer.WriteLine();

                    writer.WriteLine("=== PÓLIZAS CON ERROR ===");
                    foreach (var poliza in polizasFallidas)
                    {
                        writer.WriteLine($"{poliza.Serie}{poliza.Folio} - {poliza.Titulo} - ${poliza.Total:C}");
                    }
                }

                TBBitacora.Text = rutaCompleta;

                // Mostrar mensaje detallado en el panel de estado
                lblEstado.Text = $"✓ Bitácora creada exitosamente: {nombreArchivo} en {rutaBitacoraBase}";
                lblEstado.ForeColor = Color.Green;

                // Opcional: Mostrar también en el toolstrip status
                toolStripStatusLabel2.Text = $"Última bitácora: {nombreArchivo}";

                // Registrar en el log del sistema (opcional)
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Bitácora generada: {rutaCompleta}");
            }
            catch (Exception ex)
            {
                lblEstado.Text = "✗ Error al guardar bitácora de texto";
                lblEstado.ForeColor = Color.Red;
                Console.WriteLine($"Error al generar bitácora: {ex.Message}");
            }
        }

        #endregion

        #region Métodos Auxiliares Mejorados

        private void btnVerDuplicados_Click(object sender, EventArgs e)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "duplicados.log");

                if (File.Exists(logPath))
                {
                    string contenido = File.ReadAllText(logPath);

                    if (string.IsNullOrWhiteSpace(contenido))
                    {
                        lblEstado.Text = "No hay registros de duplicados";
                        lblEstado.ForeColor = Color.Black;
                        return;
                    }

                    Form formLog = new Form();
                    formLog.Text = "Registro de Pólizas Duplicadas";
                    formLog.Size = new Size(600, 400);

                    TextBox textBox = new TextBox();
                    textBox.Multiline = true;
                    textBox.ScrollBars = ScrollBars.Both;
                    textBox.Dock = DockStyle.Fill;
                    textBox.Text = contenido;
                    textBox.ReadOnly = true;
                    textBox.Font = new Font("Consolas", 9);

                    formLog.Controls.Add(textBox);
                    formLog.ShowDialog();
                }
                else
                {
                    lblEstado.Text = "No se encontró archivo de duplicados";
                    lblEstado.ForeColor = Color.Black;
                }
            }
            catch (Exception ex)
            {
                lblEstado.Text = "Error al leer el log";
                lblEstado.ForeColor = Color.Red;
            }
        }

        #endregion

        #region Clases DTO

        public class PolizaDTO
        {
            public int Folio { get; set; }
            public DateTime Fecha { get; set; }
            public string Concepto { get; set; }
            public string Titulo { get; set; }
            public string Serie { get; set; }
            public decimal Total { get; set; }
            public List<MovimientoPolizaDTO> Movimientos { get; set; }
        }

        public class MovimientoPolizaDTO
        {
            public int NumMovimiento { get; set; }
            public string CuentaContable { get; set; }
            public string CuentaSDK { get; set; }
            public decimal Monto { get; set; }
            public ETIPOIMPORTEMOVPOLIZA TipoMovimiento { get; set; }
            public string Referencia { get; set; }
            public string Concepto { get; set; }
        }

        #endregion

        #region Métodos Auxiliares y Eventos

        private void ManejarErrorSQL(SqlException ex)
        {
            string mensajeError = $"Error SQL (Código: {ex.Number})";

            switch (ex.Number)
            {
                case 18456: mensajeError += ": Error de autenticación"; break;
                case 4060: mensajeError += ": No se puede abrir la base de datos"; break;
                case 53: case -1: mensajeError += ": No se puede conectar al servidor"; break;
                default: mensajeError += $": {ex.Message}"; break;
            }

            lblEstado.Text = mensajeError;
            lblEstado.ForeColor = Color.Red;
        }

        private void FormatearColumnasDataGridView()
        {
            if (dgvResultados.Columns.Count == 0) return;

            if (dgvResultados.Columns.Contains("neto"))
            {
                dgvResultados.Columns["neto"].HeaderText = "Neto";
                dgvResultados.Columns["neto"].DefaultCellStyle.Format = "C2";
                dgvResultados.Columns["neto"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                dgvResultados.Columns["neto"].Width = 100;
            }

            if (dgvResultados.Columns.Contains("segcat1"))
            {
                dgvResultados.Columns["segcat1"].HeaderText = "Segmento Contable (Cuenta)";
                dgvResultados.Columns["segcat1"].Width = 180;
                dgvResultados.Columns["segcat1"].DefaultCellStyle.Font = new Font("Consolas", 9);
            }

            if (dgvResultados.Columns.Contains("titulo"))
            {
                dgvResultados.Columns["titulo"].HeaderText = "Título del Documento";
                dgvResultados.Columns["titulo"].Width = 250;
                dgvResultados.Columns["titulo"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            }

            if (dgvResultados.Columns.Contains("serie"))
            {
                dgvResultados.Columns["serie"].HeaderText = "Serie";
                dgvResultados.Columns["serie"].Width = 80;
                dgvResultados.Columns["serie"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            if (dgvResultados.Columns.Contains("Folio"))
            {
                dgvResultados.Columns["Folio"].HeaderText = "Folio";
                dgvResultados.Columns["Folio"].Width = 80;
                dgvResultados.Columns["Folio"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            if (dgvResultados.Columns.Contains("DateDocument"))
            {
                dgvResultados.Columns["DateDocument"].HeaderText = "Fecha Documento";
                dgvResultados.Columns["DateDocument"].Width = 120;
                dgvResultados.Columns["DateDocument"].DefaultCellStyle.Format = "dd/MM/yyyy";
                dgvResultados.Columns["DateDocument"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            dgvResultados.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders);
            dgvResultados.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
        }

        #endregion

        #region Eventos de UI

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                var builder = new SqlConnectionStringBuilder(connectionString);
                if (toolStripStatusLabel1 != null)
                {
                    toolStripStatusLabel1.Text = $"Conectado a: {builder.InitialCatalog} en {builder.DataSource}";
                }
            }
            catch
            {
                if (toolStripStatusLabel1 != null)
                {
                    toolStripStatusLabel1.Text = "Cadena de conexión cargada desde app.config";
                }
            }
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            dgvResultados.DataSource = null;
            dgvResultados.Rows.Clear();
            dgvResultados.Columns.Clear();
            _polizasGenerar.Clear();
            _datosOriginales = null;
            lblEstado.Text = "Listo para cargar datos";
            lblEstado.ForeColor = Color.Black;
            btnGenerarPolizas.Enabled = false;
            TBBitacora.Text = rutaBitacoraBase;
        }

        private void dgvResultados_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            DataGridViewRow filaSeleccionada = dgvResultados.Rows[e.RowIndex];
            string mensaje = $"Detalles del registro:\n\n" +
                            $"Neto: {filaSeleccionada.Cells["neto"]?.Value ?? "N/A"}\n" +
                            $"Segmento Contable (Cuenta): {filaSeleccionada.Cells["segcat1"]?.Value ?? "N/A"}\n" +
                            $"Título: {filaSeleccionada.Cells["titulo"]?.Value ?? "N/A"}\n" +
                            $"Serie: {filaSeleccionada.Cells["serie"]?.Value ?? "N/A"}\n" +
                            $"Folio: {filaSeleccionada.Cells["Folio"]?.Value ?? "N/A"}\n" +
                            $"Fecha: {filaSeleccionada.Cells["DateDocument"]?.Value ?? "N/A"}";

            MessageBox.Show(mensaje, "Detalles del Registro",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion
    }
}