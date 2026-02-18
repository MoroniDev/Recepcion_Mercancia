namespace Recepcion_Mercancia
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.dgvResultados = new System.Windows.Forms.DataGridView();
            this.panelSuperior = new System.Windows.Forms.Panel();
            this.lblTitulo = new System.Windows.Forms.Label();
            this.panelControles = new System.Windows.Forms.Panel();
            this.lblEstadoSDK = new System.Windows.Forms.Label();
            this.LabelTextoRutaBitacora = new System.Windows.Forms.Label();
            this.BtnAutomatico = new System.Windows.Forms.Button();
            this.TBBitacora = new System.Windows.Forms.TextBox();
            this.btnGenerarPolizas = new System.Windows.Forms.Button();
            this.btnLimpiar = new System.Windows.Forms.Button();
            this.btnCargar = new System.Windows.Forms.Button();
            this.btnConfigurar = new System.Windows.Forms.Button();
            this.panelEstado = new System.Windows.Forms.Panel();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblEstado = new System.Windows.Forms.Label();
            this.panelInferior = new System.Windows.Forms.Panel();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResultados)).BeginInit();
            this.panelControles.SuspendLayout();
            this.panelEstado.SuspendLayout();
            this.panelInferior.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvResultados
            // 
            this.dgvResultados.AllowUserToAddRows = false;
            this.dgvResultados.AllowUserToDeleteRows = false;
            this.dgvResultados.AllowUserToOrderColumns = true;
            this.dgvResultados.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvResultados.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvResultados.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgvResultados.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvResultados.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvResultados.Location = new System.Drawing.Point(0, 40);
            this.dgvResultados.Name = "dgvResultados";
            this.dgvResultados.ReadOnly = true;
            this.dgvResultados.RowHeadersWidth = 51;
            this.dgvResultados.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvResultados.Size = new System.Drawing.Size(1273, 364);
            this.dgvResultados.TabIndex = 0;
            this.dgvResultados.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvResultados_CellDoubleClick);
            // 
            // panelSuperior
            // 
            this.panelSuperior.Location = new System.Drawing.Point(0, 0);
            this.panelSuperior.Name = "panelSuperior";
            this.panelSuperior.Size = new System.Drawing.Size(200, 100);
            this.panelSuperior.TabIndex = 6;
            // 
            // lblTitulo
            // 
            this.lblTitulo.Location = new System.Drawing.Point(0, 0);
            this.lblTitulo.Name = "lblTitulo";
            this.lblTitulo.Size = new System.Drawing.Size(100, 23);
            this.lblTitulo.TabIndex = 0;
            // 
            // panelControles
            // 
            this.panelControles.BackColor = System.Drawing.SystemColors.Control;
            this.panelControles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelControles.Controls.Add(this.lblEstadoSDK);
            this.panelControles.Controls.Add(this.LabelTextoRutaBitacora);
            this.panelControles.Controls.Add(this.BtnAutomatico);
            this.panelControles.Controls.Add(this.TBBitacora);
            this.panelControles.Controls.Add(this.btnGenerarPolizas);
            this.panelControles.Controls.Add(this.btnLimpiar);
            this.panelControles.Controls.Add(this.btnCargar);
            this.panelControles.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelControles.Location = new System.Drawing.Point(0, 0);
            this.panelControles.Name = "panelControles";
            this.panelControles.Size = new System.Drawing.Size(1273, 40);
            this.panelControles.TabIndex = 2;
            // 
            // lblEstadoSDK
            // 
            this.lblEstadoSDK.AutoSize = true;
            this.lblEstadoSDK.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEstadoSDK.Location = new System.Drawing.Point(874, 14);
            this.lblEstadoSDK.Name = "lblEstadoSDK";
            this.lblEstadoSDK.Size = new System.Drawing.Size(164, 13);
            this.lblEstadoSDK.TabIndex = 7;
            this.lblEstadoSDK.Text = "Conectando a CONTPAQi...";
            // 
            // LabelTextoRutaBitacora
            // 
            this.LabelTextoRutaBitacora.AutoSize = true;
            this.LabelTextoRutaBitacora.Location = new System.Drawing.Point(736, 14);
            this.LabelTextoRutaBitacora.Name = "LabelTextoRutaBitacora";
            this.LabelTextoRutaBitacora.Size = new System.Drawing.Size(98, 13);
            this.LabelTextoRutaBitacora.TabIndex = 5;
            this.LabelTextoRutaBitacora.Text = "Ruta de la Bitacora";
            // 
            // BtnAutomatico
            // 
            this.BtnAutomatico.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnAutomatico.Location = new System.Drawing.Point(342, 5);
            this.BtnAutomatico.Name = "BtnAutomatico";
            this.BtnAutomatico.Size = new System.Drawing.Size(96, 28);
            this.BtnAutomatico.TabIndex = 6;
            this.BtnAutomatico.Text = "Automatico";
            this.BtnAutomatico.UseVisualStyleBackColor = true;
            // 
            // TBBitacora
            // 
            this.TBBitacora.Location = new System.Drawing.Point(444, 10);
            this.TBBitacora.Name = "TBBitacora";
            this.TBBitacora.Size = new System.Drawing.Size(286, 20);
            this.TBBitacora.TabIndex = 5;
            // 
            // btnGenerarPolizas
            // 
            this.btnGenerarPolizas.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.btnGenerarPolizas.Enabled = false;
            this.btnGenerarPolizas.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnGenerarPolizas.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
            this.btnGenerarPolizas.Location = new System.Drawing.Point(216, 5);
            this.btnGenerarPolizas.Name = "btnGenerarPolizas";
            this.btnGenerarPolizas.Size = new System.Drawing.Size(120, 28);
            this.btnGenerarPolizas.TabIndex = 4;
            this.btnGenerarPolizas.Text = "Generar Pólizas";
            this.btnGenerarPolizas.UseVisualStyleBackColor = false;
            this.btnGenerarPolizas.Click += new System.EventHandler(this.btnGenerarPolizas_Click);
            // 
            // btnLimpiar
            // 
            this.btnLimpiar.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnLimpiar.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnLimpiar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.btnLimpiar.Location = new System.Drawing.Point(116, 6);
            this.btnLimpiar.Name = "btnLimpiar";
            this.btnLimpiar.Size = new System.Drawing.Size(94, 28);
            this.btnLimpiar.TabIndex = 1;
            this.btnLimpiar.Text = "Limpiar";
            this.btnLimpiar.UseVisualStyleBackColor = false;
            this.btnLimpiar.Click += new System.EventHandler(this.btnLimpiar_Click);
            // 
            // btnCargar
            // 
            this.btnCargar.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnCargar.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.btnCargar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.btnCargar.Location = new System.Drawing.Point(15, 6);
            this.btnCargar.Name = "btnCargar";
            this.btnCargar.Size = new System.Drawing.Size(94, 28);
            this.btnCargar.TabIndex = 0;
            this.btnCargar.Text = "Cargar Datos";
            this.btnCargar.UseVisualStyleBackColor = false;
            this.btnCargar.Click += new System.EventHandler(this.btnCargar_Click);
            // 
            // btnConfigurar
            // 
            this.btnConfigurar.Location = new System.Drawing.Point(708, 228);
            this.btnConfigurar.Name = "btnConfigurar";
            this.btnConfigurar.Size = new System.Drawing.Size(75, 23);
            this.btnConfigurar.TabIndex = 5;
            // 
            // panelEstado
            // 
            this.panelEstado.BackColor = System.Drawing.SystemColors.Control;
            this.panelEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelEstado.Controls.Add(this.progressBar);
            this.panelEstado.Controls.Add(this.lblEstado);
            this.panelEstado.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelEstado.Location = new System.Drawing.Point(0, 404);
            this.panelEstado.Name = "panelEstado";
            this.panelEstado.Size = new System.Drawing.Size(1273, 35);
            this.panelEstado.TabIndex = 3;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(1056, 9);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(205, 17);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar.TabIndex = 1;
            this.progressBar.Visible = false;
            // 
            // lblEstado
            // 
            this.lblEstado.AutoSize = true;
            this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.lblEstado.Location = new System.Drawing.Point(12, 9);
            this.lblEstado.Name = "lblEstado";
            this.lblEstado.Size = new System.Drawing.Size(108, 15);
            this.lblEstado.TabIndex = 0;
            this.lblEstado.Text = "Listo para cargar...";
            // 
            // panelInferior
            // 
            this.panelInferior.Controls.Add(this.statusStrip1);
            this.panelInferior.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelInferior.Location = new System.Drawing.Point(0, 439);
            this.panelInferior.Name = "panelInferior";
            this.panelInferior.Size = new System.Drawing.Size(1273, 26);
            this.panelInferior.TabIndex = 4;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2});
            this.statusStrip1.Location = new System.Drawing.Point(0, 4);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1273, 22);
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(1184, 17);
            this.toolStripStatusLabel1.Spring = true;
            this.toolStripStatusLabel1.Text = "Sistema de Consulta SQL - ComercialSP";
            this.toolStripStatusLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(74, 17);
            this.toolStripStatusLabel2.Text = "Conectado...";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(1273, 465);
            this.Controls.Add(this.dgvResultados);
            this.Controls.Add(this.btnConfigurar);
            this.Controls.Add(this.panelEstado);
            this.Controls.Add(this.panelControles);
            this.Controls.Add(this.panelSuperior);
            this.Controls.Add(this.panelInferior);
            this.MinimumSize = new System.Drawing.Size(859, 504);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Generación de Pólizas - Sistema ComercialSP";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvResultados)).EndInit();
            this.panelControles.ResumeLayout(false);
            this.panelControles.PerformLayout();
            this.panelEstado.ResumeLayout(false);
            this.panelEstado.PerformLayout();
            this.panelInferior.ResumeLayout(false);
            this.panelInferior.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvResultados;
        private System.Windows.Forms.Panel panelSuperior;
        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.Panel panelControles;
        private System.Windows.Forms.Button btnConfigurar;
        private System.Windows.Forms.Button btnLimpiar;
        private System.Windows.Forms.Button btnCargar;
        private System.Windows.Forms.Panel panelEstado;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblEstado;
        private System.Windows.Forms.Panel panelInferior;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.Button btnGenerarPolizas;
        private System.Windows.Forms.Button BtnAutomatico;
        private System.Windows.Forms.TextBox TBBitacora;
        private System.Windows.Forms.Label LabelTextoRutaBitacora;
        // 
        // NUEVO: Declaración del control lblEstadoSDK
        // 
        private System.Windows.Forms.Label lblEstadoSDK;
    }
}