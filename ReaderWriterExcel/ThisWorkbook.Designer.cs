﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан программой.
//     Исполняемая версия:4.0.30319.42000
//
//     Изменения в этом файле могут привести к неправильной работе и будут потеряны в случае
//     повторной генерации кода.
// </auto-generated>
//------------------------------------------------------------------------------

#pragma warning disable 414
namespace ReaderWriterExcel {
    
    
    /// 
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [global::System.Security.Permissions.PermissionSetAttribute(global::System.Security.Permissions.SecurityAction.Demand, Name="FullTrust")]
    public sealed partial class ThisWorkbook : Microsoft.Office.Tools.Excel.WorkbookBase {
        
        internal Microsoft.Office.Tools.ActionsPane ActionsPane;
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        private global::System.Object missing = global::System.Type.Missing;
        
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        internal Microsoft.Office.Interop.Excel.Application ThisApplication;
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        public ThisWorkbook(global::Microsoft.Office.Tools.Excel.Factory factory, global::System.IServiceProvider serviceProvider) : 
                base(factory, serviceProvider, "ThisWorkbook", "ThisWorkbook") {
            Globals.Factory = factory;
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void Initialize() {
            base.Initialize();
            this.ThisApplication = this.GetHostItem<Microsoft.Office.Interop.Excel.Application>(typeof(Microsoft.Office.Interop.Excel.Application), "Application");
            Globals.ThisWorkbook = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            this.InitializeCachedData();
            this.InitializeControls();
            this.InitializeComponents();
            this.InitializeData();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void FinishInitialization() {
            this.InternalStartup();
            this.OnStartup();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void InitializeDataBindings() {
            this.BeginInitialization();
            this.BindToData();
            this.EndInitialization();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeCachedData() {
            if ((this.DataHost == null)) {
                return;
            }
            if (this.DataHost.IsCacheInitialized) {
                this.DataHost.FillCachedData(this);
            }
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BindToData() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StartCaching(string MemberName) {
            this.DataHost.StartCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private void StopCaching(string MemberName) {
            this.DataHost.StopCaching(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool IsCached(string MemberName) {
            return this.DataHost.IsCached(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void BeginInitialization() {
            this.BeginInit();
            this.ActionsPane.BeginInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void EndInitialization() {
            this.ActionsPane.EndInit();
            this.EndInit();
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeControls() {
            this.ActionsPane = Globals.Factory.CreateActionsPane(null, null, "ActionsPane", "ActionsPane", this);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        private void InitializeComponents() {
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        private bool NeedsFill(string MemberName) {
            return this.DataHost.NeedsFill(this, MemberName);
        }
        
        /// 
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Never)]
        protected override void OnShutdown() {
            this.ActionsPane.Dispose();
            base.OnShutdown();
        }
    }
    
    /// 
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
    internal sealed partial class Globals {
        
        /// 
        private Globals() {
        }
        
        private static ThisWorkbook _ThisWorkbook;
        
        private static global::Microsoft.Office.Tools.Excel.Factory _factory;
        
        private static ThisRibbonCollection _ThisRibbonCollection;
        
        internal static ThisWorkbook ThisWorkbook {
            get {
                return _ThisWorkbook;
            }
            set {
                if ((_ThisWorkbook == null)) {
                    _ThisWorkbook = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
        
        internal static global::Microsoft.Office.Tools.Excel.Factory Factory {
            get {
                return _factory;
            }
            set {
                if ((_factory == null)) {
                    _factory = value;
                }
                else {
                    throw new System.NotSupportedException();
                }
            }
        }
        
        internal static ThisRibbonCollection Ribbons {
            get {
                if ((_ThisRibbonCollection == null)) {
                    _ThisRibbonCollection = new ThisRibbonCollection(_factory.GetRibbonFactory());
                }
                return _ThisRibbonCollection;
            }
        }
    }
    
    /// 
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
    internal sealed partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonCollectionBase {
        
        /// 
        internal ThisRibbonCollection(global::Microsoft.Office.Tools.Ribbon.RibbonFactory factory) : 
                base(factory) {
        }
    }
}
