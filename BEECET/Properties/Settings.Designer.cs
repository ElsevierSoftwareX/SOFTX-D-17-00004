﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AnalyticalSupportData_Info.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "11.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=.\\MSSQLSERVERR2;AttachDbFilename=|DataDirectory|\\SteelSections.mdf;In" +
            "tegrated Security=True;User Instance=True")]
        public string SteelSectionsConnectionString {
            get {
                return ((string)(this["SteelSectionsConnectionString"]));
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("RBP100")]
        public string PID {
            get {
                return ((string)(this["PID"]));
            }
            set {
                this["PID"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("1")]
        public string DON {
            get {
                return ((string)(this["DON"]));
            }
            set {
                this["DON"] = value;
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=|DataDirectory|\\Database1.sdf")]
        public string Database1ConnectionString {
            get {
                return ((string)(this["Database1ConnectionString"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=.\\MSSQLSERVERR2;AttachDbFilename=|DataDirectory|\\SteelSections.mdf;In" +
            "tegrated Security=True;Connect Timeout=30;User Instance=True")]
        public string SteelSectionsConnectionString1 {
            get {
                return ((string)(this["SteelSectionsConnectionString1"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=.\\MSSQLSERVERR2;AttachDbFilename=|DataDirectory|\\Database1.mdf;Integr" +
            "ated Security=True;User Instance=True")]
        public string Database1ConnectionString1 {
            get {
                return ((string)(this["Database1ConnectionString1"]));
            }
        }
        
        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.ConnectionString)]
        [global::System.Configuration.DefaultSettingValueAttribute("Data Source=.\\MSSQLSERVERR2;AttachDbFilename=|DataDirectory|\\ICEDatabase.mdf;Inte" +
            "grated Security=True;Connect Timeout=30;User Instance=True")]
        public string ICEDatabaseConnectionString {
            get {
                return ((string)(this["ICEDatabaseConnectionString"]));
            }
        }
    }
}
