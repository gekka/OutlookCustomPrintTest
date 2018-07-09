

namespace OutlookCustomPrintTest
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Threading;
    using System.Windows.Forms;
    using System.Reflection;

    [Flags]
    public enum DLCTL : int
    {
        DLIMAGES = 0x00000010,
        VIDEOS = 0x00000020,
        BGSOUNDS = 0x00000040,
        NO_SCRIPTS = 0x00000080,
        NO_JAVA = 0x00000100,
        NO_RUNACTIVEXCTLS = 0x00000200,
        NO_DLACTIVEXCTLS = 0x00000400,
        DOWNLOADONLY = 0x00000800,
        NO_FRAMEDOWNLOAD = 0x00001000,
        RESYNCHRONIZE = 0x00002000,
        PRAGMA_NO_CACHE = 0x00004000,
        NO_BEHAVIORS = 0x00008000,
        NO_METACHARSET = 0x00010000,
        URL_ENCODING_DISABLE_UTF8 = 0x00020000,
        URL_ENCODING_ENABLE_UTF8 = 0x00040000,
        NOFRAMES = 0x00080000,
        FORCEOFFLINE = 0x10000000,
        NO_CLIENTPULL = 0x20000000,
        SILENT = 0x40000000,
        OFFLINEIFNOTCONNECTED = unchecked((int)0x80000000),
        OFFLINE = OFFLINEIFNOTCONNECTED,
    }


    public class CustomWebBrowser : WebBrowser
    {
        private const int DISPID_AMBIENT_DLCONTROL = -5512;
        private DLCTL _downloadControlFlags = DLCTL. OFFLINE;

        protected override WebBrowserSiteBase CreateWebBrowserSiteBase()
        {
            return new CustomWebBrowserSite(this);
        }

        protected override void AttachInterfaces(object nativeActiveXObject)
        {
            base.AttachInterfaces(nativeActiveXObject);
            IOleControl ctl = (IOleControl)ActiveXInstance;
            ctl.OnAmbientPropertyChange(DISPID_AMBIENT_DLCONTROL);
        }

        [DispId(DISPID_AMBIENT_DLCONTROL)]
        public int DownloadControlFlags
        {
            get { return (int)_downloadControlFlags; }
            set
            {
                if ((int)_downloadControlFlags != value)
                {
                    _downloadControlFlags = (DLCTL)value;
                    IOleControl ctl = (IOleControl)ActiveXInstance;
                    ctl.OnAmbientPropertyChange(DISPID_AMBIENT_DLCONTROL);
                }
            }
        }

        protected class CustomWebBrowserSite : WebBrowserSite, IReflect
        {
            private Dictionary<int, PropertyInfo> _dispidCache;
            private CustomWebBrowser _host;

            public CustomWebBrowserSite(CustomWebBrowser host)
                : base(host)
            {
                _host = host;

                _dispidCache = new Dictionary<int, PropertyInfo>();
                foreach (PropertyInfo pi in _host.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                {
                    if ((!pi.CanRead) || (pi.GetIndexParameters().Length > 0))
                    {
                        continue;
                    }
                    object[] atts = pi.GetCustomAttributes(typeof(DispIdAttribute), true);
                    if ((atts != null) && (atts.Length > 0))
                    {
                        DispIdAttribute da = (DispIdAttribute)atts[0];
                        _dispidCache[da.Value] = pi;
                    }
                }
            }

            object IReflect.InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, System.Globalization.CultureInfo culture, string[] namedParameters)
            {
                object ret = null;
                const string dispidToken = "[DISPID=";
                if (name.StartsWith(dispidToken) && name.EndsWith("]"))
                {
                    int dispid = int.Parse(name.Substring(dispidToken.Length, name.Length - dispidToken.Length - 1));
                    PropertyInfo property;
                    if (_dispidCache.TryGetValue(dispid, out property))
                    {
                        ret = property.GetValue(_host, null);
                    }
                }
                return ret;
            }

            FieldInfo[] IReflect.GetFields(BindingFlags bindingAttr) { return GetType().GetFields(bindingAttr); }
            MethodInfo[] IReflect.GetMethods(BindingFlags bindingAttr) { return GetType().GetMethods(bindingAttr); }
            PropertyInfo[] IReflect.GetProperties(BindingFlags bindingAttr) { return GetType().GetProperties(bindingAttr); }

            FieldInfo IReflect.GetField(string name, BindingFlags bindingAttr) { throw new NotImplementedException(); }
            MemberInfo[] IReflect.GetMember(string name, BindingFlags bindingAttr) { throw new NotImplementedException(); }
            MemberInfo[] IReflect.GetMembers(BindingFlags bindingAttr) { throw new NotImplementedException(); }
            MethodInfo IReflect.GetMethod(string name, BindingFlags bindingAttr) { throw new NotImplementedException(); }
            MethodInfo IReflect.GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers) { throw new NotImplementedException(); }
            PropertyInfo IReflect.GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers) { throw new NotImplementedException(); }
            PropertyInfo IReflect.GetProperty(string name, BindingFlags bindingAttr) { throw new NotImplementedException(); }
            Type IReflect.UnderlyingSystemType { get { throw new NotImplementedException(); } }
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("B196B288-BAB4-101A-B69C-00AA00341D07")]
        internal interface IOleControl
        {
            void Reserved0();
            void Reserved1();
            void OnAmbientPropertyChange(int dispID);
            void Reserved2();
        }
    }

}