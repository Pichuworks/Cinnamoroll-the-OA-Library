using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AddInDesignerObjects;
using Office;
using Word;

namespace Cinnamoroll_the_WPS_OA_Library
{
    public class CinnamorollMain : IDTExtensibility2, IRibbonExtensibility
    {
        public static Word.Application app = null;
        public static object wps;
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            wps = Application;
            app = wps as Word.Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            return;
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            return;
        }

        public void OnStartupComplete(ref Array custom)
        {
            return;
        }

        public void OnBeginShutdown(ref Array custom)
        {
            return;
        }

        public string GetCustomUI(string RibbonID)
        {
            return Properties.Resource1.Ribbon;
        }

        public Bitmap GetRibbonImage(IRibbonControl ctrl)
        {
            switch (ctrl.Id)
            {
                case "BtnAddItems":
                    return Properties.Resource1.PNGAddItems;
                case "BtnDeleteItems":
                    return Properties.Resource1.PNGDeleteItems;
                case "BtnDocCommDemo":
                    return Properties.Resource1.PNGDocComm;
            }
            return null;
        }

        public void AddComment(IRibbonControl ctrl)
        {
            MessageBox.Show("Hello World");
            return;
        }

        public void AddAddinField(IRibbonControl ctrl)
        {
            return;
        }

        public void AddMarking(IRibbonControl ctrl)
        {
            return;
        }

        public void DeleteComment(IRibbonControl ctrl)
        {
            return;
        }

        public void DeleteAddinField(IRibbonControl ctrl)
        {
            return;
        }

        public void DeleteMarking(IRibbonControl ctrl)
        {
            return;
        }

        public void DocCommDemo(IRibbonControl ctrl)
        {
            return;
        }
    }
}
