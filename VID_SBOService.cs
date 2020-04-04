using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using SAPbobsCOM;

namespace Llacolen_SBOService
{
    public partial class VID_SBOService : ServiceBase
    {
        private Timer _timer = new System.Timers.Timer();
        private SBOControl SBOCtrl;
        private Boolean FirstTime;
        public Logs.Logger oLog;

        public VID_SBOService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            SBOCtrl = new SBOControl();
            oLog = new Logs.Logger();
            FirstTime = true;
            EventLog.WriteEntry("Servicio iniciado.");
            oLog.LogMsg("Servicio iniciado", "A", "I");


            SBOCtrl.oLog = oLog;
            _timer.Interval = 20*1000; // 20 segundos 
            _timer.AutoReset = true;
            _timer.Elapsed += OnElapsedEvent;
            _timer.Start();

            oLog.LogMsg("Timer enabled", "F", "I");
        }

        protected override void OnStop()
        {
            _timer.Stop();
            EventLog.WriteEntry("Servicio detenido.");
            oLog.LogMsg("Servicio detenido - timer off", "A", "I");
        }

        private void OnElapsedEvent(object sender, ElapsedEventArgs e)
        {
            int nError = 0;
            string sMsg = "";

            // Write an entry to the Application log in the Event Viewer.
            //oLog.LogMsg("The service timer's Elapsed event was triggered.", "A", "D");

            _timer.Stop();
            oLog.LogMsg("Timer paused", "F", "D");

            if (FirstTime)
            {
                FirstTime = false;
                _timer.Interval = Llacolen_SBOService.Properties.Settings.Default.IntervaloEnSegundos * 1000;
            }

            SBOCtrl.Doit(ref nError, ref sMsg);
            
            _timer.Start();
            oLog.LogMsg("Timer restart", "F", "D");
        }
    }
}
