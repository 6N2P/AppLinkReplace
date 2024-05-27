using System;
using System.Windows;

namespace AppLinkReplace
{
    /// <summary>
    /// Interaction logic for WindowProgress.xaml
    /// </summary>
    public partial class WindowProgress
    {
        #region events

        public delegate void OnStopVoidHandler(WindowProgress sender);

        public event OnStopVoidHandler OnStopEvent;

        #endregion events

        public WindowProgress()
        {
            //SourceInitialized += WindowSourceInitialized;
            InitializeComponent();
        }

        public void SetProgress(string text, double progress)
        {
            if (text == null || text_status == null) return;
            text_status.Text = text;
            if (!(progress > 0) || progress_bar == null) return;
            progress_bar.Value = progress;
        }

        public Double ProgressMaximum
        {
            set
            {
                progress_bar.Maximum = value;
            }
        }

        private void BtStopClick(object sender, RoutedEventArgs e)
        {
            if (OnStopEvent != null)
            {
                OnStopEvent(this);
            }
        }

        //private void WindowSourceInitialized(object sender, EventArgs e)
        //{
        //    //var wih = new WindowInteropHelper(this);
        //    //var style = GetWindowLong(wih.Handle, GWL_STYLE);
        //    //SetWindowLong(wih.Handle, GWL_STYLE, style & ~WS_SYSMENU);
        //}

        #region win api

        //private const int GWL_STYLE = -16;
        //private const int WS_SYSMENU = 0x00080000;

        //[DllImport("user32.dll")]
        //private static extern int SetWindowLong(IntPtr hwnd, int index, int value);

        //[DllImport("user32.dll")]
        //private static extern int GetWindowLong(IntPtr hwnd, int index);

        #endregion win api
    }
}